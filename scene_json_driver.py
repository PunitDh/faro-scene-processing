import argparse
import json
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class Job:
    source_fls: Path
    site: str
    asset: str
    inspection_date: str

    @property
    def source_folder_name(self) -> str:
        # Folder containing the .fls file
        return self.source_fls.parent.name

    @property
    def fls_stem(self) -> str:
        return self.source_fls.stem

    def target_relpath(self) -> Path:
        # {SiteName}/{AssetName}/{InspectionDate}/{SourceFolderNameAsFileName}.jpg
        # NOTE: This can collide if multiple fls are in the same folder for same site/asset/date.
        # We handle collisions by suffixing _2, _3, ...
        filename = f"{self.source_folder_name}.jpg"
        return Path(self.site) / self.asset / self.inspection_date / filename


def safe_name(s: str) -> str:
    # Keep folder names Windows-friendly (you can loosen if you want)
    s = s.strip()
    # Replace reserved characters: \ / : * ? " < > |
    return re.sub(r'[\\/:*?"<>|]+', "_", s)


def load_jobs(json_path: Path) -> list[Job]:
    data = json.loads(json_path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("Input JSON must be a list of objects.")

    jobs: list[Job] = []
    for i, obj in enumerate(data):
        if not isinstance(obj, dict):
            raise ValueError(f"Item {i} is not an object.")

        try:
            src = Path(obj["SourceFilePath"])
            site = str(obj["SiteName"])
            asset = str(obj["AssetName"])
            date = str(obj["InspectionDate"])
        except KeyError as e:
            raise ValueError(f"Item {i} missing key: {e}") from e

        jobs.append(
            Job(
                source_fls=src,
                site=safe_name(site),
                asset=safe_name(asset),
                inspection_date=safe_name(date),
            )
        )

    return jobs


def write_manifest(fls_paths: list[Path], manifest_path: Path) -> None:
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    with manifest_path.open("w", encoding="utf-8") as f:
        for p in fls_paths:
            f.write(str(p) + "\n")


def list_jpgs(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    out = []
    out.extend(folder.rglob("*.jpg"))
    out.extend(folder.rglob("*.jpeg"))
    return sorted(out)


def pick_unique_target(target: Path) -> Path:
    if not target.exists():
        return target
    stem = target.stem
    suffix = target.suffix
    parent = target.parent
    k = 2
    while True:
        cand = parent / f"{stem}_{k}{suffix}"
        if not cand.exists():
            return cand
        k += 1


def match_exported_jpg(job: Job, exported: list[Path]) -> Path | None:
    """
    Heuristic matcher:
      - Prefer JPG filenames containing the .fls stem (case-insensitive)
      - If multiple matches, pick shortest name (usually most direct)
    """
    needle = job.fls_stem.lower()
    matches = [p for p in exported if needle in p.stem.lower()]
    if not matches:
        return None
    matches.sort(key=lambda p: (len(p.name), p.name.lower()))
    return matches[0]


def append_csv(log_path: Path, row: list[str]) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    # Minimal CSV escaping
    def esc(s: str) -> str:
        s = s.replace('"', '""')
        return f'"{s}"'
    line = ",".join(esc(x) for x in row) + "\n"
    log_path.write_text("", encoding="utf-8") if not log_path.exists() else None
    with log_path.open("a", encoding="utf-8") as f:
        f.write(line)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Drive FARO SCENE export using AHK with JSON input + structured output folders."
    )
    ap.add_argument("input_json", type=Path, help="JSON file like [{SourceFilePath,SiteName,AssetName,InspectionDate}, ...]")
    ap.add_argument("output_dir", type=Path, help="Root output folder")
    ap.add_argument("--ahk", type=Path, default=Path("scene_export.ahk"), help="Your working AHK script")
    ap.add_argument("--autohotkey-exe", type=Path, default=None, help="Path to AutoHotkey.exe (optional)")
    ap.add_argument("--batch-size", type=int, default=10)
    ap.add_argument("--work-dir", type=Path, default=Path("work_scene"))
    ap.add_argument("--resume-log", type=Path, default=Path("work_scene/processed_targets.txt"))
    ap.add_argument("--csv-log", type=Path, default=Path("work_scene/run_log.csv"))
    args = ap.parse_args()

    input_json: Path = args.input_json.resolve()
    out_root: Path = args.output_dir.resolve()
    ahk_script: Path = args.ahk.resolve()
    work_dir: Path = args.work_dir.resolve()
    resume_log: Path = args.resume_log.resolve()
    csv_log: Path = args.csv_log.resolve()

    if not input_json.exists():
        print(f"JSON not found: {input_json}", file=sys.stderr)
        return 2
    if not ahk_script.exists():
        print(f"AHK script not found: {ahk_script}", file=sys.stderr)
        return 2

    out_root.mkdir(parents=True, exist_ok=True)
    work_dir.mkdir(parents=True, exist_ok=True)

    jobs_all = load_jobs(input_json)

    # Resume set is "target file path already exists OR is listed"
    done_targets: set[str] = set()
    if resume_log.exists():
        for line in resume_log.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line:
                done_targets.add(line)

    # Filter jobs that already have their target jpg on disk OR are logged done
    jobs: list[Job] = []
    for j in jobs_all:
        target = out_root / j.target_relpath()
        if str(target) in done_targets:
            continue
        if target.exists():
            done_targets.add(str(target))
            continue
        jobs.append(j)

    if not jobs:
        print("Nothing to do (everything already processed).")
        return 0

    ahk_exe = str(args.autohotkey_exe) if args.autohotkey_exe else "AutoHotkey.exe"

    total = len(jobs)
    batch_size = max(1, int(args.batch_size))
    batch_num = 0
    idx = 0

    # CSV header if new
    if not csv_log.exists():
        append_csv(csv_log, ["event", "batch", "index", "total", "source_fls", "staging", "target", "note"])

    while idx < total:
        batch_num += 1
        batch_jobs = jobs[idx: idx + batch_size]
        idx_end = idx + len(batch_jobs)

        staging = work_dir / f"staging_batch_{batch_num:04d}"
        if staging.exists():
            shutil.rmtree(staging)
        staging.mkdir(parents=True, exist_ok=True)

        manifest = work_dir / f"manifest_batch_{batch_num:04d}.txt"
        write_manifest([j.source_fls for j in batch_jobs], manifest)

        append_csv(csv_log, ["BATCH_START", str(batch_num), f"{idx+1}", str(total), "-", str(staging), "-", f"count={len(batch_jobs)}"])

        # Run AHK: (manifest, staging)
        proc = subprocess.run([ahk_exe, str(ahk_script), str(manifest), str(staging)])
        if proc.returncode != 0:
            append_csv(csv_log, ["ERROR_AHK", str(batch_num), f"{idx+1}", str(total), "-", str(staging), "-", f"returncode={proc.returncode}"])
            print(f"AHK failed on batch {batch_num} (returncode={proc.returncode}).", file=sys.stderr)
            return proc.returncode or 1

        # After AHK returns, exported JPGs should be in staging
        exported = list_jpgs(staging)
        append_csv(csv_log, ["BATCH_EXPORTED", str(batch_num), f"{idx+1}", str(total), "-", str(staging), "-", f"jpgs_found={len(exported)}"])

        # Move + rename for each job; mark processed ONLY after the target exists
        used: set[Path] = set()
        for j in batch_jobs:
            target = out_root / j.target_relpath()
            target.parent.mkdir(parents=True, exist_ok=True)

            match = match_exported_jpg(j, [p for p in exported if p not in used])
            if match is None:
                append_csv(csv_log, ["ERROR_NO_MATCH", str(batch_num), "-", str(total), str(j.source_fls), str(staging), str(target), "no jpg matched fls stem"])
                print(f"No JPG matched {j.source_fls.name} in staging batch {batch_num}. Stopping.", file=sys.stderr)
                return 1

            used.add(match)

            # Handle collisions safely
            final_target = pick_unique_target(target)

            # Move file
            final_target.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(match), str(final_target))

            # Verify it exists (your directive)
            if not final_target.exists():
                append_csv(csv_log, ["ERROR_MOVE_FAILED", str(batch_num), "-", str(total), str(j.source_fls), str(staging), str(final_target), "move did not result in file on disk"])
                print(f"Move failed for {match} -> {final_target}", file=sys.stderr)
                return 1

            # Mark processed
            with resume_log.open("a", encoding="utf-8") as f:
                f.write(str(final_target) + "\n")
            done_targets.add(str(final_target))

            append_csv(csv_log, ["OK", str(batch_num), "-", str(total), str(j.source_fls), str(staging), str(final_target), ""])

        # Optional: If staging still has JPGs we didn't map, log it (doesn't fail)
        leftovers = [p for p in list_jpgs(staging)]
        if leftovers:
            append_csv(csv_log, ["WARN_LEFTOVERS", str(batch_num), "-", str(total), "-", str(staging), "-", f"leftover_jpgs={len(leftovers)}"])

        # Clean staging (optional)
        shutil.rmtree(staging, ignore_errors=True)

        append_csv(csv_log, ["BATCH_DONE", str(batch_num), f"{idx_end}", str(total), "-", "-", "-", ""])
        idx = idx_end

    append_csv(csv_log, ["DONE", "-", str(total), str(total), "-", "-", "-", ""])
    print(f"Done. Processed {total} items.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())