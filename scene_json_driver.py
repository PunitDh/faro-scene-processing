import argparse
import json
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


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

    def target_relpath(self, ext: str) -> Path:
        """
        Final output path:
          {SiteName}/{AssetName}/{InspectionDate}/{SourceFolderName}.{ext}

        Notes:
          - Same basename for both JPG + PNG so they sit side-by-side.
          - Collision-handled via pick_unique_target() at move time.
        """
        ext = ext.lstrip(".").lower()
        filename = f"{self.source_folder_name}.{ext}"
        return Path(self.site) / self.asset / self.inspection_date / filename


def safe_name(s: str) -> str:
    s = s.strip()
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


def list_files_by_ext(folder: Path, exts: tuple[str, ...]) -> list[Path]:
    if not folder.exists():
        return []
    out: list[Path] = []
    for ext in exts:
        out.extend(folder.rglob(f"*.{ext.lstrip('.')}"))
    return sorted(out)


def list_jpgs(folder: Path) -> list[Path]:
    return list_files_by_ext(folder, ("jpg", "jpeg"))


def list_pngs(folder: Path) -> list[Path]:
    return list_files_by_ext(folder, ("png",))


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


def match_exported_file(job: Job, exported: list[Path]) -> Path | None:
    """
    Heuristic matcher:
      1) Prefer filenames containing job.fls_stem (case-insensitive)
      2) Fallback: filenames containing source_folder_name (case-insensitive)
      3) If multiple, pick shortest name then lexicographic
    """
    if not exported:
        return None

    needle1 = job.fls_stem.lower()
    needle2 = job.source_folder_name.lower()

    matches = [p for p in exported if needle1 in p.stem.lower()]
    if not matches:
        matches = [p for p in exported if needle2 in p.stem.lower()]
    if not matches:
        return None

    matches.sort(key=lambda p: (len(p.name), p.name.lower()))
    return matches[0]


def append_csv(log_path: Path, row: list[str]) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)

    def esc(s: str) -> str:
        s = s.replace('"', '""')
        return f'"{s}"'

    if not log_path.exists():
        log_path.write_text("", encoding="utf-8")

    line = ",".join(esc(x) for x in row) + "\n"
    with log_path.open("a", encoding="utf-8") as f:
        f.write(line)


def run_ahk(ahk_exe: str, ahk_script: Path, manifest: Path, staging: Path) -> int:
    # NOTE: AHK gets (manifest, output_folder). We pass staging as the output folder.
    proc = subprocess.run([ahk_exe, str(ahk_script), str(manifest), str(staging)])
    return proc.returncode


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Drive FARO SCENE export using AHK with JSON input + structured output folders."
    )
    ap.add_argument(
        "input_json",
        type=Path,
        help="JSON file like [{SourceFilePath,SiteName,AssetName,InspectionDate}, ...]",
    )
    ap.add_argument("output_dir", type=Path, help="Root output folder")
    ap.add_argument("--ahk", type=Path, default=Path("scene_export.ahk"), help="Your AHK script")
    ap.add_argument("--autohotkey-exe", type=Path, default=None, help="Path to AutoHotkey.exe (optional)")
    ap.add_argument("--batch-size", type=int, default=10)
    ap.add_argument("--work-dir", type=Path, default=Path("work"))
    ap.add_argument("--resume-log", type=Path, default=Path("log/processed_targets.txt"))
    ap.add_argument("--csv-log", type=Path, default=Path("log/run_log.csv"))
    ap.add_argument(
        "--require-png",
        action="store_true",
        help="Fail the batch if a PNG cannot be found for a job (default: PNG is optional).",
    )
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

    # Filter jobs that already have their target JPG on disk OR are logged done.
    # (JPG is treated as the primary “done” artifact; PNG can be optional.)
    jobs: list[Job] = []
    for j in jobs_all:
        target_jpg = out_root / j.target_relpath("jpg")
        if str(target_jpg) in done_targets:
            continue
        if target_jpg.exists():
            done_targets.add(str(target_jpg))
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

    if not csv_log.exists():
        append_csv(csv_log, ["event", "batch", "index", "total", "source_fls", "staging", "target", "note"])

    while idx < total:
        batch_num += 1
        batch_jobs = jobs[idx : idx + batch_size]
        idx_end = idx + len(batch_jobs)

        staging = work_dir / f"staging_batch_{batch_num:04d}"
        if staging.exists():
            shutil.rmtree(staging)
        staging.mkdir(parents=True, exist_ok=True)

        manifest = work_dir / f"manifest_batch_{batch_num:04d}.txt"
        write_manifest([j.source_fls for j in batch_jobs], manifest)

        append_csv(csv_log, ["BATCH_START", str(batch_num), f"{idx+1}", str(total), "-", str(staging), "-", f"count={len(batch_jobs)}"])

        rc = run_ahk(ahk_exe, ahk_script, manifest, staging)
        if rc != 0:
            append_csv(csv_log, ["ERROR_AHK", str(batch_num), f"{idx+1}", str(total), "-", str(staging), "-", f"returncode={rc}"])
            print(f"AHK failed on batch {batch_num} (returncode={rc}).", file=sys.stderr)
            return rc or 1

        # AHK might export into:
        # - staging root, OR
        # - staging/ScanResolution + staging/FullColor (recommended),
        # so we search in BOTH the whole staging tree and the specific subfolders if present.
        scan_dir = staging / "ScanResolution"
        fc_dir = staging / "FullColor"

        # Collect JPGs
        exported_jpgs = []
        if scan_dir.exists():
            exported_jpgs = list_jpgs(scan_dir)
        if not exported_jpgs:
            exported_jpgs = list_jpgs(staging)

        # Collect PNGs
        exported_pngs = []
        if fc_dir.exists():
            exported_pngs = list_pngs(fc_dir)
        if not exported_pngs:
            exported_pngs = list_pngs(staging)

        append_csv(csv_log, ["BATCH_EXPORTED", str(batch_num), f"{idx+1}", str(total), "-", str(staging), "-", f"jpgs_found={len(exported_jpgs)} pngs_found={len(exported_pngs)}"])

        used_jpg: set[Path] = set()
        used_png: set[Path] = set()

        for j in batch_jobs:
            # ---- Move JPG (required) ----
            target_jpg = out_root / j.target_relpath("jpg")
            target_jpg.parent.mkdir(parents=True, exist_ok=True)

            jpg_match = match_exported_file(j, [p for p in exported_jpgs if p not in used_jpg])
            if jpg_match is None:
                append_csv(csv_log, ["ERROR_NO_MATCH_JPG", str(batch_num), "-", str(total), str(j.source_fls), str(staging), str(target_jpg), "no jpg matched fls stem/folder"])
                print(f"No JPG matched {j.source_fls.name} in staging batch {batch_num}. Stopping.", file=sys.stderr)
                return 1

            used_jpg.add(jpg_match)
            final_jpg = pick_unique_target(target_jpg)
            shutil.move(str(jpg_match), str(final_jpg))

            if not final_jpg.exists():
                append_csv(csv_log, ["ERROR_MOVE_FAILED_JPG", str(batch_num), "-", str(total), str(j.source_fls), str(staging), str(final_jpg), "move did not result in file on disk"])
                print(f"Move failed for {jpg_match} -> {final_jpg}", file=sys.stderr)
                return 1

            # Mark processed (JPG)
            with resume_log.open("a", encoding="utf-8") as f:
                f.write(str(final_jpg) + "\n")
            done_targets.add(str(final_jpg))

            # ---- Move PNG (optional unless --require-png) ----
            target_png = out_root / j.target_relpath("png")
            target_png.parent.mkdir(parents=True, exist_ok=True)

            png_match = match_exported_file(j, [p for p in exported_pngs if p not in used_png])
            if png_match is None:
                note = "no png matched fls stem/folder"
                if args.require_png:
                    append_csv(csv_log, ["ERROR_NO_MATCH_PNG", str(batch_num), "-", str(total), str(j.source_fls), str(staging), str(target_png), note])
                    print(f"No PNG matched {j.source_fls.name} in staging batch {batch_num}. Stopping (require-png).", file=sys.stderr)
                    return 1
                else:
                    append_csv(csv_log, ["WARN_NO_PNG", str(batch_num), "-", str(total), str(j.source_fls), str(staging), str(target_png), note])
            else:
                used_png.add(png_match)
                final_png = pick_unique_target(target_png)
                shutil.move(str(png_match), str(final_png))

                if not final_png.exists():
                    append_csv(csv_log, ["ERROR_MOVE_FAILED_PNG", str(batch_num), "-", str(total), str(j.source_fls), str(staging), str(final_png), "move did not result in file on disk"])
                    print(f"Move failed for {png_match} -> {final_png}", file=sys.stderr)
                    return 1

                # Log PNG as processed too (handy for auditing)
                with resume_log.open("a", encoding="utf-8") as f:
                    f.write(str(final_png) + "\n")
                done_targets.add(str(final_png))

            append_csv(csv_log, ["OK", str(batch_num), "-", str(total), str(j.source_fls), str(staging), str(final_jpg), "jpg ok; png optional"])

        # Log leftovers (non-fatal)
        leftovers_j = list_jpgs(staging)
        leftovers_p = list_pngs(staging)
        if leftovers_j or leftovers_p:
            append_csv(csv_log, ["WARN_LEFTOVERS", str(batch_num), "-", str(total), "-", str(staging), "-", f"leftover_jpgs={len(leftovers_j)} leftover_pngs={len(leftovers_p)}"])

        # Clean staging
        shutil.rmtree(staging, ignore_errors=True)

        append_csv(csv_log, ["BATCH_DONE", str(batch_num), f"{idx_end}", str(total), "-", "-", "-", ""])
        idx = idx_end

    append_csv(csv_log, ["DONE", "-", str(total), str(total), "-", "-", "-", ""])
    print(f"Done. Processed {total} items.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())