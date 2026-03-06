from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path


TARGET_DIRS = ("log", "out", "work")
KEEP_FILENAME = ".keep"


def purge_dir(base: Path, dry_run: bool) -> tuple[int, int]:
    """Remove everything inside base except files named .keep."""
    deleted_files = 0
    deleted_dirs = 0

    if not base.exists() or base.is_symlink():
        return deleted_files, deleted_dirs

    for entry in base.iterdir():
        if entry.name == KEEP_FILENAME and entry.is_file():
            continue

        if dry_run:
            print(f"[DRY-RUN] DELETE {entry}")
            continue

        if entry.is_dir() and not entry.is_symlink():
            shutil.rmtree(entry)
            deleted_dirs += 1
        else:
            entry.unlink()
            deleted_files += 1

    return deleted_files, deleted_dirs


def ensure_keep_file(base: Path, dry_run: bool) -> None:
    if base.is_symlink():
        return
    keep_path = base / KEEP_FILENAME
    if keep_path.exists():
        return
    if dry_run:
        print(f"[DRY-RUN] CREATE {keep_path}")
        return
    base.mkdir(parents=True, exist_ok=True)
    keep_path.write_text("", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Clean log/out/work while preserving .keep files."
    )
    parser.add_argument(
        "--root",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="Project root containing log/out/work (default: script directory).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would be deleted without making changes.",
    )
    args = parser.parse_args()

    root = args.root.resolve()
    if not root.exists():
        print(f"Root path does not exist: {root}", file=sys.stderr)
        return 2

    total_files = 0
    total_dirs = 0

    for dirname in TARGET_DIRS:
        target = root / dirname
        if not target.exists():
            if args.dry_run:
                print(f"[DRY-RUN] SKIP missing folder {target}")
            else:
                target.mkdir(parents=True, exist_ok=True)
            ensure_keep_file(target, args.dry_run)
            continue

        files, dirs = purge_dir(target, args.dry_run)
        total_files += files
        total_dirs += dirs
        ensure_keep_file(target, args.dry_run)

    if args.dry_run:
        print("Dry run complete.")
    else:
        print(
            f"Cleanup complete. Deleted files: {total_files}, deleted directories: {total_dirs}."
        )

    return 0


if __name__ == "__main__":
    sys.exit(main())
