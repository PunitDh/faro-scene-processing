import argparse
import subprocess
from pathlib import Path
import sys


def find_fls(root: Path) -> list[Path]:
    """Recursively find real .fls files (not folders named .fls)."""
    return sorted(
        p for p in root.rglob("*.fls")
        if p.is_file()
    )


def write_manifest(paths: list[Path], manifest_path: Path) -> None:
    """Write a manifest file containing one .fls path per line."""
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    with manifest_path.open("w", encoding="utf-8") as f:
        for p in paths:
            f.write(str(p) + "\n")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Batch export FARO SCENE panoramas by driving SCENE with AutoHotkey."
    )
    parser.add_argument("input_dir", type=Path, help="Root folder to search recursively for *.fls")
    parser.add_argument("output_dir", type=Path, help="Folder where panoramas should be exported")
    parser.add_argument("--ahk", type=Path, default=Path("scene_export.ahk"),
                        help="Path to the AutoHotkey script")
    parser.add_argument("--manifest", type=Path, default=Path("work/manifest.txt"),
                        help="Where to write the list of *.fls files")
    parser.add_argument("--autohotkey-exe", type=Path, default=None,
                        help="Optional full path to AutoHotkey.exe. If omitted, uses AutoHotkey.exe from PATH.")
    args = parser.parse_args()

    input_dir: Path = args.input_dir.resolve()
    output_dir: Path = args.output_dir.resolve()
    ahk_script: Path = args.ahk.resolve()
    manifest: Path = args.manifest.resolve()

    # Basic validation
    if not input_dir.exists():
        print(f"Input dir not found: {input_dir}", file=sys.stderr)
        return 2
    if not ahk_script.exists():
        print(f"AHK script not found: {ahk_script}", file=sys.stderr)
        return 2

    # Find all .fls
    fls_files = find_fls(input_dir)
    if not fls_files:
        print("No .fls files found.")
        return 0

    # Ensure output folder exists
    output_dir.mkdir(parents=True, exist_ok=True)

    # Create manifest file
    write_manifest(fls_files, manifest)

    # Decide which AutoHotkey executable to use
    ahk_exe = str(args.autohotkey_exe) if args.autohotkey_exe else "AutoHotkey.exe"

    print(f"Found {len(fls_files)} .fls files.")
    print(f"Manifest written to: {manifest}")
    print(f"Output folder:       {output_dir}")
    print(f"Running AHK:         {ahk_exe}")
    print(f"Script:              {ahk_script}")

    # Launch AutoHotkey and pass (manifest, output_dir) as args
    result = subprocess.run([ahk_exe, str(ahk_script), str(manifest), str(output_dir)])

    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())