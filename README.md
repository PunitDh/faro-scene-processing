# Faro Scene Processing

Internal tool to batch export panoramas from FARO SCENE `.fls` files with automated UI control via AutoHotkey.

## Prerequisites

- **AutoHotkey v2** – Download from [autohotkey.com](https://www.autohotkey.com/). Typically installs to `C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe`
- **FARO SCENE** – With a default project open, switched to the **Import tab**, and **no imports currently in the Import section**
- **Python 3.10+**

## Setup

### 1. Create & activate virtual environment
```pwsh
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### 2. Install dependencies
```pwsh
pip install -r requirements.txt
```

## Input File Format

Create input/input.json with job specifications:
```json
[
  {
    "SourceFilePath": "C:\\path\\to\\scan1.fls",
    "SiteName": "Site A",
    "AssetName": "Building 1",
    "InspectionDate": "2026-03-06"
  },
  {
    "SourceFilePath": "C:\\path\\to\\scan2.fls",
    "SiteName": "Site B",
    "AssetName": "Bridge",
    "InspectionDate": "2026-03-05"
  }
]
```

## Running

### Basic command
```pwsh
python .\scene_json_driver.py .\input\input.json .\out\ --autohotkey-exe "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
```

### With options
```pwsh
python .\scene_json_driver.py .\input\input.json .\out\ 
  --autohotkey-exe "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe" 
  --batch-size 10 
  --work-dir .\work\ 
  --resume-log .\log\processed_targets.txt 
  --csv-log .\log\run_log.csv
```

**Options:**
- `--batch-size` – Number of files per batch (default: 10)
- `--work-dir` – Temp staging directory (default: work/)
- `--resume-log` – Track completed jobs to skip re-processing
- `--csv-log` – Detailed run log with timestamps and errors
- `--require-png` – Fail batch if PNG export fails (default: optional)

## Cleanup

### Clean working folders

To clean up the `log/`, `out/`, and `work/` directories while preserving `.keep` files:

```pwsh
python .\clear_output.py
```

For a dry run (to see what would be deleted without making changes):

```pwsh
python .\clear_output.py --dry-run
```

## Output Structure

Exports organized by inspection hierarchy:
```
out/
├── SiteName/
│   └── AssetName/
│       └── InspectionDate/
│           ├── folder_name.jpg
│           └── folder_name.png
```

## Critical Notes

**During Execution:**
- Do **NOT** interact with FARO SCENE or any other windows—the script controls the UI automatically
- Press `Ctrl+Q` in FARO SCENE window to abort the AutoHotkey script if needed
- Keep FARO SCENE and terminal visible and focused as needed

**System Requirements:**
- **Disable PC sleep/hibernation** – Script can take hours; PC must not sleep
- **Disable screensaver** – Any screensaver interruption will break UI automation
- **Disable Windows Update notifications/auto-restart** – Can pause/restart mid-batch
- **Close unnecessary applications** – Only run FARO SCENE + terminal (IDE optional if code edits needed)
- **Disable mouse tracking/accessibility features** that might interfere with automation

**Troubleshooting:**
- Logs and manifests are written to `log/` and `work/` directories for debugging
- Use `--resume-log` to skip already-processed files if a batch fails mid-way
- Check CSV log (`log/run_log.csv`) for per-file status and error messages
