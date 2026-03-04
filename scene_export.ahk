^q:: {
    ToolTip "Script aborted."
    Sleep 500
    ExitApp
}

#SingleInstance Force
SetTitleMatchMode 2
SendMode "Input"

; ---------- Args ----------
if A_Args.Length < 2 {
    MsgBox "Usage: scene_export.ahk <manifest.txt> <output_folder>"
    ExitApp 2
}

manifest := A_Args[1]
outDir   := A_Args[2]

if !FileExist(manifest) {
    MsgBox "Manifest not found: " manifest
    ExitApp 2
}
if !DirExist(outDir)
    DirCreate outDir

; =========================
; TUNE THESE
; =========================
sceneTitle := "SCENE 2026.0.1"
BATCH_SIZE := 10

waitImportMs := 2000
importResultsOkDelayMs := 1500

; --- SCENE window coords ---
IMPORT_BTN_X := 60
IMPORT_BTN_Y := 120

TREE_X := 80
TREE_Y := 235

SCANS_ROOT_X := 77
SCANS_ROOT_Y := 232

; --- screen coords for SCENE custom import OK ---
IMPORT_RESULTS_OK_X := 1250
IMPORT_RESULTS_OK_Y := 690

exportDlgTitle := "Select folder for images export"

useBlockInput := true

; Export dialog (WINDOW RELATIVE to export dialog)
EXPORT_FOLDER_FIELD_WX := 460
EXPORT_FOLDER_FIELD_WY := 512
EXPORT_SELECT_BTN_WX   := 715
EXPORT_SELECT_BTN_WY   := 550

; WAIT logic (existing)
EXPORT_IDLE_MS := 25000         ; must be QUIET for 25s to consider export finished
EXPORT_MAX_SEC := 7200          ; 2 hours max safety timeout

; --- NEW: logging / resume ---
logCsvPath := A_ScriptDir "\log\scene_export_log.csv"
donePath   := A_ScriptDir "\log\processed.txt"

; --- NEW: JPG verification (on EXPORT only) ---
; We only mark processed if jpg count increases by at least batchCount.
JPG_WAIT_MAX_SEC := 3600        ; 1 hour max wait for jpg growth after export completes
JPG_IDLE_MS      := 8000        ; if jpg count doesn't change for 8s, we treat as stalled (fail)

; =========================

; ---------- Logging helpers ----------
NowStamp() {
    return FormatTime(, "yyyy-MM-dd HH:mm:ss")
}

CsvEscape(s) {
    ; CSV rule: double any internal quotes, then wrap whole field in quotes
    s := StrReplace(s, '"', '""')
    return '"' s '"'
}

LogEvent(eventType, batchNum := "-", idx := "-", total := "-", flsPath := "-") {
    global logCsvPath
    line := NowStamp() "," CsvEscape(eventType) "," batchNum "," idx "," total "," CsvEscape(flsPath) "`n"
    FileAppend line, logCsvPath, "UTF-8"
}

LoadDoneSet(path) {
    done := Map()
    if !FileExist(path)
        return done

    raw := FileRead(path, "UTF-8")
    raw := StrReplace(raw, "`r", "")
    lines := StrSplit(raw, "`n")
    for _, line in lines {
        p := Trim(line)
        if (p = "")
            continue
        done[p] := true
    }
    return done
}

AppendDoneList(path, flsArray) {
    out := ""
    for _, p in flsArray
        out .= p "`n"
    FileAppend out, path, "UTF-8"
}

; ---------- JPG verification helpers (EXPORT only) ----------
CountJpgs(dir) {
    c := 0
    Loop Files dir "\*.jpg", "FR"
        c++
    Loop Files dir "\*.jpeg", "FR"
        c++
    return c
}

WaitForJpgGrowth(dir, expectedIncrease, baseCount, maxSec := 3600, idleMs := 8000) {
    target := baseCount + expectedIncrease
    start := A_TickCount
    lastChange := A_TickCount
    lastCount := baseCount

    Loop {
        if (A_TickCount - start) > (maxSec * 1000)
            return false

        cur := CountJpgs(dir)
        if (cur >= target)
            return true

        if (cur != lastCount) {
            lastCount := cur
            lastChange := A_TickCount
        } else {
            if (A_TickCount - lastChange) > idleMs
                return false
        }
        Sleep 500
    }
}

; =========================
; EXISTING WORKING FUNCTIONS (UNCHANGED BEHAVIOR)
; =========================

ActivateScene() {
    global sceneTitle
    WinActivate sceneTitle
    WinWaitActive sceneTitle, , 10
}

ClickAtWindow(x, y, button := "Left") {
    CoordMode "Mouse", "Window"
    MouseMove x, y, 0
    Sleep 60
    Click button
    Sleep 150
}

; ---------- Import ----------
ImportOneFls(flsPath) {
    global IMPORT_BTN_X, IMPORT_BTN_Y, waitImportMs
    global IMPORT_RESULTS_OK_X, IMPORT_RESULTS_OK_Y
    global TREE_X, TREE_Y
    global importResultsOkDelayMs, useBlockInput

    ActivateScene()

    ClickAtWindow(IMPORT_BTN_X, IMPORT_BTN_Y, "Left")
    WinWaitActive "Import Scans", , 15

    ControlSetText flsPath, "Edit1", "Import Scans"
    Sleep 150
    Send "{Enter}"
    Sleep 300

    Sleep importResultsOkDelayMs

    if useBlockInput
        BlockInput true

    CoordMode "Mouse", "Screen"
    MouseMove IMPORT_RESULTS_OK_X, IMPORT_RESULTS_OK_Y, 0
    Sleep 50
    Click "Left"
    Sleep 200

    ; park focus away from ribbon
    CoordMode "Mouse", "Window"
    MouseMove TREE_X, TREE_Y, 0
    Sleep 50
    Click "Left"
    Sleep 150

    if useBlockInput
        BlockInput false

    Sleep waitImportMs
}

; ---------- Optional dialog: Confirm "File already exists" -> pick "No All" ----------
IsFileExistsDialog() {
    hwnd := WinExist("Confirm")
    if !hwnd
        return false

    Loop 10 {
        try txt := ControlGetText("Static" A_Index, "ahk_id " hwnd)
        catch {
            continue
        }
        if InStr(txt, "File already exists")
            return true
    }
    return false
}

HandleFileExistsConfirmOnce() {
    if !IsFileExistsDialog()
        return false

    WinActivate "Confirm"
    WinWaitActive "Confirm", , 2
    Sleep 120

    Send "{Tab}{Enter}"
    Sleep 200
    return true
}

IsScanModifiedDialog(hwnd) {
    Loop 10 {
        try txt := ControlGetText("Static" A_Index, "ahk_id " hwnd)
        catch {
            continue
        }
        if InStr(txt, "Scan was modified")
            return true
    }
    return false
}

; ---------- Scan modified prompt (repeat) ----------
HandleScanModifiedOnce() {
    hwnd := WinExist("ahk_class #32770 ahk_exe SCENE.exe")
    if !hwnd
        hwnd := WinExist("ahk_class #32770 ahk_exe Scene.exe")
    if !hwnd
        return false

    if !IsScanModifiedDialog(hwnd)
        return false

    WinActivate "ahk_id " hwnd
    WinWaitActive "ahk_id " hwnd, , 1
    Sleep 80

    Send "!n"
    Sleep 120

    if WinExist("ahk_id " hwnd) {
        Send "{Right}{Enter}"
        Sleep 120
    }
    return true
}

; ---------- Confirm delete scans? -> press Yes ----------
IsDeleteScansDialog() {
    hwnd := WinExist("SCENE")
    if !hwnd
        return false

    Loop 10 {
        try txt := ControlGetText("Static" A_Index, "ahk_id " hwnd)
        catch {
            continue
        }
        if InStr(txt, "delete '/Scans'")
            return true
    }
    return false
}

HandleDeleteScansConfirm() {
    if !WinWait("SCENE", , 5)
        return false
    if !IsDeleteScansDialog()
        return false

    WinActivate "SCENE"
    WinWaitActive "SCENE", , 2
    Sleep 100

    Send "!y"
    Sleep 150
    if WinExist("SCENE") {
        Send "{Enter}"
        Sleep 150
    }
    return true
}

; ---------- THIS is the critical “wait until export is done” gate ----------
DrainExportCompletion(maxSeconds := 7200, idleMs := 25000) {
    start := A_TickCount
    lastDialogSeen := A_TickCount

    Loop {
        if (A_TickCount - start) > (maxSeconds * 1000)
            return false

        handled := false

        if WinExist("Confirm") {
            HandleFileExistsConfirmOnce()
            handled := true
            lastDialogSeen := A_TickCount
        }

        if HandleScanModifiedOnce() {
            handled := true
            lastDialogSeen := A_TickCount
        }

        if !handled {
            if (A_TickCount - lastDialogSeen) > idleMs
                return true
            Sleep 200
        }
    }
}

; ---------- Export once from Scans root ----------
ExportAllFromScansRoot(outDir) {
    global SCANS_ROOT_X, SCANS_ROOT_Y
    global exportDlgTitle
    global useBlockInput
    global EXPORT_FOLDER_FIELD_WX, EXPORT_FOLDER_FIELD_WY
    global EXPORT_SELECT_BTN_WX, EXPORT_SELECT_BTN_WY
    global EXPORT_MAX_SEC, EXPORT_IDLE_MS

    ActivateScene()

    ClickAtWindow(SCANS_ROOT_X, SCANS_ROOT_Y, "Left")
    Sleep 200

    if useBlockInput
        BlockInput false

    Send "{AppsKey}"
    Sleep 150
    if !WinWait("ahk_class #32768", , 1) {
        Send "+{F10}"
        Sleep 150
    }
    if !WinWait("ahk_class #32768", , 2) {
        MsgBox "Export context menu did not open. SCANS_ROOT_X/Y likely wrong."
        return false
    }

    Send "e"
    Sleep 200
    Send "p"
    Sleep 200
    Send "s"
    Sleep 600

    if !WinWaitActive(exportDlgTitle, , 20) {
        MsgBox "Did not reach export folder picker."
        return false
    }

    WinGetPos &dlgX, &dlgY, &dlgW, &dlgH, exportDlgTitle
    fieldX := dlgX + EXPORT_FOLDER_FIELD_WX
    fieldY := dlgY + EXPORT_FOLDER_FIELD_WY
    btnX   := dlgX + EXPORT_SELECT_BTN_WX
    btnY   := dlgY + EXPORT_SELECT_BTN_WY

    if useBlockInput
        BlockInput true

    CoordMode "Mouse", "Screen"

    MouseMove fieldX, fieldY, 0
    Sleep 50
    Click "Left"
    Sleep 120

    Send "^a"
    Sleep 80
    A_Clipboard := outDir
    Sleep 80
    Send "^v"
    Sleep 200

    MouseMove btnX, btnY, 0
    Sleep 50
    Click "Left"
    Sleep 250

    if useBlockInput
        BlockInput false

    CoordMode "Mouse", "Window"

    ToolTip "Waiting for export to finish (draining dialogs)..."
    ok := DrainExportCompletion(EXPORT_MAX_SEC, EXPORT_IDLE_MS)
    ToolTip
    return ok
}

; ---------- Delete Scans root after export finished ----------
DeleteScansRoot() {
    global SCANS_ROOT_X, SCANS_ROOT_Y
    ActivateScene()

    ClickAtWindow(SCANS_ROOT_X, SCANS_ROOT_Y, "Left")
    Sleep 150

    Send "{Del}"
    Sleep 150

    HandleDeleteScansConfirm()
    Sleep 500
}

ReadManifest(path) {
    raw := FileRead(path, "UTF-8")
    raw := StrReplace(raw, "`r", "")
    lines := StrSplit(raw, "`n")

    out := []
    for _, line in lines {
        line := Trim(line)
        if (line = "")
            continue
        out.Push(line)
    }
    return out
}

; =========================
; Main (adds logging + resume + JPG verification ON EXPORT)
; =========================

flsFilesAll := ReadManifest(manifest)
totalAll := flsFilesAll.Length
if (totalAll = 0) {
    MsgBox "Manifest has 0 paths."
    ExitApp 2
}

doneSet := LoadDoneSet(donePath)

; Filter out already-processed paths (resume)
flsFiles := []
for _, p in flsFilesAll {
    if !doneSet.Has(p)
        flsFiles.Push(p)
}

total := flsFiles.Length
if (total = 0) {
    MsgBox "All scans already processed (per processed.txt). Nothing to do."
    ExitApp 0
}

LogEvent("START", "-", "-", total, "-")

idx := 1
batchNum := 0

while (idx <= total) {
    batchNum++
    batchStart := idx
    batchEnd := Min(idx + BATCH_SIZE - 1, total)
    batchCount := batchEnd - batchStart + 1

    batchPaths := []
    LogEvent("BATCH_START", batchNum, batchStart, total, "-")

    ; IMPORT batch
    Loop batchCount {
        flsPath := flsFiles[idx]
        batchPaths.Push(flsPath)

        ToolTip "BATCH " batchNum ": Importing " idx "/" total "`n" flsPath
        LogEvent("IMPORT_BEGIN", batchNum, idx, total, flsPath)

        ImportOneFls(flsPath)

        LogEvent("IMPORT_OK", batchNum, idx, total, flsPath)

        idx++
        Sleep 300
    }

    ; --- JPG baseline BEFORE export ---
    baseJpg := CountJpgs(outDir)
    LogEvent("BATCH_EXPORT_BASE_JPG", batchNum, "-", total, "base=" baseJpg)

    ToolTip "BATCH " batchNum ": Exporting once from Scans root..."
    LogEvent("BATCH_EXPORT_BEGIN", batchNum, "-", total, "-")

    ok := ExportAllFromScansRoot(outDir)
    ToolTip

    if !ok {
        LogEvent("ERROR_EXPORT_TIMEOUT", batchNum, "-", total, "-")
        MsgBox "Export did not reach idle/finish state. Stopping to avoid deleting scans mid-export."
        ExitApp 1
    }

    LogEvent("BATCH_EXPORT_OK", batchNum, "-", total, "-")

    ; --- Verify JPGs exist AFTER export (by growth) ---
    ToolTip "BATCH " batchNum ": Verifying JPGs were created..."
    jpgOk := WaitForJpgGrowth(outDir, batchCount, baseJpg, JPG_WAIT_MAX_SEC, JPG_IDLE_MS)
    ToolTip

    if !jpgOk {
        LogEvent("ERROR_JPG_NOT_CREATED", batchNum, "-", total, "expected_inc=" batchCount ", base=" baseJpg)
        MsgBox "Export finished but JPG count did NOT increase by at least " batchCount ". Stopping (won't delete scans, won't mark processed)."
        ExitApp 1
    }

    afterJpg := CountJpgs(outDir)
    LogEvent("BATCH_JPG_OK", batchNum, "-", total, "after=" afterJpg)

    ToolTip "BATCH " batchNum ": Export finished. Deleting Scans..."
    LogEvent("BATCH_DELETE_BEGIN", batchNum, "-", total, "-")

    DeleteScansRoot()

    LogEvent("BATCH_DELETE_OK", batchNum, "-", total, "-")

    ; Commit processed ONLY after export OK + JPG verified + delete done
    AppendDoneList(donePath, batchPaths)
    for _, p in batchPaths
        doneSet[p] := true

    LogEvent("BATCH_COMMIT_OK", batchNum, "-", total, "-")

    ToolTip
    Sleep 800
}

LogEvent("DONE", "-", "-", total, "-")
MsgBox "Done. Processed " total " scans in batches of " BATCH_SIZE "."
ExitApp 0