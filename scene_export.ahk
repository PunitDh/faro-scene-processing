^q:: {
    MsgBox "Script aborted."
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

; Split staging outputs (so Python can pick them up reliably)
outDirFull := outDir "\FullColor"
outDirScan := outDir "\ScanResolution"

if !DirExist(outDirFull)
    DirCreate outDirFull
if !DirExist(outDirScan)
    DirCreate outDirScan

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

; --- Load All Scans wait ---
LOADALL_QUIET_MS := 5000     ; must be gone for 5s
LOADALL_MAX_SEC  := 7200     ; safety timeout (2h)

; --- Full Color export progress wait ---
FULLCOLOR_QUIET_MS := 5000    ; must be gone for 5s
FULLCOLOR_MAX_SEC  := 7200    ; safety timeout (2h)
FULLCOLOR_MUSTSEE_MS := 60000 ; must detect the dialog at least once within 60s

exportDlgTitle := "Select folder for images export"

useBlockInput := true

; Export dialog (WINDOW RELATIVE to export dialog)
EXPORT_FOLDER_FIELD_WX := 460
EXPORT_FOLDER_FIELD_WY := 512
EXPORT_SELECT_BTN_WX   := 715
EXPORT_SELECT_BTN_WY   := 550

; WAIT logic (existing)
EXPORT_IDLE_MS := 10000         ; must be QUIET for 10s to consider export finished
EXPORT_MAX_SEC := 7200          ; 2 hours max safety timeout

; --- logging / resume ---
logCsvPath := A_ScriptDir "\log\scene_export_log.csv"
donePath   := A_ScriptDir "\log\processed.txt"

; --- JPG verification (on Scan-Resolution EXPORT only) ---
JPG_WAIT_MAX_SEC := 3600        ; 1 hour max wait for jpg growth after export completes
JPG_IDLE_MS      := 8000        ; if jpg count doesn't change for 8s, we treat as stalled (fail)

; --- NEW: PNG verification (FullColor) ---
PNG_WAIT_MAX_SEC := 7200        ; can be long
PNG_IDLE_MS      := 10000       ; if png count doesn't change for 10s, we treat as stalled (fail)

; =========================

; Ensure log dir exists
if !DirExist(A_ScriptDir "\log")
    DirCreate A_ScriptDir "\log"

; ---------- Logging helpers ----------
NowStamp() {
    return FormatTime(, "yyyy-MM-dd HH:mm:ss")
}

CsvEscape(s) {
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

; ---------- JPG verification helpers ----------
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

; ---------- PNG verification helpers ----------
CountPngs(dir) {
    c := 0
    Loop Files dir "\*.png", "FR"
        c++
    return c
}

WaitForPngGrowth(dir, expectedIncrease, baseCount, maxSec := 7200, idleMs := 10000) {
    target := baseCount + expectedIncrease
    start := A_TickCount
    lastChange := A_TickCount
    lastCount := baseCount

    Loop {
        if (A_TickCount - start) > (maxSec * 1000)
            return false

        cur := CountPngs(dir)
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
; EXISTING WORKING FUNCTIONS
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
; ---------- Import (ROBUST) ----------
FindImportDialogHwnd(timeoutMs := 15000) {
    start := A_TickCount
    Loop {
        if (A_TickCount - start) > timeoutMs
            return 0

        ; scan all dialogs
        list := WinGetList("ahk_class #32770")
        for _, hwnd in list {
            ; must be SCENE-owned dialog (helps avoid random dialogs)
            try proc := WinGetProcessName("ahk_id " hwnd)
            catch
                continue
            if !(proc = "SCENE.exe" || proc = "Scene.exe")
                continue

            ; Import dialog should have an Edit1 path box
            if !ControlExists("Edit1", "ahk_id " hwnd)
                continue

            ; Usually has a button like Open/OK/Import as Button1 or Button2
            ; We don't care which—Edit1 is the key.
            return hwnd
        }
        Sleep 100
    }
}

ImportOneFls(flsPath) {
    global IMPORT_BTN_X, IMPORT_BTN_Y, waitImportMs
    global IMPORT_RESULTS_OK_X, IMPORT_RESULTS_OK_Y
    global TREE_X, TREE_Y
    global importResultsOkDelayMs, useBlockInput

    ActivateScene()

    ; click Import
    ClickAtWindow(IMPORT_BTN_X, IMPORT_BTN_Y, "Left")
    Sleep 150

    ; wait for the import dialog by HWND, not title
    hwnd := FindImportDialogHwnd(15000)
    if !hwnd {
        MsgBox "Import dialog not found after clicking Import."
        ExitApp 1
    }

    WinActivate "ahk_id " hwnd
    WinWaitActive "ahk_id " hwnd, , 5

    ; Set path and confirm
    ControlSetText flsPath, "Edit1", "ahk_id " hwnd
    Sleep 150
    Send "{Enter}"
    Sleep 300

    ; wait for SCENE’s custom “Import results” OK
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

; ---------- “wait until export is done” gate for Scan Resolution ----------
FindExportFolderPickerHwnd(timeoutMs := 10000) {
    start := A_TickCount
    Loop {
        if (A_TickCount - start) > timeoutMs
            return 0

        list := WinGetList("ahk_class #32770")
        for _, hwnd in list {
            try proc := WinGetProcessName("ahk_id " hwnd)
            catch
                continue
            if !(proc = "SCENE.exe" || proc = "Scene.exe")
                continue

            try title := WinGetTitle("ahk_id " hwnd)
            catch
                continue
            if !InStr(title, "Select folder for images export")
                continue

            ; the real picker has an Edit1 path field
            if !ControlExists("Edit1", "ahk_id " hwnd)
                continue

            return hwnd
        }
        Sleep 100
    }
}

CancelTreeRenameAndMenus() {
    ; ESC once usually cancels rename; twice is cheap insurance
    Send "{Esc}"
    Sleep 80
    Send "{Esc}"
    Sleep 80
}

DrainExportCompletion(maxSeconds := 7200, idleMs := 10000) {
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

ControlExists(ctrl, winTitle) {
    try {
        ControlGetHwnd(ctrl, winTitle)
        return true
    } catch {
        return false
    }
}

WindowHasStaticText(hwnd, needle) {
    Loop 20 {
        try txt := ControlGetText("Static" A_Index, "ahk_id " hwnd)
        catch
            continue
        if (txt != "" && InStr(txt, needle))
            return true
    }
    return false
}

; ---------- Load All Scans progress detection ----------
IsLoadAllScansProgress(hwnd) {
    try cls := WinGetClass("ahk_id " hwnd)
    catch
        return false
    if (cls != "#32770")
        return false

    title := ""
    try title := WinGetTitle("ahk_id " hwnd)
    catch
        return false
    if !InStr(title, "Load all scans")
        return false

    if ControlExists("Button1", "ahk_id " hwnd) {
        b := ""
        try b := ControlGetText("Button1", "ahk_id " hwnd)
        catch
            b := ""
        if InStr(b, "Abort")
            return true
    }

    if WindowHasStaticText(hwnd, "Load data")
        return true

    if ControlExists("msctls_progress321", "ahk_id " hwnd)
        return true
    if ControlExists("msctls_progress322", "ahk_id " hwnd)
        return true

    return false
}

FindLoadAllScansProgressHwnd() {
    list := WinGetList("ahk_class #32770")
    for _, hwnd in list {
        if IsLoadAllScansProgress(hwnd)
            return hwnd
    }
    return 0
}

WaitForLoadAllScansDone(maxSec := 7200, quietMs := 5000, mustSeeMs := 60000) {
    start := A_TickCount
    firstSeen := 0
    lastSeen := 0

    Loop {
        now := A_TickCount
        if (now - start) > (maxSec * 1000)
            return false

        hwnd := FindLoadAllScansProgressHwnd()
        if (hwnd) {
            if (!firstSeen)
                firstSeen := now
            lastSeen := now
            Sleep 200
            continue
        }

        if (!firstSeen) {
            if ((now - start) > mustSeeMs)
                return false
            Sleep 200
            continue
        }

        if ((now - lastSeen) > quietMs)
            return true

        Sleep 200
    }
}

; ---------- Panoramic success dialog (OK) ----------
IsPanoramicSuccessDialog(hwnd) {
    try cls := WinGetClass("ahk_id " hwnd)
    catch
        return false
    if (cls != "#32770")
        return false

    if WindowHasStaticText(hwnd, "Successfully created panoramic images")
        return true

    return false
}

FindPanoramicSuccessHwnd() {
    list := WinGetList("ahk_class #32770")
    for _, hwnd in list {
        if IsPanoramicSuccessDialog(hwnd)
            return hwnd
    }
    return 0
}

DismissPanoramicSuccessIfPresent() {
    hwnd := FindPanoramicSuccessHwnd()
    if !hwnd
        return false

    WinActivate "ahk_id " hwnd
    WinWaitActive "ahk_id " hwnd, , 2
    Sleep 80

    if ControlExists("Button1", "ahk_id " hwnd) {
        b := ""
        try b := ControlGetText("Button1", "ahk_id " hwnd)
        catch
            b := ""
        if (b = "OK") {
            ControlClick "Button1", "ahk_id " hwnd
            Sleep 150
            return true
        }
    }

    Send "{Enter}"
    Sleep 150
    return true
}

; ---------- Full Color export progress detection ----------
IsFullColorExportProgress(hwnd) {
    ; Dialog title like: "Creating full resolution panoramic images 8%"
    try cls := WinGetClass("ahk_id " hwnd)
    catch
        return false
    if (cls != "#32770")
        return false

    title := ""
    try title := WinGetTitle("ahk_id " hwnd)
    catch
        return false

    if !InStr(title, "Creating full resolution panoramic images")
        return false

    if ControlExists("Button1", "ahk_id " hwnd) {
        b := ""
        try b := ControlGetText("Button1", "ahk_id " hwnd)
        catch
            b := ""
        if InStr(b, "Abort")
            return true
    }

    if WindowHasStaticText(hwnd, "Creating full resolution panoramic images")
        return true

    return true
}

FindFullColorExportProgressHwnd() {
    list := WinGetList("ahk_class #32770")
    for _, hwnd in list {
        if IsFullColorExportProgress(hwnd)
            return hwnd
    }
    return 0
}

WaitForFullColorExportDone(maxSec := 7200, quietMs := 5000, mustSeeMs := 60000) {
    start := A_TickCount
    firstSeen := 0
    lastSeen := 0

    Loop {
        now := A_TickCount
        if (now - start) > (maxSec * 1000)
            return false

        ; drain known popups while waiting
        if WinExist("Confirm")
            HandleFileExistsConfirmOnce()
        HandleScanModifiedOnce()
        DismissPanoramicSuccessIfPresent()

        hwnd := FindFullColorExportProgressHwnd()
        if (hwnd) {
            if (!firstSeen)
                firstSeen := now
            lastSeen := now
            Sleep 200
            continue
        }

        if (!firstSeen) {
            if ((now - start) > mustSeeMs)
                return false
            Sleep 200
            continue
        }

        if ((now - lastSeen) > quietMs) {
            ; success dialog can appear after progress ends
            endStart := A_TickCount
            while (A_TickCount - endStart) < 8000 {
                if DismissPanoramicSuccessIfPresent()
                    break
                Sleep 200
            }
            return true
        }

        Sleep 200
    }
}

; ---------- Load All Scans ONCE (Right click Scans, Down x4, Enter) ----------
LoadAllScansOnce() {
    global SCANS_ROOT_X, SCANS_ROOT_Y, useBlockInput
    global LOADALL_MAX_SEC, LOADALL_QUIET_MS

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
        MsgBox "Context menu did not open for Scans root (Load All Scans)."
        return false
    }

    Send "{Down 4}{Enter}"
    Sleep 250

    ToolTip "Loading all scans... waiting for progress dialog to finish (quiet " (LOADALL_QUIET_MS/1000) "s)"
    ok := WaitForLoadAllScansDone(LOADALL_MAX_SEC, LOADALL_QUIET_MS, 60000)
    ToolTip

    return ok
}

; ---------- Helper: fill export folder picker ----------
FillExportFolderPicker(targetDir) {
    global exportDlgTitle
    global useBlockInput
    global EXPORT_FOLDER_FIELD_WX, EXPORT_FOLDER_FIELD_WY
    global EXPORT_SELECT_BTN_WX, EXPORT_SELECT_BTN_WY

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
    A_Clipboard := targetDir
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
}

; ---------- Full Color export from Scans root (creates PNG) ----------
ExportAllFromScansRoot_FullColor(outDir) {
    global SCANS_ROOT_X, SCANS_ROOT_Y
    global useBlockInput
    global EXPORT_FOLDER_FIELD_WX, EXPORT_FOLDER_FIELD_WY
    global EXPORT_SELECT_BTN_WX, EXPORT_SELECT_BTN_WY
    global FULLCOLOR_MAX_SEC, FULLCOLOR_QUIET_MS, FULLCOLOR_MUSTSEE_MS

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
        MsgBox "Export context menu did not open (Full Color)."
        return "FAIL"
    }

    ; Export -> Panoramic Images -> Full Color Resolution
    Send "e"
    Sleep 200
    Send "p"
    Sleep 200
    Send "{Down}{Enter}"
    Sleep 250

    ; HARD GATE: the folder picker must REALLY exist
    dlgHwnd := FindExportFolderPickerHwnd(2500)
    if !dlgHwnd {
        ; Full Color likely disabled; prevent rename/paste chaos
        CancelTreeRenameAndMenus()
        DismissPanoramicSuccessIfPresent()
        return "SKIPPED"
    }

    WinActivate "ahk_id " dlgHwnd
    WinWaitActive "ahk_id " dlgHwnd, , 5

    WinGetPos &dlgX, &dlgY, &dlgW, &dlgH, "ahk_id " dlgHwnd
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

    ToolTip "Full Color export running..."
    ok := WaitForFullColorExportDone(FULLCOLOR_MAX_SEC, FULLCOLOR_QUIET_MS, FULLCOLOR_MUSTSEE_MS)
    ToolTip

    return ok ? "OK" : "FAIL"
}

; ---------- Scan Resolution export from Scans root (JPG) ----------
ExportAllFromScansRoot_Scan(outDirScan) {
    global SCANS_ROOT_X, SCANS_ROOT_Y
    global exportDlgTitle
    global useBlockInput
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

    ; Export -> Panoramic Images -> Scan Resolution
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

    FillExportFolderPicker(outDirScan)

    ToolTip "Waiting for Scan Resolution export to finish (draining dialogs)..."
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
; Main
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

    ; Load All Scans ONCE
    ToolTip "BATCH " batchNum ": Load All Scans (Down x4, Enter)..."
    LogEvent("BATCH_LOAD_ALL_BEGIN", batchNum, "-", total, "-")

    okLoad := LoadAllScansOnce()
    ToolTip

    if !okLoad {
        LogEvent("ERROR_LOAD_ALL_FAILED", batchNum, "-", total, "-")
        MsgBox "Load All Scans failed. Stopping before export."
        ExitApp 1
    }

    LogEvent("BATCH_LOAD_ALL_OK", batchNum, "-", total, "-")

    ; 1) FULL COLOR export (PNG)
    basePng := CountPngs(outDirFull)
    LogEvent("BATCH_EXPORT_BASE_PNG", batchNum, "-", total, "base=" basePng)

    ToolTip "BATCH " batchNum ": Exporting Full Color Resolution (PNG)..."
    LogEvent("BATCH_EXPORT_FULLCOLOR_BEGIN", batchNum, "-", total, "-")

    fc := ExportAllFromScansRoot_FullColor(outDirFull)
    ToolTip

    if (fc = "SKIPPED") {
        LogEvent("BATCH_EXPORT_FULLCOLOR_SKIPPED", batchNum, "-", total, "-")
    } else if (fc = "FAIL") {
        LogEvent("ERROR_FULLCOLOR_EXPORT_FAILED", batchNum, "-", total, "-")
        MsgBox "Full Color export failed. Stopping before Scan Resolution export."
        ExitApp 1
    } else {
        LogEvent("BATCH_EXPORT_FULLCOLOR_OK", batchNum, "-", total, "-")

        ; Do NOT require PNGs (some scans are B/W-only).
        ; Just log how many PNGs we ended up with.
        afterPng := CountPngs(outDirFull)
        incPng := afterPng - basePng
        LogEvent("BATCH_PNG_COUNT", batchNum, "-", total, "base=" basePng ", after=" afterPng ", inc=" incPng)
    }

    ; 2) Scan Resolution export (JPG)
    baseJpg := CountJpgs(outDirScan)
    LogEvent("BATCH_EXPORT_BASE_JPG", batchNum, "-", total, "base=" baseJpg)

    ToolTip "BATCH " batchNum ": Exporting Scan Resolution (JPG)..."
    LogEvent("BATCH_EXPORT_BEGIN", batchNum, "-", total, "-")

    ok := ExportAllFromScansRoot_Scan(outDirScan)
    ToolTip

    if !ok {
        LogEvent("ERROR_EXPORT_TIMEOUT", batchNum, "-", total, "-")
        MsgBox "Scan Resolution export did not reach idle/finish state. Stopping to avoid deleting scans mid-export."
        ExitApp 1
    }

    LogEvent("BATCH_EXPORT_OK", batchNum, "-", total, "-")

    ; Verify JPGs exist AFTER Scan Resolution export (by growth)
    ToolTip "BATCH " batchNum ": Verifying JPGs were created..."
    jpgOk := WaitForJpgGrowth(outDirScan, batchCount, baseJpg, JPG_WAIT_MAX_SEC, JPG_IDLE_MS)
    ToolTip

    if !jpgOk {
        LogEvent("ERROR_JPG_NOT_CREATED", batchNum, "-", total, "expected_inc=" batchCount ", base=" baseJpg)
        MsgBox "Export finished but JPG count did NOT increase by at least " batchCount ". Stopping (won't delete scans, won't mark processed)."
        ExitApp 1
    }

    afterJpg := CountJpgs(outDirScan)
    LogEvent("BATCH_JPG_OK", batchNum, "-", total, "after=" afterJpg)

    ; Delete scans
    ToolTip "BATCH " batchNum ": Exports finished. Deleting Scans..."
    LogEvent("BATCH_DELETE_BEGIN", batchNum, "-", total, "-")

    DeleteScansRoot()

    LogEvent("BATCH_DELETE_OK", batchNum, "-", total, "-")

    ; Commit processed ONLY after exports OK + JPG verified + delete done
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