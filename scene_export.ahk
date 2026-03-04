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

; WAIT logic
EXPORT_IDLE_MS := 25000        ; must be QUIET for 5s to consider export finished
EXPORT_MAX_SEC := 7200        ; 2 hours max safety timeout

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
    ; window title is Confirm, but verify its text
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
    ; checks for text fragment in any Static control
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

    ; NEW: verify it's the right dialog
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

        ; 1) file exists confirm
        if WinExist("Confirm") {
            HandleFileExistsConfirmOnce()
            handled := true
            lastDialogSeen := A_TickCount
        }

        ; 2) scan modified prompt (can be MANY)
        ; handle repeatedly in case a new one pops instantly
        if HandleScanModifiedOnce() {
            handled := true
            lastDialogSeen := A_TickCount
        }

        ; If nothing handled, check idle timer
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

    ; focus Scans root
    ClickAtWindow(SCANS_ROOT_X, SCANS_ROOT_Y, "Left")
    Sleep 200

    ; open context menu
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

    ; export -> panoramic images -> scan resolution
    Send "e"
    Sleep 200
    Send "p"
    Sleep 200
    Send "s"
    Sleep 600

    ; wait folder picker
    if !WinWaitActive(exportDlgTitle, , 20) {
        MsgBox "Did not reach export folder picker."
        return false
    }

    ; dialog-relative -> screen coords
    WinGetPos &dlgX, &dlgY, &dlgW, &dlgH, exportDlgTitle
    fieldX := dlgX + EXPORT_FOLDER_FIELD_WX
    fieldY := dlgY + EXPORT_FOLDER_FIELD_WY
    btnX   := dlgX + EXPORT_SELECT_BTN_WX
    btnY   := dlgY + EXPORT_SELECT_BTN_WY

    if useBlockInput
        BlockInput true

    CoordMode "Mouse", "Screen"

    ; paste outDir
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

    ; click Select Folder ONCE
    MouseMove btnX, btnY, 0
    Sleep 50
    Click "Left"
    Sleep 250

    if useBlockInput
        BlockInput false

    CoordMode "Mouse", "Window"

    ; >>> CRITICAL: WAIT until export has really finished <<<
    ToolTip "Waiting for export to finish (draining dialogs)..."
    ok := DrainExportCompletion(EXPORT_MAX_SEC, EXPORT_IDLE_MS)
    ToolTip
    return ok
}

; ---------- Delete Scans root after export finished ----------
DeleteScansRoot() {
    global SCANS_ROOT_X, SCANS_ROOT_Y
    ActivateScene()

    ; select scans root
    ClickAtWindow(SCANS_ROOT_X, SCANS_ROOT_Y, "Left")
    Sleep 150

    ; press Delete key
    Send "{Del}"
    Sleep 150

    ; confirm Yes
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

; ---------- Main ----------
flsFiles := ReadManifest(manifest)
total := flsFiles.Length
if (total = 0) {
    MsgBox "Manifest has 0 paths."
    ExitApp 2
}

idx := 1
batchNum := 0

while (idx <= total) {
    batchNum++
    batchStart := idx
    batchEnd := Min(idx + BATCH_SIZE - 1, total)
    batchCount := batchEnd - batchStart + 1

    ; IMPORT batch
    Loop batchCount {
        flsPath := flsFiles[idx]
        ToolTip "BATCH " batchNum ": Importing " (idx) "/" total "`n" flsPath
        ImportOneFls(flsPath)
        idx++
        Sleep 300
    }

    ToolTip "BATCH " batchNum ": Exporting once from Scans root..."
    ok := ExportAllFromScansRoot(outDir)
    ToolTip

    if !ok {
        MsgBox "Export did not reach idle/finish state. Stopping to avoid deleting scans mid-export."
        ExitApp 1
    }

    ToolTip "BATCH " batchNum ": Export finished. Deleting Scans..."
    DeleteScansRoot()
    ToolTip

    Sleep 800
}

MsgBox "Done. Processed " total " scans in batches of " BATCH_SIZE "."
ExitApp 0