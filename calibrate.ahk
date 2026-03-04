#SingleInstance Force
CoordMode "Mouse", "Window"
SetTitleMatchMode 2

; Hover your mouse over the thing you want to click, then press Ctrl+Alt+C
^!c::
{
    MouseGetPos &x, &y
    ToolTip "Window-relative: x=" x " y=" y
    A_Clipboard := x "," y
    SetTimer () => ToolTip(), -1500
}

^!q:: {
  ExitApp
}