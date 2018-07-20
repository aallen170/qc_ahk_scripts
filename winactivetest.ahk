WinWait, Untitled - Notepad, 
IfWinNotActive, Untitled - Notepad, , WinActivate, Untitled - Notepad, 
WinWaitActive, Untitled - Notepad, 
IfWinActive, Untitled - Notepad
{
    WinMaximize  ; Maximizes the Notepad window found by IfWinActive above.
    Send, Some text.{Enter}
    return
}