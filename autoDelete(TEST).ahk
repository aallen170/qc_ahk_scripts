SetTitleMatchMode, 2
matCount = 0
matMax = 0
posY = 0
currentClip =
prevClip =
noteLength = 0
continuing = 0
emptyStr =
currentMat =
InputBox, UserInput, Number of Materials, Please enter the number of materials to delete
matMax = %UserInput%
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 200
Send, {CTRLDOWN}{HOME}{CTRLUP}{DOWN}{RIGHT}
while (matCount < matMax)
{
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 200
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control,
Sleep, 200
Send, {CTRLDOWN}v{CTRLUP}
currentMat = %clipboard%
Sleep, 200
Send, {SHIFTDOWN}{TAB}{TAB}{SHIFTUP}{ENTER}
Sleep, 300
MouseClick, left,  1065,  200
SetTitleMatchMode, 1
WinWait, Notes, , 0.5
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes, , 0.5
SetTitleMatchMode, 2
IfWinNotActive, Notes
{
MouseClick, left,  1065,  200
}
clipboard = 
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 200
if(clipboard == emptyStr)
{
Send, {F3}
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control,
MouseClick, left,  105,  200
MouseClick, left,  105,  200
WinWait, ZQUAL editing screen, , 0.5
IfWinNotActive, ZQUAL editing screen, , WinActivate, ZQUAL editing screen, 
WinWaitActive, ZQUAL editing screen, , 0.5
IfWinNotActive, ZQUAL editing screen
{
MouseClick, left,  105,  200
MouseClick, left,  105,  200
}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {BACKSPACE}
Sleep, 100
Send, {F3}
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control,
Sleep, 200
MouseClick, left,  1065,  200
SetTitleMatchMode, 1
WinWait, Notes, , 0.5
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes, , 0.5
SetTitleMatchMode, 2
IfWinNotActive, Notes
{
MouseClick, left,  1065,  200
}
while(continuing == 0)
{
Send, {CTRLDOWN}c{CTRLUP}
currentClip = %clipboard%
; MsgBox, Clipboard: %clipboard%`ncurrentClip: %currentClip%
if(noteLength == 12)
{
MsgBox, Too many notes!
}
else
{
if(currentClip != prevClip)
{
prevClip = %currentClip%
noteLength++
Send, {TAB}
Sleep, 100
}
else
{
continuing = 1
}
}
}
continuing = 0
noteLength = 0
prevClip = emptyStr
currentClip = emptyStr
WinWait, phyllis_remove_note.txt - Notepad, 
IfWinNotActive, phyllis_remove_note.txt - Notepad, , WinActivate, phyllis_remove_note.txt - Notepad, 
WinWaitActive, phyllis_remove_note.txt - Notepad, 
Sleep, 100
Send, {CTRLDOWN}a{CTRLUP}
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 100
SetTitleMatchMode, 1
WinWait, Notes, 
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes,
SetTitleMatchMode, 2
Sleep, 200
Send, {CTRLDOWN}v{CTRLUP}
Sleep, 200
 ; Send, {CTRLDOWN}s{CTRLUP}
 ; Sleep, 800
}
else
{
WinWait, yesNote.txt - Notepad, 
IfWinNotActive, yesNote.txt - Notepad, , WinActivate, yesNote.txt - Notepad, 
WinWaitActive, yesNote.txt - Notepad, 
Send, %currentMat%{ENTER}
SetTitleMatchMode, 1
WinWait, Notes, , 0.5
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes, , 0.5
SetTitleMatchMode, 2
}
Send, {F3}
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control,
Sleep, 100 
Send, {SHIFTDOWN}{F3}{SHIFTUP}
WinWait, SAP Easy Access, 
IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
WinWaitActive, SAP Easy Access,
Send, {ENTER}
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control,
Send, {TAB}
Sleep, 200
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Send, {DOWN}
matCount++
}
ExitApp
Esc::ExitApp