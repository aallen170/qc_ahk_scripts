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
note1 = PLANT HOLD REMOVED. 7-9-18. PA.
note2 = PLANT HOLD REMOVED. 7-10-18. PA.
note3 = PLANT HOLD REMOVED. 7-11-18. PA.
note4 = PLANT HOLD REMOVED. 7-12-18. PA.
note5 = PLANT HOLD REMOVED. 7-13-18. PA.
InputBox, UserInput, Number of Materials, Please enter the number of materials to delete
matMax = %UserInput%
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 200
Send, {CTRLDOWN}{HOME}{CTRLUP}{DOWN} ; {RIGHT}
while (matCount < matMax)
{
stuck =
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
Sleep, 200
MouseClick, left,  1065,  200
SetTitleMatchMode, 1
WinWait, Notes, , 0.5
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes, , 0.5
SetTitleMatchMode, 2
atNotes = 0
while(atNotes == 0)
{
IfWinNotActive, Notes
{
MouseClick, left,  1065,  200
}
IfWinActive, Notes
{
atNotes = 1
}
}
clipboard = 
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 200
delete = 0
rotate = 0
currentCopy =
prevCopy =
currentCopy = %clipboard%
while(rotate == 0)
{
if(currentCopy == prevCopy)
{
rotate = 1
}
else
{
prevCopy = %currentCopy%
Send, {TAB}
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 200
currentCopy = %clipboard%
}
}
if((clipboard == note1) or (clipboard == note2) or (clipboard == note3) or (clipboard == note4) or (clipboard == note5))
{
delete = 1
}
if(true)
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
atEdit = 0
while(atEdit == 0)
{
IfWinNotActive, ZQUAL editing screen
{
MouseClick, left,  105,  200
MouseClick, left,  105,  200
}
IfWinActive, ZQUAL editing screen
{
atEdit = 1
}
}
Sleep, 100
MouseClick, left, 243, 328
Sleep, 100
Send, {BACKSPACE}
Sleep, 100
Send, {CTRLDOWN}s{CTRLUP}
Sleep, 100
EPA = 0
IfWinActive, Information
{
EPA = 1
Sleep, 100
Send, {ENTER}
WinWait, ZQUAL editing screen, , 0.5
IfWinNotActive, ZQUAL editing screen, , WinActivate, ZQUAL editing screen, 
WinWaitActive, ZQUAL editing screen, , 0.5
clipboard =
MouseClick, left, 237, 746
Sleep, 100
Send, {CTRLDOWN}a{CTRLUP}
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 100
if(clipboard == emptyStr)
{
Send, N
Sleep, 100
Send, {TAB}N
Sleep, 100
Send, {TAB}N
WinWait, ZQUAL editing screen, , 0.5
IfWinNotActive, ZQUAL editing screen, , WinActivate, ZQUAL editing screen, 
WinWaitActive, ZQUAL editing screen, , 0.5
Sleep, 100
MouseClick, left, 243, 328
Send, {BACKSPACE}
Sleep, 100
Send, {CTRLDOWN}s{CTRLUP}
Sleep, 200
WinWait, Information, , 0.5
IfWinNotActive, Information, WinActivate, Information, 
WinWaitActive, Information, , 0.5
IfWinActive Information
{
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
active = 0
while(active = 0)
{
IfWinActive, Microsoft Excel
{
Send, {RIGHT}RX{LEFT}
active = 1
}
}
WinWait, ZQUAL editing screen,
IfWinNotActive, ZQUAL editing screen, , WinActivate, ZQUAL editing screen, 
WinWait, Information, 
IfWinNotActive, Information, WinActivate, Information, 
WinWaitActive, Information,
Sleep, 100
Send, {ENTER}
Sleep, 200
}
else
{
IfWinActive, Changed attributes will be copied to following items:
{
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {ENTER}
}
Sleep, 500
Send, {ENTER}
WinWait, Information, 
IfWinNotActive, Information, WinActivate, Information, 
WinWaitActive, Information,
IfWinActive, Information
{
Sleep, 100
Send, {ENTER}
}
Sleep, 1000
}
}
else
{
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
active = 0
while(active = 0)
{
IfWinActive, Microsoft Excel
{
Send, {RIGHT}RX{LEFT}
active = 1
}
}
WinWait, ZQUAL editing screen,
IfWinNotActive, ZQUAL editing screen, , WinActivate, ZQUAL editing screen, 
WinWaitActive, ZQUAL editing screen,
}
}
else if(EPA == 0)
{
MouseClick, left,  278,  135
Sleep, 100
IfWinActive, Changed attributes will be copied to following items:
{
Sleep, 100
Send, {TAB}
Sleep, 100
Send, {ENTER}
}
Sleep, 500
Send, {ENTER}
WinWait, Information, , 1
IfWinNotActive, Information, WinActivate, Information, 
WinWaitActive, Information, , 1
IfWinActive, Information
{
Sleep, 100
Send, {ENTER}
}
else
{
stuck = 1
}
if(stuck == 1)
{
MouseClick, left, 538, 55
WinWait, ZQUAL editing screen,
IfWinNotActive, ZQUAL editing screen, , WinActivate, ZQUAL editing screen, 
WinWaitActive, ZQUAL editing screen,
Send, {ALTDOWN}{F4}{ALTUP}
WinWait, SAP Easy Access, 
IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
WinWaitActive, SAP Easy Access,
Sleep, 200
MouseClick, left, 187, 278
MouseClick, left, 187, 278
WinWait, invalid.txt - Notepad, 
IfWinNotActive, invalid.txt - Notepad, , WinActivate, invalid.txt - Notepad, 
WinWaitActive, invalid.txt - Notepad, 
Send, %currentMat%{ENTER}
}
if(stuck != 1)
{
Sleep, 1000
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
if(delete == 0)
{
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
Send, {CTRLDOWN}s{CTRLUP}
Sleep, 100
WinWait, successfullyDeleted.txt - Notepad, 
IfWinNotActive, successfullyDeleted.txt - Notepad, , WinActivate, successfullyDeleted.txt - Notepad, 
WinWaitActive, successfullyDeleted.txt - Notepad, 
Send, %currentMat%{ENTER}
}
SetTitleMatchMode, 1
WinWait, Notes, 
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes,
SetTitleMatchMode, 2
}
}
}
else
{
if(delete == 0)
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
}
if(stuck != 1)
{
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
}
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