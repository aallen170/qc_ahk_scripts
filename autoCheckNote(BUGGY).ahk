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
itemNumber = 0
InputBox, UserInput, Number of Materials, Please enter the number of materials for this division
matMax = %UserInput%
InputBox, UserInput, Number of Item Number, Please enter the number of the first material as it appears on the sheet
itemNumber = %UserInput%
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 200
Send, {CTRLDOWN}{HOME}{CTRLUP}{DOWN}{RIGHT}
while (matCount < matMax)
{
noNotes = 0
yesNotes = 0
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 200
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 200
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control,
Sleep, 200
Send, {CTRLDOWN}v{CTRLUP}
Sleep, 200
Send, {SHIFTDOWN}{TAB}{TAB}{SHIFTUP}{ENTER}
Sleep, 200
MouseClick, left,  1065,  200
WinWait, Notes, 
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes,
clipboard = %emptyStr%
MsgBox, clipboard after clearing = %clipboard%
Send, {CTRLDOWN}c{CTRLUP}
MsgBox, clipboard after copying = %clipboard%
if(clipboard == emptyStr)
{
noNotes = 1
}
else
{
yesNotes = 1
}
if(noNotes == 1)
{
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Send, {CTRLDOWN}c{CTRLUP}
WinWait, noNotes.txt - Notepad, 
IfWinNotActive, noNotes.txt - Notepad, , WinActivate, noNotes.txt - Notepad, 
WinWaitActive, noNotes.txt - Notepad,
Send, Item #%itemNumber%: %clipboard%{ENTER}{ENTER}
}
if(yesNotes == 1)
{
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Send, {CTRLDOWN}c{CTRLUP}
WinWait, yesNotes.txt - Notepad, 
IfWinNotActive, yesNotes.txt - Notepad, , WinActivate, yesNotes.txt - Notepad, 
WinWaitActive, yesNotes.txt - Notepad,
Send, Item #%itemNumber%: %clipboard%{ENTER}{ENTER}
}
WinWait, Notes, 
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes,
Send, {F3}
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control, 
Send, {SHIFTDOWN}{F3}{SHIFTUP}
WinWait, SAP Easy Access, 
IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
WinWaitActive, SAP Easy Access,
Send, {ENTER}
Sleep, 200
Send, {TAB}
Sleep, 200
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Send, {DOWN}
itemNumber++
matCount++
Sleep, 200
}
ExitApp
Esc::ExitApp