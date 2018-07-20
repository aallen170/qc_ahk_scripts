SetTitleMatchMode, 2
matCount = 0
matMax = 0
posY = 0
currentClip =
prevClip =
noteLength = 0
continuing = 0
emptyStr =
currentStr =
prevStr =
currentMat =
InputBox, UserInput, Number of Materials, Please enter the number of materials to check
matMax = %UserInput%
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 200
Send, {CTRLDOWN}{HOME}{CTRLUP}{DOWN}{RIGHT}{RIGHT}
while (matCount < matMax)
{
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 200
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 500
currentMat = %clipboard%
Sleep, 500
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control,
Sleep, 200
Send, {CTRLDOWN}v{CTRLUP}
Sleep, 200
Send, {SHIFTDOWN}{TAB}{TAB}{SHIFTUP}{ENTER}
Sleep, 200
MouseClick, left,  980, 200
MouseClick, left,  980, 200
Sleep, 200
; InputBox, UserInput, Number of Materials, Please enter the plant numbers you see
; plantNums = %UserInput%
plantNums =
posY = 262
MouseClickDrag, L, 215, posY, 194, posY
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
currentStr = %clipboard%
WinWait, plantNums.txt - Notepad, 
IfWinNotActive, plantNums.txt - Notepad, , WinActivate, plantNums.txt - Notepad, 
WinWaitActive, plantNums.txt - Notepad,
Sleep, 100
Send, %currentMat%
Sleep, 100
while(currentStr != prevStr)
{
WinWait, plantNums.txt - Notepad, 
IfWinNotActive, plantNums.txt - Notepad, , WinActivate, plantNums.txt - Notepad, 
WinWaitActive, plantNums.txt - Notepad,
Sleep, 100
prevStr = %currentStr%
plantNums = %currentStr%
Sleep, 200
Send, {ENTER}
Send, %plantNums%
posY += 22
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing,
Sleep, 100
MouseClickDrag, L, 215, posY, 194, posY
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
currentStr = %clipboard%
}
WinWait, plantNums.txt - Notepad, 
IfWinNotActive, plantNums.txt - Notepad, , WinActivate, plantNums.txt - Notepad, 
WinWaitActive, plantNums.txt - Notepad,
Send, {ENTER}{ENTER}
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing, 
MsgBox, wait
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
Sleep, 200
Send, {TAB}
Sleep, 200
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Send, {DOWN}
matCount++
Sleep, 200
currentStr =
prevStr =
}
ExitApp
Esc::ExitApp