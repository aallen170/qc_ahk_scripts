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
Send, {CTRLDOWN}c{CTRLUP}
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
; WinWait, plantNums.txt - Notepad, 
; IfWinNotActive, plantNums.txt - Notepad, , WinActivate, plantNums.txt - Notepad, 
; WinWaitActive, plantNums.txt - Notepad,
; Sleep, 100
; Send, %currentMat% = %plantNums%
; Send, {ENTER}{ENTER}
Sleep, 100
IfWinActive, Information
{
MsgBox, HOLD IS SET TO ALL PLANTS
Sleep, 100
Send, {ENTER}
}
else
{
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing,
MsgBox, WRITE DOWN PLANT NUMS
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing,
Send, {F3}
WinWait, ZQUALN - Quality control, , 0.5
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control, , 0.5
ifWinActive, ZQUALN - Quality control
{
Send, {SHIFTDOWN}{F3}{SHIFTUP}
WinWait, SAP Easy Access, 
IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
WinWaitActive, SAP Easy Access,
MouseClick, left, 190, 406
MouseClick, left, 190, 406
WinWait, Stock Overview By Material, 
IfWinNotActive, Stock Overview By Material, , WinActivate, Stock Overview By Material, 
WinWaitActive, Stock Overview By Material, 
Send, {SHIFTDOWN}{TAB}{SHIFTUP}{ENTER}
MsgBox, CHECK IF DISCONTINUED
WinWait, Stock Overview By Material, 
IfWinNotActive, Stock Overview By Material, , WinActivate, Stock Overview By Material, 
WinWaitActive, Stock Overview By Material, 
Send, {F3}
Sleep, 200
Send, {SHIFTDOWN}{F3}{SHIFTUP}
WinWait, SAP Easy Access, 
IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
WinWaitActive, SAP Easy Access,
MouseClick, left, 180, 280
MouseClick, left, 180, 280
}
else
{
MsgBox, CHECK WHICH PLANT IT DOESN'T EXIST IN
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing,
MouseClick, left,  536, 55
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing,
Send, {ALTDOWN}{F4}{ALTUP}
WinWait, SAP Easy Access, 
IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
WinWaitActive, SAP Easy Access,
MouseClick, left, 190, 406
MouseClick, left, 190, 406
WinWait, Stock Overview By Material, 
IfWinNotActive, Stock Overview By Material, , WinActivate, Stock Overview By Material, 
WinWaitActive, Stock Overview By Material, 
Send, {SHIFTDOWN}{TAB}{SHIFTUP}{ENTER}
MsgBox, CHECK IF DISCONTINUED
Send, {F3}
Sleep, 200
Send, {SHIFTDOWN}{F3}{SHIFTUP}
WinWait, SAP Easy Access, 
IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
WinWaitActive, SAP Easy Access,
MouseClick, left, 180, 280
MouseClick, left, 180, 280
}
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
Sleep, 1000
}
ExitApp
Esc::ExitApp
