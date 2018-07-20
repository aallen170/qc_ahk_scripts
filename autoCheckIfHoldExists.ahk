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
startAt = 0
InputBox, UserInput, Number of Materials, Please enter the number of materials for this division
matMax = %UserInput%
InputBox, startAt, Number to start at, enter the number shown on excel to start at
startAt := startAt - 2
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 200
Send, {CTRLDOWN}{HOME}{CTRLUP}{DOWN} ; {RIGHT}{RIGHT}
while(startAt > 0)
{
Send, {DOWN}
startAt--
matCount++
}
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
Sleep, 500
MouseClick, left, 982, 198
MouseClick, left, 982, 198
Sleep, 300
IfWinActive, Information
{
WinWait, Information, 
IfWinNotActive, Information, , WinActivate, Information, 
WinWaitActive, Information,
Sleep, 100
Send, {SHIFTDOWN}{END}{SHIFTUP}{CTRLDOWN}c{CTRLUP}
Sleep, 100
allHold = Hold is Set to All Plants
if(clipboard == allHold)
{
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 100
active = 0
while(active = 0)
{
IfWinActive, Microsoft Excel
{
Sleep, 100
Send, {RIGHT}A{LEFT}
Sleep, 100
active = 1
}
}
Sleep, 100
WinWait, Information, 
IfWinNotActive, Information, , WinActivate, Information, 
WinWaitActive, Information,
}
WinWait, Information, 
IfWinNotActive, Information, , WinActivate, Information, 
WinWaitActive, Information,
Send, {ENTER}
}
IfWinActive, Plant Hold Editing
{
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 100
active = 0
while(active = 0)
{
IfWinActive, Microsoft Excel
{
Sleep, 100
Send, {RIGHT}P{LEFT}
Sleep, 100
active = 1
}
}
Sleep, 100
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing, 
Send,{F3}
}
WinWait, ZQUALN - Quality control, , 0.5
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control, , 0.5
stuck = 0
IfWinNotActive ZQUALN - Quality control
{
stuck = 1
MouseClick, left, 533, 54
WinWait, SAP Easy Access, 
IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
WinWaitActive, SAP Easy Access,
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing, 
Send, {ALTDOWN}{F4}{ALTUP}
WinWait, SAP Easy Access, 
IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
WinWaitActive, SAP Easy Access,
MouseClick, left, 175, 278
MouseClick, left, 175, 278
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control,
}
if(stuck == 0)
{
Send, {SHIFTDOWN}{F3}{SHIFTUP}
WinWait, SAP Easy Access, 
IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
WinWaitActive, SAP Easy Access,
Send, {ENTER}
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control, 
}
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