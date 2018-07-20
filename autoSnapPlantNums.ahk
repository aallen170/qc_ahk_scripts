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
Send, {CTRLDOWN}{HOME}{CTRLUP}{DOWN} ; {RIGHT}{RIGHT}
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
posY = 262
MouseClickDrag, L, 215, posY, 194, posY
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
currentNum = %clipboard%
count = 0
prevNum =
while(currentNum != prevNum)
{
prevNum = %currentNum%
posY += 22
Sleep, 100
MouseClickDrag, L, 215, posY, 194, posY
Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}
currentNum = %clipboard%
count++
if(count == 31)
{
posY += 10
break
}
}
posY += 10
WinWait, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
IfWinNotActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, , WinActivate, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
WinWaitActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
Send, {CTRLDOWN}r{CTRLUP}
Sleep, 500
MouseClickDrag, L, 260, posY, 35, 160
WinWait, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
IfWinNotActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, , WinActivate, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
WinWaitActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 200
WinWait, Microsoft Word, 
IfWinNotActive, Microsoft Word, , WinActivate, Microsoft Word, 
WinWaitActive, Microsoft Word, 
Send, {CTRLDOWN}v{CTRLUP}
Sleep, 100
Send, {DOWN}
Sleep, 100
Send, {RIGHT}
Sleep, 200
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
MouseClick, left, 180, 280
MouseClick, left, 180, 280
}
else
{
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing,
MouseClick, left,  536, 55
Sleep, 200
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing,
Send, {ALTDOWN}{F4}{ALTUP}
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
Sleep, 200
}
ExitApp
Esc::ExitApp