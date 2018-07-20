SetTitleMatchMode, 2
matCount = 0
matMax = 0
posY = 0
currentClip =
prevClip =
noteLength = 0
continuing = 0
emptyStr =
currentNum =
prevNum =
currentMat =
InputBox, UserInput, Number of Materials, Please enter the number of materials for this division
matMax = %UserInput%
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 200
Send, {CTRLDOWN}{HOME}{CTRLUP}
while (matCount < matMax)
{
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Sleep, 200
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 500
currentMat = %clipboard%
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control,
Sleep, 200
Send, {CTRLDOWN}v{CTRLUP}
Sleep, 200
Send, {SHIFTDOWN}{TAB}{TAB}{SHIFTUP}{ENTER}
Sleep, 200
MouseClick, left,  1065,  200
Sleep, 200
while(continuing == 0)
{
Send, {CTRLDOWN}c{CTRLUP}
currentClip = %clipboard%
; MsgBox, Clipboard: %clipboard%`ncurrentClip: %currentClip%
if(noteLength == 12)
{
break
}
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
if(noteLength == 0)
{
posY = 275
}
if(noteLength == 1)
{
posY = 275
}
if(noteLength == 2)
{
posY = 300
}
if(noteLength == 3)
{
posY = 320
}
if(noteLength == 4)
{
posY = 340
}
if(noteLength == 5)
{
posY = 360
}
if(noteLength == 6)
{
posY = 385
}
if(noteLength == 7)
{
posY = 410
}
if(noteLength == 8)
{
posY = 430
}
if(noteLength == 9)
{
posY = 450
}
if(noteLength == 10)
{
posY = 475
}
if(noteLength == 11)
{
posY = 500
}
if(noteLength == 12)
{
posY = 500
}
noteLength = 0
continuing = 0
prevClip = 
currentClip = 
Sleep, 200
WinWait, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
IfWinNotActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, , WinActivate, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
WinWaitActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
Send, {CTRLDOWN}r{CTRLUP}
Sleep, 500
MouseClickDrag, L, 500, posY, 35, 160
Sleep, 200
WinWait, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
IfWinNotActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, , WinActivate, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
WinWaitActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 200
WinWait, Microsoft Word, 
IfWinNotActive, Microsoft Word, , WinActivate, Microsoft Word, 
WinWaitActive, Microsoft Word, 
Send, {CTRLDOWN}v{CTRLUP}{ENTER}
Sleep, 200
WinWait, Notes, 
IfWinNotActive, Notes, , WinActivate, Notes, 
WinWaitActive, Notes, 
Send, {F3}
MouseClickDrag, L, 180, 965, 400, 965
WinWait, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
IfWinNotActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, , WinActivate, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
WinWaitActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
Send, {CTRLDOWN}r{CTRLUP}
Sleep, 500
MouseClickDrag, L, 307, 220, 833, 140
WinWait, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
IfWinNotActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, , WinActivate, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
WinWaitActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
Send, {CTRLDOWN}c{CTRLUP}
Sleep, 200
WinWait, Microsoft Word, 
IfWinNotActive, Microsoft Word, , WinActivate, Microsoft Word, 
WinWaitActive, Microsoft Word, 
Send, {CTRLDOWN}v{CTRLUP}{ENTER}
Sleep, 200
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control, 
MouseClick, left,  330, 198
MouseClick, left,  330, 198
Sleep, 150
IfWinActive, Information
{
WinWait, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
IfWinNotActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, , WinActivate, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
WinWaitActive, PrintKey 2000   v5.10 Full (  English  )        Copyright (c) 1999 By Alfred Bolliger, 
Send, {CTRLDOWN}o{CTRLUP}
Sleep, 500
MouseClick, left, 291, 357
Sleep, 200
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
WinWait, Information, 
IfWinNotActive, Information, , WinActivate, Information, 
WinWaitActive, Information, 
Send, {ENTER}
WinWait, ZQUALN - Quality control, 
IfWinNotActive, ZQUALN - Quality control, , WinActivate, ZQUALN - Quality control, 
WinWaitActive, ZQUALN - Quality control, 
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
WinWait, Plant Hold Editing, 
IfWinNotActive, Plant Hold Editing, , WinActivate, Plant Hold Editing, 
WinWaitActive, Plant Hold Editing,
Sleep, 200
}
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
}
ExitApp
Esc::ExitApp