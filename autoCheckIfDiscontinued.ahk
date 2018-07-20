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
Send, {CTRLDOWN}{HOME}{CTRLUP}{DOWN}{RIGHT} ; {RIGHT}
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
WinWait, Stock Overview By Material, 
IfWinNotActive, Stock Overview By Material, , WinActivate, Stock Overview By Material, 
WinWaitActive, Stock Overview By Material, 
Sleep, 200
Send, {CTRLDOWN}v{CTRLUP}
currentMat = %clipboard%
Sleep, 200
Send, {SHIFTDOWN}{TAB}{SHIFTUP}{ENTER}
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
InputBox, disco, Active/Discontinued, ENTER "A"/"D" FOR ACTIVE/DISCONTINUED
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel,
Send, {LEFT}%disco%{RIGHT}
WinWait, Stock Overview By Material, 
IfWinNotActive, Stock Overview By Material, , WinActivate, Stock Overview By Material, 
WinWaitActive, Stock Overview By Material, 
Send,{F3}
WinWait, Microsoft Excel, 
IfWinNotActive, Microsoft Excel, , WinActivate, Microsoft Excel, 
WinWaitActive, Microsoft Excel, 
Send, {DOWN}
matCount++
}
ExitApp
Esc::ExitApp