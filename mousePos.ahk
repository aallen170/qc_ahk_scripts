; This example allows you to move the mouse around to see
; the title of the window currently under the cursor:
#Persistent
SetTimer, WatchCursor, 100
return

WatchCursor:
MouseGetPos, xpos, ypos
ToolTip, Xpos: %xpos%`nYpos: %ypos%

esc::exitapp