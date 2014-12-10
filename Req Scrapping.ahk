#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetKeyDelay, 100
esc::exitapp
ralt::pause
#r::reload

CpyCell(x,y)
{
clipboard =  
winActivate, Microsoft Excel
	click %x%,%y%
	send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
}

PsteCellinVis:
winActivate, newbox.r2w - WRQ Reflection for UNIX and Digital
click 932,200
send %clipboard%
send {enter}
return

CpyVis(x1,x2)
{
Msgbox, flag6
clipboard =  
winActivate, newbox.r2w
	click 1,%x1%,%my%
	click down
	mousemove, %x2%,%my%
	click up
	send !e
	send c
	Msgbox, flag5 
}



#d::
loop,
{
winActivate, newbox.r2w
i = 5
if (i<6){
	Msgbox, flag1 %i%
	y = 400
	my := 400+(i*72)
	CpyVis(365,170)
	Msgbox, flag2
	winActivate, Microsoft Excel
	click 67,194
	send ^v
	send {down}
	i := i+1
	} else{
		Msgbox, flag3 
		send enter
		i=0
	}
}




