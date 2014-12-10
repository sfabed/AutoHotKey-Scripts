#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 50
#p::Pause
ESC::ExitApp
ralt::pause
#r::reload

lalt::
i=0
send {enter}
pause
return



; Function area!

CpyCell(x,y)
{
clipboard =  
winActivate, Microsoft Excel - DanPartsv1
	click %x%,%y%
	send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
}

#d::
loop,
{
i=1
CpyCell(67,194)
winActivate, Microsoft Excel - Q1 '14 Physical Count
send ^f
click 143,77, 2
StringLen, partnum, clipboard
;if (%partnum% < 7)
;{
;send 0
;send ^v{enter}
;} else{
;send ^v{enter}
;}
send ^v{enter}
pause
if(i>0)
	{
	winActivate, Microsoft Excel - Q1 '14 Physical Count
	send {right 3}
	clipboard =  
	send ^c
	ClipWait
	StringTrimRight clipboard, clipboard, 2
	winActivate, Microsoft Excel - DanPartsv1
	click 523,194, 2
	send ^v{enter}
	} else{
	winActivate, Microsoft Excel - DanPartsv1
	send {enter}
	}
}