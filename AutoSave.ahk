#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#p::Pause
ESC::ExitApp
ralt::pause
#r::reload
SetKeyDelay, 250


CpyCell(x,y)
{
clipboard =  
winActivate, Microsoft Excel - ScrapSheet
	click %x%,%y%
	send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
}

#x::


loop
{
CpyCell(66,198)
prefx=%clipboard%
sleep 500
CpyCell(129,198)
product=%clipboard%
sleep 500
CpyCell(191,198)
commad=%clipboard%
sleep 500
;listvars
;pause
winActivate, Microsoft Excel - ScrapSheet
send {down}
winActivate, *new  2 - Notepad++
send prod%prefx%=%product%
sleep 50
send {enter 2}
sleep 50
send comm%prefx%=%commad%
sleep 50
send {enter 2}
sleep 50
}