#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetKeyDelay, 100
esc::exitapp
#p::pause
ralt::pause
#r::reload

CpyCell(x,y)
{
clipboard =  
winActivate, Microsoft Excel - Item Master Log
	click %x%,%y%
	send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
}

^+`::
loop,
{
click (40)
}