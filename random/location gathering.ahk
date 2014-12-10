#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetKeyDelay, 200
esc::exitapp
#p::pause
ralt::pause
#r::reload

enter::
	send {down}
	winActivate, Microsoft Excel - Book1
	click 67,198
	sleep 50
	send ^c
	sleep 50
	winActivate, Part Locations Lookup (frmPartLocations)
	click 474,194
	send {backspace 20}
	send ^v{enter}
	winActivate, Microsoft Excel - Book1
	click 131,198
	return
	

