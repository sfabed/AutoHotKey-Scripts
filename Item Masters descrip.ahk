#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetKeyDelay, 50
esc::exitapp
ralt::pause
#r::reload

CpyCell(x,y)
{
clipboard =  
;winActivate, Microsoft Excel - Item Master
winActivate, Microsoft Excel - Scrap
	click %x%,%y%
	send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
}

PsteCellinVis:
winActivate, newbox
click 932,200
sendRaw %clipboard%
send {enter}
return

#a::
loop,
{
winActivate, newbox
send {F4 2}EN{enter}1{enter}s{enter}
	CpyCell(648,195)
	GoSub, PsteCellinVis
	;send {enter}6{enter 1}CGS{enter}7{enter}CHE{enter}U{enter}
	send {enter}94{enter}
	sleep 500
	send {enter}A{enter 2}
	;pause
	send REF PO#:9007730
	;pause
	;CpyCell(136,195)
	;GoSub, PsteCellinVis
	send {enter}//{enter}UR{enter}
	winActivate, Microsoft Excel - Item
send {down}
}


#ss::
loop,
{
winActivate, newbox
send {f4 2}IC{enter}2{enter}E{enter}
	CpyCell(640,195)
	GoSub, PsteCellinVis
send ST{enter 2}
	clipboard=ANNEX
	GoSub, PsteCellinVis
send 99{enter}U{enter 6}//{enter}US{enter}
winActivate, newbox
winActivate, Microsoft Excel - Item Master Log
send {down}
}

#s::
loop,
{
winActivate, newbox
send {F4 2}EN{enter}1{enter}s{enter}
	CpyCell(129,200)
	GoSub, PsteCellinVis
	send {enter}94{enter 2}
	sleep 10
	send A{enter 2}
	;msgbox, hey
send END USE:DIE CAUSTIC
send {enter 3}
send CLEANING SYSTEM
;send {enter 3}
;send REF PO:9006946
winActivate, Microsoft Excel - Scrap
send {down}
winActivate, newbox
send {enter}//{enter}UR{enter}
winActivate, newbox
}
