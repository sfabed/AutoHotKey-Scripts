#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetKeyDelay, 50
esc::exitapp
ralt::pause
#r::reload

/*
#d: Purchase History and Vendor Information
	Dep: Parts to Order & Vis
^d: Check for Open Reqs
	Dep: Parts to Order & Vis
Insert: Get Part Descriptions
	Dep: Parts to Order & eMaint
^r:Creates Reqs and Prints out two copies.
	Dep: Parts to Order (all info filled) & Req Screen 
*/


CpyCell(x,y)
{
clipboard =  
winActivate, Microsoft Excel - ScrapSheet
	click %x%,%y%
	send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
}

PsteCellinVis:
winActivate, newbox
click 932,200
send %clipboard%
send {enter}
return


; Purchase History (Takes Vendor number and checks PH after specified date (currently 01/01/08))
#d::
loop,
{
winActivate, newbox
send {enter}
CPYCELL(129,200)
GoSub, PsteCellinVis
sleep 500

send 33{enter 3}
;send 010108{enter 2}
winActivate, Microsoft Excel - ScrapSheet
click 179,200
pause
winActivate, Microsoft Excel - ScrapSheet
click 520,200
pause

;winActivate, newbox
;send 6{enter}
;winActivate, Microsoft Excel - ScrapSheet
;click 570,200
;pause
winActivate, newbox
send {f4}{enter}
winActivate, Microsoft Excel - ScrapSheet
send {down}
}

;Uses Emaint to get the descriptions of part number
Insert::
loop,
{
CPYCELL(129,200)
winActivate, eMaint X3 - Google Chrome
click 35,295
click 916,280
sendRaw %clipboard%
click 29,281
sleep  1500
click 494,338, 3
clipboard =  
send ^c
ClipWait
winActivate, Microsoft Excel - ScrapSheet 
	click 340,196
	send {Raw}%clipboard% 
	
winActivate, eMaint X3 - Google Chrome
click 644,337, 3
clipboard =  
send ^c
ClipWait
winActivate, Microsoft Excel - ScrapSheet 
	click 666,196
	send {Raw}%clipboard%
	send {down}
}



;Checks for Open Reqs
^d::
loop,
{
winActivate, newbox
send {enter}
CPYCELL(129,200)
	GoSub, PsteCellinVis
sleep 500
send 31{enter}
send O{enter}
pause
winActivate, newbox
send {f4}{enter}
winActivate, Microsoft Excel - ScrapSheet
send {down}
}

;Creates Reqs
#w::
SetKeyDelay, 200
winActivate, req_header (req_entry.vbp)
click 525,184
send STOCK ITEMS
click 908,218
send {enter}
clipboard =
click 514,63, 2
send ^c
ClipWait
reqnum = %clipboard%
click 489,319
send 5400350
loop,
{
	CpyCell(67,195)
	IfInString, clipboard, END 
	{
		MsgBox, END THE REQ MAN!!!
		SetKeyDelay, 100
		winActivate, req_header (req_entry.vbp)
		winWait, req_header
		click 654,459
		WinWait, Impromptu
		winActivate, Impromptu
		click 128,62
		sleep 1000
		WinWait, Impromptu
		winActivate, Impromptu
		click 128,62
		winActivate, req_header (req_entry.vbp)
		winWait, req_header
		send {f4}
		return
		} else {
			winActivate, req_header (req_entry.vbp)
			click 164,311
			send ^v{tab}
			CpyCell(175,195)
			winActivate, req_header (req_entry.vbp)
			click  329,312
			send ^v
			CpyCell(443,200)
			click 504,200
			send %reqnum%
			send {down}
			winActivate, req_header (req_entry.vbp)
			click 845,299
			send ^v{f2}
			sleep 500
			send {enter}
			sleep 500
	}
}
return

^b::
;used to update the o/h qty of parts ordered
loop,
{
winActivate Microsoft Excel - ScrapSheet
click 62,87
clipboard = 
send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
partnum=%clipboard%
;msgbox, %partnum%
click 130,87
clipboard = 
send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
ohqty=%clipboard%
;msgbox, %ohqty%
send {down}
winActivate, Part Trasactions
click 422,179
send %partnum%
click 420,299
send %ohqty%
click 450,588
sleep 500
ifWinExist, CALL_STOREROOM_PROC
{
pause
}
sleep 500
;pause
}



