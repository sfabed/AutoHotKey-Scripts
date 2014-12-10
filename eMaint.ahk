#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetKeyDelay, 250
esc::exitapp
ralt::pause
#r::reload

waitForPageLoad(x, y, color)
{
;Computer1(75, 28, "0x68A4FC")
;Computer2(30, 32, "0x4992F9")
Loop
{
Loaded =
PixelGetColor, Loaded, %x%, %y%
if Loaded = %color%
break
}

Loop
{
Loaded =
PixelGetColor, Loaded, %x%, %y%
if Loaded = %color%
break
}

}

CpyCell(x,y)
{
SLEEP 500
click 90,220

clipboard =  
SLEEP 1000
click %x%,%y%,2
send ^c
ClipWait
if clipboard is not alnum
{
	click 99,198,2
	clipboard =
	send ^c
	if (clipboard = "DELETION")
	{
	click 101,239
	waitForPageLoad(30, 32, "0x4992F9")
	CpyCell(733,571)
	}else
{
	CpyCell(733,571)
	Msgbox, check this shit
	pause
}
}
}

;waitForPageLoad(75, 28, "0x68A4FC")
;waitForPageLoad(30, 32, "0x4992F9")

Insert::
loop,
{
	winActivate, Inventory - Google Chrome
	CpyCell(733,571)
	partnum = %clipboard%
	;pause
	click 176,451,2
	clipboard =
	send ^c
	;msgbox, %clipboard%
	if clipboard =
	{
		;msgbox, "flag 1"
		sendraw %partnum%
		click 68,222
		;waitForPageLoad(75, 28, "0x68A4FC")
		 waitForPageLoad(30, 32, "0x4992F9")
		sleep 500
		click 290,220
		;waitForPageLoad(75, 28, "0x68A4FC")
		waitForPageLoad(30, 32, "0x4992F9")
		sleep 1000
		
	}
		else if clipboard <>
		{
		;msgbox, "flag 2"
			click 68,222
			;waitForPageLoad(75, 28, "0x68A4FC")
			waitForPageLoad(30, 32, "0x4992F9")
			sleep 500
			click 290,220
			;waitForPageLoad(75, 28, "0x68A4FC")
			waitForPageLoad(30, 32, "0x4992F9")
			sleep 1000
		}
}

/*

loop,
{
winActivate, Inventory - Google Chrome
CpyCell(733,571)
partnum = %clipboard%
;msgbox, %partnum%
click 153,222
waitForPageLoad(75, 28, "0x68A4FC")
MouseMove 44,222
sleep 500
click 62,263
waitForPageLoad(75, 28, "0x68A4FC")
sleep 1000
click 198,451
sendraw %partnum%
click 190,591
send y
click 190,575
send t
click 190,492
;pause
click 68,222
waitForPageLoad(75, 28, "0x68A4FC")
sleep 500
CpyCell(733,571)
click 153,222
waitForPageLoad(75, 28, "0x68A4FC")
sleep 500
entrynum = %clipboard%
winActivate, Microsoft Excel - Emaint Changes
click 60,195
send %partnum%{tab}%entrynum%{down}
winActivate, Inventory - Google Chrome
click 250,220
send %partnum%{enter}
waitForPageLoad(75, 28, "0x68A4FC")
click 290,220
waitForPageLoad(75, 28, "0x68A4FC")
sleep 1000

*/