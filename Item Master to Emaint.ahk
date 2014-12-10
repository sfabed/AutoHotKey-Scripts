#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 10
#p::Pause
ESC::ExitApp
ralt::pause
#r::reload
/*
Excel Coor:
All y coordinates are 200.
X's are as follows: Description(164),Location(420),Part#(657)
Computer1(75, 28, "0x68A4FC")
Computer2(30, 32, "0x4992F9")
*/
SetVar()
{
global
;spreadsheet = Microsoft Excel - Item Master
spreadsheet = Microsoft Excel - ScrapSheet
xp = 29
yp = 30
color = 0xE6F0FD
}

CpyCell(x,y,spreadsheet)
{
clipboard =  
winActivate, %spreadsheet%
;msgbox, %spreadsheet%
	click %x%,%y%,1
	send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
sleep 500
if clipboard = 
	{
	msgbox, Please Check Me! "%clipboard%"
	pause
	}
}

PsteCellinVis:
winActivate, newbox
click 700,200
sendRaw %clipboard%
send {enter}
return

waitForPageLoad()
{
loop
	{
	Loaded =
	PixelGetColor, Loaded, %xp%, %yp%
	if Loaded = %color%
	break
	}
sleep 50
loop
	{
	Loaded =
	PixelGetColor, Loaded, %x%, %y%
	if Loaded = %color%
	break
	}
}

Insert::
{
winActivate, ahk_class Chrome_WidgetWin_1
send ^n
sleep 500
send #{left}
sleep 500
click 338,86
sleep 1000
click 625,86
sleep 1000
/*
loop,
{
*/
SetVar()
sleep 500
CpyCell(164,200,spreadsheet)
winActivate, Inventory - Google Chrome
click 288,394
send ^v
CpyCell(658,200,spreadsheet)
winActivate, Inventory - Google Chrome
click 205,452
send ^v
click 759,571
send {backspace 8}^v
CpyCell(420,200,spreadsheet)
winActivate, Inventory - Google Chrome
click 184,377
send ^v
;}
}