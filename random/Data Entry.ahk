SetKeyDelay, 50
esc::exitapp
ralt::pause
#r::reload

CpyCell(x,y)
{
clipboard =  
winActivate, Microsoft Excel - Q4
	click %x%,%y%
	sleep 500
	click %x%,%y%
	send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
}

PasteCell(x,y)
{
winActivate, Part Trasactions (frmPartTrans)
	click %x%,%y%
	send %clipboard%
}

Insert::
{
LOOP,
{
	winActivate, Part Trasactions (frmPartTrans)
	click 493,173
	send {BS 20}
	click 493,220
	send {BS 20}
	click 445,301
	send {BS 20}
	send 0
	sleep 50

Loop,
{
	;Part Number
	CPYCELL(83,197)
	PASTECELL(493,173)
	
	;Location
	CPYCELL(429,197)
	PASTECELL(427,220)
	
	;OnHand
	CPYCELL(483,197)
	PASTECELL(445,301)
			
	;Update
	winActivate, Part Trasactions (frmPartTrans)
	click 499,586
	sleep 100
	Ifwinexist, CALL_STOREROOM_PROC
	{
	;sleep 1000
	send {enter}
	winActivate, Microsoft Excel - Q4
	sleep 50
	click 673,197
	sleep 50
	send U
	send {Down}
	sleep 50
	break
	;pause
	return
	}else
	{
	;msgbox, AOK
	;Mark,Ready Next Line
		sleep 50
		winActivate, Microsoft Excel - Q4
		sleep 50
		click 673,197
		sleep 50
		send X
		send {Down}
		sleep 50
	}
	
}
}
}