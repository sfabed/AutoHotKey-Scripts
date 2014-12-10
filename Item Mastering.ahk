#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
;SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 10
#p::Pause
ESC::ExitApp
ralt::pause
#r::reload

;#c - this checks the description length, note that it can't be over 30 or else Vis will cut it
;^i - this is an automated item master, taking the part number and description. it pauses so you can put in the product and commodity codes. #p unpauses it so it can finish. Auto prints out sticker as well.
;#x - prints out part label
;#d - prints out part description for yellow label. 
;#q - sets part location in warehouse
;^#q - print location 

; Function area!

CpyCell(x,y)
{
clipboard =  
winActivate, Microsoft Excel - Item Master Log
	click %x%,%y% 
	send ^c
ClipWait
StringTrimRight clipboard, clipboard, 2
}

PsteCellinVis:
winActivate, newbox
click 700,200
sendRaw %clipboard%
send {enter}
return


#d::
winActivate, newbox
;send {F4}IQ{enter}1{enter}
send {enter}
	CpyCell(640,195)
	GoSub, PsteCellinVis
sleep 1000
send 6{enter}
;send {f2}
;sleep 1000
pause
winActivate, newbox
send {f4}{enter}
winActivate, Microsoft Excel - Item Master Log
return

#q::
winActivate, newbox
send {f4}IC{enter}2{enter}E{enter}
	CpyCell(640,195)
	GoSub, PsteCellinVis
send ST{enter 2}
	CpyCell(421,195)
	GoSub, PsteCellinVis
send 99{enter}U{enter 6}//{enter}US{enter}
;Print Location Label
	CpyCell(421,195)
/*
winActivate, Location Labels (frmLocation)
click 423,170
SetKeyDelay, 0
send {backspace 10}^v
SetKeyDelay, 50
click 423,241
click 734,164
click 299,508
sleep 500
*/
winActivate, newbox
winActivate, Microsoft Excel - Item Master Log
return

^#q::
;Print Location Label
	CpyCell(421,195)
winActivate, Location Labels (frmLocation)
click 423,170
SetKeyDelay, 0
send {backspace 10}^v
SetKeyDelay, 50
click 423,241
click 734,164
click 299,508
sleep 500
winActivate, Microsoft Excel - Item Master Log
return

#c::
	CpyCell(268,197)
MsgBox % "The length of Description is " . StrLen(clipboard)
return

^i::
SetTable()
winActivate, newbox
send {F4 2}EN{enter}1{enter}e{enter}
	CpyCell(62,195)
prefx=%clipboard%
prod:=prod%prefx%
comm:=comm%prefx%
CpyCell(640,195)
GoSub, PsteCellinVis
CpyCell(239,195)
GoSub, PsteCellinVis
send {enter 3}%prod%{enter}%comm%
send {enter 7}92{enter}1{enter}UA{enter}2{enter}Z{enter}Y{enter}9{enter}N{enter}7{enter}Y{enter}{f4}US{enter}
sleep 500
winActivate, newbox
winActivate, Microsoft Excel - Item Master Log
;send {down}
return

#x::
;loop,
;{
CpyCell(640,195)
winActivate, Part Labels (frmPart)
click 423,170
send ^v
click 734,164
click 299,508
sleep 500
;winActivate, newbox
winActivate, Microsoft Excel - Item Master Log
;send {down}
;}
return

#s::
;loop,
{
winActivate, newbox
send {F4 2}EN{enter}1{enter}S{enter}
	CpyCell(640,195)
	GoSub, PsteCellinVis
	SEND {enter}2{ENTER}DELETED PART{ENTER}US{enter}
	winActivate, Microsoft Excel - Item Master
	send {down}
}

SetTable()
{
global
prod00=MINV
comm00=MAC
prod01=MINV
comm01=MAC
prod02=MINV
comm02=MAC
prod03=FEM
comm03=HAR
prod04=FEM
comm04=SAF
prod05=MINV
comm05=WOD
prod06=FEM
comm06=GLO
prod07=CGS
comm07=CHE
prod08=MINV
comm08=STL
prod09=FEM
comm09=HAR
prod10=MINV
comm10=HAR
prod11=MINV
comm11=WEL
prod12=FEM
comm12=OFF
prod13=CGS
comm13=PAK
prod14=CGS
comm14=PAK
prod15=MINV
comm15=BEA
prod16=MINV
comm16=BEA
prod17=MINV
comm17=HAR
prod18=MINV
comm18=HAR
prod19=MINV
comm19=HAR
prod20=MINV
comm20=MAC
prod21=MINV
comm21=MAC
prod22=MINV
comm22=PLU
prod23=MINV
comm23=PLU
prod24=MINV
comm24=PLU
prod25=MINV
comm25=PUM
prod26=MINV
comm26=PLU
prod27=MINV
comm27=HAR
prod28=MINV
comm28=HAR
prod29=MINV
comm29=HAR
prod30=MINV
comm30=FIL
prod31=MINV
comm31=MAC
prod32=MINV
comm32=BEA
prod33=MINV
comm33=BEA
prod34=MINV
comm34=BEA
prod35=MINV
comm35=BEA
prod36=MINV
comm36=BEA
prod37=MINV
comm37=BEA
prod38=MINV
comm38=BEA
prod39=MINV
comm39=BEA
prod40=MINV
comm40=BEA
prod41=MINV
comm41=BEA
prod42=MINV
comm42=BEA
prod43=MINV
comm43=BEA
prod44=MINV
comm44=BEA
prod45=MINV
comm45=BEA
prod46=MINV
comm46=BEA
prod48=MINV
comm48=ELE
prod49=MINV
comm49=CRA
prod50=MINV
comm50=ELE
prod51=MINV
comm51=ELE
prod52=MINV
comm52=ELE
prod53=MINV
comm53=ELE
prod54=MINV
comm54=ELE
prod55=MINV
comm55=ELE
prod56=MINV
comm56=ELE
prod57=MINV
comm57=ELE
prod58=MINV
comm58=ELE
prod59=MINV
comm59=ELE
prod60=MINV
comm60=ELE
prod61=MINV
comm61=ELE
prod62=MINV
comm62=ELE
prod63=MINV
comm63=ELE
prod64=MINV
comm64=ELE
prod65=MINV
comm65=MOT
prod66=MINV
comm66=ELE
prod67=MINV
comm67=ELE
prod68=MINV
comm68=ELE
prod70=MINV
comm70=ELE
prod71=MINV
comm71=ELE
prod73=MINV
comm73=ELE
prod74=MINV
comm74=CRA
prod75=MINV
comm75=MFG
prod76=CGS
comm76=PAK
prod77=CGS
comm77=PAK
prod78=MFG
comm78=MFG
prod79=MINV
comm79=JAN
prod80=MINV
comm80=SAW
prod81=MINV
comm81=MAC
prod82=MINV
comm82=MAC
prod83=MINV
comm83=CRA
prod84=MINV
comm84=MAC
prod85=MINV
comm85=MAC
prod86=MINV
comm86=MAC
prod87=MINV
comm87=SAW
prod88=MINV
comm88=MAC
prod89=MINV
comm89=REF
prod90=MINV
comm90=MAC
prod92=CGS
comm92=MFG
prod93=FEM
comm93=UNI
prod94=MINV
comm94=HAR
prod95=MINV
comm95=HAR
prod96=MINV
comm96=HAR
prod97=MINV
comm97=HAR
prod98=MINV
comm98=HAR
prod99=NP
comm99=MAC
prod100=NP
comm100=MAC
prod101=MINV
comm101=STL
prod102=NP
comm102=JAN
}
