#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent

; GUI Layout
; -------------
Gui, Show, w500 h500, I Love Automation

GUI, Add, Button, x10 y10 w50 g19_to_21, 19-21
GUI, Add, Button, x+10 w50 g20_to_21, 20-21
GUI, Add, Button, x+10 w50 gOnly_20, 20
GUI, Add, Button, x+10 w50 gOnly_21, 21

Gui, Add, Checkbox, x11 vOne, 1
Gui, Add, Checkbox, x+20 vTwo, 2
Gui, Add, Checkbox, x+20 vThree, 3
return

; Labels
; -------------

19_to_21:
	Gui, Submit, NoHide
	If One = 1
		msgbox, you checked the first box
	sleep, 1000
	WinWait, I Love Automation
	WinMinimize
	BtnValue := 1
	LoopVariable := 3
	do_GMK_WL()
	do_GMK_PWL(LoopVariable, BtnValue)
return

20_to_21:
	sleep, 1000
	WinWait, I Love Automation
	WinMinimize
	BtnValue := 2
	LoopVariable := 2
	do_GMK_WL()
	do_GMK_PWL(LoopVariable, BtnValue)
return

Only_20:
	sleep, 1000
	WinWait, I Love Automation
	WinMinimize
	BtnValue := 3
	LoopVariable := 1
	do_GMK_WL()
	do_GMK_PWL(LoopVariable, BtnValue)
return

Only_21:
	sleep, 1000
	WinWait, I Love Automation
	WinMinimize
	BtnValue := 4
	LoopVariable := 1
	do_GMK_WL()
	do_GMK_PWL(LoopVariable, BtnValue)
return

GuiClose:
	ExitApp
return

/*
CoordMode, ToolTip, screen
SetTimer, WatchCursor, 100 
return

WatchCursor:
	CoordMode, mouse, Screen
	MouseGetPos, x_1, y_1, id_1, control_1 ; Coordinates are relative to the desktop (entire screen).

	CoordMode, mouse, Window ; Synonymous with Relative and recommended for clarity.
	MouseGetPos, x_2, y_2, id_2, control_2

	CoordMode, mouse, Client ; Coordinates are relative to the active window's client area
	MouseGetPos, x_3, y_3, id_3, control_3

	ToolTip, Screen :`t`tx %x_1% y %y_1%`nWindow :`tx %x_2% y %y_2%`nClient :`t`tx %x_3% y %y_3%, % A_ScreenWidth-200, % A_ScreenHeight-100
return
*/

; Functions
; -------------

check_boxes(){
	Gui, Submit, NoHide
	If One = 1
		msgbox, you checked the first box
	If Two = 1
		msgbox, you checked the first box
	If Three = 1
		msgbox, you checked the first box
}

do_GMK_WL(){
	sleep, 500
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	CoordMode, mouse, Window
	
	sleep, 1000
	MouseClick, left, 38, 577, 1 ; PRZYPIS 2
	sleep, 1000
	MouseClick, left, 757, 617, 1 ; OPERACJE
	sleep, 1000
	MouseClick, left, 411, 124, 1 ; ZMIEN DATE DO
	sleep, 100
	Send 31-12-2021
	sleep, 1000
	MouseClick, left, 501, 344, 1 ; ZAPISZ (data)
	sleep, 1500
	Send, {Enter} ; OK
	sleep, 1000
	MouseClick, left, 747, 301, 1 ; ZAPISZ
	sleep, 1000
	MouseClick, left, 38, 577, 1 ; PRZYPIS 2
	sleep, 1000
	MouseClick, left, 745, 365, 1 ; KARTOTEKA
	sleep, 1000
	MouseClick, left, 171, 165, 1 ; DOK. FINANSOWE
	sleep, 1000
	MouseClick, left, 650, 230, 1 ; ZAZNACZ1
	sleep, 1000
	MouseClick, left, 604, 377, 1 ; GENERUJ KOREKTE
	sleep, 1000
	MouseClick, left, 415, 416, 1 ; KOREKTA DO ZERA
	sleep, 1000
	MouseClick, left, 484, 470, 1 ; ZATWIERDŹ
	sleep, 1500
	Send, {Enter} ; OK
	sleep, 1000
	MouseClick, left, 651, 249, 1 ; ZAZNACZ2
	sleep, 1000
	MouseClick, left, 604, 377, 1 ; GENERUJ KOREKTE
	sleep, 1000
	MouseClick, left, 415, 416, 1 ; KOREKTA DO ZERA
	sleep, 1000
	MouseClick, left, 484, 470, 1 ; ZATWIERDŹ
	sleep, 1500
	Send, {Enter} ; OK
	sleep, 1000
	MouseClick, left, 777, 93, 1 ; WYJSCIE
	sleep, 1000
	MouseClick, left, 47, 559, 1 ; PRZYPIS 1
	sleep, 1000
	MouseClick, left, 745, 365, 1 ; KARTOTEKA
	sleep, 1000
	MouseClick, left, 171, 165, 1 ; DOK. FINANSOWE
	sleep, 1000
	MouseClick, left, 650, 230, 1 ; ZAZNACZ1
	sleep, 1000
	MouseClick, left, 604, 377, 1 ; GENERUJ KOREKTE
	sleep, 1000
	MouseClick, left, 415, 416, 1 ; KOREKTA DO ZERA
	sleep, 1000
	MouseClick, left, 484, 470, 1 ; ZATWIERDŹ
	sleep, 1500
	Send, {Enter} ; OK
	sleep, 1000
	MouseClick, left, 777, 93, 1 ; WYJSCIE
	sleep, 1000
	MouseClick, left, 777, 93, 1 ; WYJSCIE
	return
}

do_GMK_PWL(LoopVariable, BtnValue){
	sleep, 500
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	CoordMode, mouse, Window
	
	
	sleep, 1000
	MouseClick, left, 136, 432, 1 ; NAZWISKO 2

	sleep, 1000
	MouseClick, left, 47, 559, 1 ; PRZYPIS 1
	sleep, 100
	Send, ^c

	sleep, 1000
	MouseClick, left, 38, 577, 1 ; PRZYPIS 2
	sleep, 1000
	MouseClick, left, 38, 577, 1
	sleep, 100
	Send, ^v
	sleep, 50
	Send, {Tab}{Tab}01-01-2019{Tab}31-12-2021
/*
	sleep, 1000
	MouseClick, left, 747, 301, 1 ; ZAPISZ

	sleep, 1000
	MouseClick, left, 38, 577, 1 ; PRZYPIS 2

	sleep, 1000
	MouseClick, left, 745, 365, 1 ; KARTOTEKA

	sleep, 1000
	MouseClick, left, 697, 293, 1 ; ZAZNACZ
*/
	Loop, %LoopVariable%{
/*		sleep, 1000
		MouseClick, left, 715, 165, 1 ; OPERACJE

		sleep, 1000
		MouseClick, left, 410, 249, 1 ; GENERUJ FAKTURY/DOKUMENTY DLA ZAZN.
		sleep, 1500
		Send, {Enter} ; OK

		sleep, 1000
		MouseClick, left, 537, 119, 1 ; RODZAJ DOK. FINANSOWEGO
		sleep, 1500
		Send, {Enter} ; OK

		sleep, 1000
		MouseClick, left, 539, 140, 1 ; REJESTR VAT
		sleep, 1500
		Send, {Enter} ; OK
		
		sleep, 1000
		MouseClick, left, 363, 322, 1 ; Termin zaplaty/ termin raty/ poczatkowa data okresu
		sleep, 100
*/
		Send, %BtnValue%
		Switch BtnValue{
		case 1: ; 19-21
			if (A_Index = 1)
				Send, 29-02-2020{Tab}29-02-2020{Tab}01-01-2019
			else if (A_Index = 2)
				Send, 30-06-2020{Tab}30-06-2020{Tab}01-01-2020
			else if (A_Index = 3)
				Send, 31-03-2021{Tab}31-03-2021{Tab}01-01-2021
			else ExitApp
		case 2: ; 20-21
			if (A_Index = 1)
				Send, 30-06-2020{Tab}30-06-2020{Tab}01-01-2020
			else if (A_Index = 2)
				Send, 31-03-2021{Tab}31-03-2021{Tab}01-01-2021
			else ExitApp
		case 3: ; 20
			Send, 30-06-2020{Tab}30-06-2020{Tab}01-01-2020
		case 4: ; 21
			Send, 31-03-2021{Tab}31-03-2021{Tab}01-01-2021
		default: ExitApp
		}

		sleep, 1000
		MouseClick, left, 608, 127, 1 ; Generuj Dokument
		sleep, 1500
		Send, {Enter}
	}

	sleep, 1000
	MouseClick, left, 777, 93, 1 ; WYJSCIE

	sleep, 1000
	MouseClick, left, 777, 93, 1 ; WYJSCIE

	MouseClick, left, 51, 413, 1 ; KOD
	Send, ^c
	return
}

do_FAKTUROWANIE(){
	sleep, 500
	if !WinActive("FAKTUROWANIE")
		WinActivate FAKTUROWANIE
	CoordMode, mouse, Window

	sleep, 200
	MouseClick, left, 757, 218, 1 ; Zaznacz1

	sleep, 200
	MouseClick, left, 756, 237, 1 ; Zaznacz2

	sleep, 200
	MouseClick, left, 758, 256, 1 ; Zaznacz3

	sleep, 500
	MouseClick, left, 716, 502, 1 ; Przeslij do NZ

	sleep, 1000
	MouseClick, left, 318, 163, 1 ; JED, RD i TOK - z faktury
	sleep, 2000
	Send, {Enter}
	
	sleep, 1000
	MouseClick, left, 649, 353, 1 ; Kalendarz
	sleep, 1500
	Send, {Enter}
	
	sleep, 1000
	MouseClick, left, 624, 424, 1 ; Zatwierdź
	sleep, 4000
	Send, {Enter}
	sleep, 4000
	Send, {Enter}
	sleep, 4000
	Send, {Enter}
	return
}

clear(){
	sleep, 500
	if WinActive("GOSPODARKA MIENIEM KOMUNALNYM") {	
		MouseClick, left, 92, 224, 1 ; NR UMOWY
		sleep, 50
		Send {F7}

		MouseClick, left, 51, 413, 1 ; KOD
	} else if WinActive("FAKTUROWANIE") {
		MouseClick, left, 97, 217, 1 ; NR REJESTRU
		sleep, 50
		Send {F7}

		MouseClick, left, 94, 414, 1 ; KOD
	} else {
		MsgBox No windows active
	}
	return
}

testing(){
	WinWait, I Love Automation
	WinMinimize
	return
}

; Hotkeys
; -------------

;Test
>+o::
	testing()
return

;GMK WŁ
>+g::
	do_GMK_WL()
return

;GMK PWŁ
>+h::
	;do_GMK_PWL()
return

;FAKTUROWANIE
>+f::
	do_FAKTUROWANIE()
return

;CLEAR
>+c::
	clear()
return

;Ewidencja
\::
	sleep, 250
	MouseClick, left, 748, 388, 1
return

;Konto Kontrahenta
]::
	sleep, 250
	MouseClick, left, 747, 427, 1

	sleep, 1000
	MouseClick, left, 441, 394, 1
return

;Wyjscie z kontrahenta
[::
	sleep, 250
	MouseClick, left, 797, 94, 1
	
	sleep, 250
	MouseClick, left, 236, 399, 1
return

>^x::ExitApp
