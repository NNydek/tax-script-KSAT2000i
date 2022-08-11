#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
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

>+l::
	Loop, 3{
		if (A_Index = 1){
			Send, 29-02-2020{Tab}29-02-2020{Tab}01-01-2019
		} else if (A_Index = 2){
			Send, 30-06-2020{Tab}30-06-2020{Tab}01-01-2020
		} else if (A_Index = 3){
			Send, 31-03-2021{Tab}31-03-2021{Tab}01-01-2021
		}
	}
return

;Test
>+o::
	sleep, 500
	Loop, 3{
		sleep, 1000
		;MouseMove 715, 165, 0 ; OPERACJE
		MouseClick, left, 715, 165, 1

		sleep, 1000
		;MouseMove 410, 249, 0 ; GENERUJ FAKTURY/DOKUMENTY DLA ZAZN.
		MouseClick, left, 410, 249, 1
		sleep, 100
		Send, {Enter} ; OK

		sleep, 1000
		;MouseMove 537, 119, 0 ; RODZAJ DOK. FINANSOWEGO
		MouseClick, left, 537, 119, 1
		sleep, 200
		Send, {Enter} ; OK

		sleep, 1000
		;MouseMove 539, 140, 0 ; REJESTR VAT
		MouseClick, left, 539, 140, 1
		sleep, 200
		Send, {Enter} ; OK
		
		sleep, 1000
		;MouseMove 363, 322, 0 ; Termin zaplaty/ termin raty/ poczatkowa data okresu
		MouseClick, left, 363, 322, 1
		sleep, 50
		if (A_Index = 1){
			Send, 29-02-2020{Tab}29-02-2020{Tab}01-01-2019
		} else if (A_Index = 2){
			Send, 30-06-2020{Tab}30-06-2020{Tab}01-01-2020
		} else if (A_Index = 3){
			Send, 31-03-2021{Tab}31-03-2021{Tab}01-01-2021
		}

		sleep, 1000
		;MouseMove 608, 127, 0 ; Generuj Dokument
		MouseClick, left, 608, 127, 1
		sleep, 200
		Send, {Enter}
	}
return

;GMK WŁ
>+g::
	sleep, 500
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	CoordMode, mouse, Window
	sleep, 1000
	;MouseMove 38, 577, 0 ; PRZYPIS 2
	MouseClick, left, 38, 577, 1

	sleep, 1000
	;MouseMove 757, 617, 0 ; OPERACJE
	MouseClick, left, 757, 617, 1
	
	sleep, 1000
	;MouseMove 411, 124, 0 ; ZMIEN DATE DO
	MouseClick, left, 411, 124, 1
	sleep, 100
	Send 31-12-2021
	
	sleep, 1000
	;MouseMove 501, 344, 0 ; ZAPISZ (data)
	MouseClick, left, 501, 344, 1
	sleep, 1500
	Send, {Enter} ; OK

	sleep, 1000
	;MouseMove 747, 301, 0 ; ZAPISZ
	MouseClick, left, 747, 301, 1

	sleep, 1000
	;MouseMove 38, 577, 0 ; PRZYPIS 2
	MouseClick, left, 38, 577, 1

	sleep, 1000
	;MouseMove 745, 365, 0 ; KARTOTEKA
	MouseClick, left, 745, 365, 1	

	sleep, 1000
	;MouseMove 171, 165, 0 ; DOK. FINANSOWE
	MouseClick, left, 171, 165, 1

	sleep, 1000
	;MouseMove 650, 230, 0 ; ZAZNACZ1
	MouseClick, left, 650, 230, 1
	
	sleep, 1000
	;MouseMove 604, 377, 0 ; GENERUJ KOREKTE
	MouseClick, left, 604, 377, 1
	;mozliwe okienko
	
	sleep, 1000
	;MouseMove 415, 416, 0 ; KOREKTA DO ZERA
	MouseClick, left, 415, 416, 1
	;mozliwe okienko
	
	sleep, 1000
	;MouseMove 484, 470, 0 ; ZATWIERDŹ
	MouseClick, left, 484, 470, 1
	sleep, 1500
	Send, {Enter} ; OK
	
	sleep, 1000
	;MouseMove 651, 249, 0 ; ZAZNACZ2
	MouseClick, left, 651, 249, 1

	sleep, 1000
	;MouseMove 604, 377, 0 ; GENERUJ KOREKTE
	MouseClick, left, 604, 377, 1
	;mozliwe okienko

	sleep, 1000
	;MouseMove 415, 416, 0 ; KOREKTA DO ZERA
	MouseClick, left, 415, 416, 1
	;mozliwe okienko

	sleep, 1000
	;MouseMove 484, 470, 0 ; ZATWIERDŹ
	MouseClick, left, 484, 470, 1
	sleep, 1500
	Send, {Enter} ; OK

	sleep, 1000
	;MouseMove 777, 93, 0 ; WYJSCIE
	MouseClick, left, 777, 93, 1

	sleep, 1000
	;MouseMove 47, 559, 0 ; PRZYPIS 1
	MouseClick, left, 47, 559, 1

	sleep, 1000
	;MouseMove 745, 365, 0 ; KARTOTEKA
	MouseClick, left, 745, 365, 1	

	sleep, 1000
	;MouseMove 171, 165, 0 ; DOK. FINANSOWE
	MouseClick, left, 171, 165, 1

	sleep, 1000
	;MouseMove 650, 230, 0 ; ZAZNACZ1
	MouseClick, left, 650, 230, 1
	
	sleep, 1000
	;MouseMove 604, 377, 0 ; GENERUJ KOREKTE
	MouseClick, left, 604, 377, 1
	;mozliwe okienko
	
	sleep, 1000
	;MouseMove 415, 416, 0 ; KOREKTA DO ZERA
	MouseClick, left, 415, 416, 1
	;mozliwe okienko
	
	sleep, 1000
	;MouseMove 484, 470, 0 ; ZATWIERDŹ
	MouseClick, left, 484, 470, 1
	sleep, 1500
	Send, {Enter} ; OK

	sleep, 1000
	;MouseMove 777, 93, 0 ; WYJSCIE
	MouseClick, left, 777, 93, 1

	sleep, 1000
	;MouseMove 777, 93, 0 ; WYJSCIE
	MouseClick, left, 777, 93, 1
return

;GMK PWŁ
>+h::
	sleep, 500
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	CoordMode, mouse, Window
	
	sleep, 1000
	;MouseMove 136, 432, 0 ; NAZWISKO 2
	MouseClick, left, 136, 432, 1

	sleep, 1000
	;MouseMove 47, 559, 0 ; PRZYPIS 1
	MouseClick, left, 47, 559, 1
	sleep, 100
	Send, ^c

	sleep, 1000
	;MouseMove 38, 577, 0 ; PRZYPIS 2
	MouseClick, left, 38, 577, 1
	sleep, 1000
	MouseClick, left, 38, 577, 1
	sleep, 100
	Send, ^v
	sleep, 50
	Send, {Tab}{Tab}01-01-2019{Tab}31-12-2021

	sleep, 1000
	;MouseMove 747, 301, 0 ; ZAPISZ
	MouseClick, left, 747, 301, 1

	sleep, 1000
	;MouseMove 38, 577, 0 ; PRZYPIS 2
	MouseClick, left, 38, 577, 1

	sleep, 1000
	;MouseMove 745, 365, 0 ; KARTOTEKA
	MouseClick, left, 745, 365, 1

	sleep, 1000
	;MouseMove 697, 293, 0 ; ZAZNACZ
	MouseClick, left, 697, 293, 1

	Loop, 3{
		sleep, 1000
		;MouseMove 715, 165, 0 ; OPERACJE
		MouseClick, left, 715, 165, 1

		sleep, 1000
		;MouseMove 410, 249, 0 ; GENERUJ FAKTURY/DOKUMENTY DLA ZAZN.
		MouseClick, left, 410, 249, 1
		sleep, 1500
		Send, {Enter} ; OK

		sleep, 1000
		;MouseMove 537, 119, 0 ; RODZAJ DOK. FINANSOWEGO
		MouseClick, left, 537, 119, 1
		sleep, 1500
		Send, {Enter} ; OK

		sleep, 1000
		;MouseMove 539, 140, 0 ; REJESTR VAT
		MouseClick, left, 539, 140, 1
		sleep, 1500
		Send, {Enter} ; OK
		
		sleep, 1000
		;MouseMove 363, 322, 0 ; Termin zaplaty/ termin raty/ poczatkowa data okresu
		MouseClick, left, 363, 322, 1
		sleep, 100
		if (A_Index = 1){
			Send, 29-02-2020{Tab}29-02-2020{Tab}01-01-2019
		} else if (A_Index = 2){
			Send, 30-06-2020{Tab}30-06-2020{Tab}01-01-2020
		} else if (A_Index = 3){
			Send, 31-03-2021{Tab}31-03-2021{Tab}01-01-2021
		}

		sleep, 1000
		;MouseMove 608, 127, 0 ; Generuj Dokument
		MouseClick, left, 608, 127, 1
		sleep, 1500
		Send, {Enter}
	}

	sleep, 1000
	;MouseMove 777, 93, 0 ; WYJSCIE
	MouseClick, left, 777, 93, 1

	sleep, 1000
	;MouseMove 777, 93, 0 ; WYJSCIE
	MouseClick, left, 777, 93, 1

	;MouseMove 51, 413, 0 ; KOD
	MouseClick, left, 51, 413, 1
	Send, ^c
return

;FAKTUROWANIE
>+f::
	sleep, 500
	if !WinActive("FAKTUROWANIE")
		WinActivate FAKTUROWANIE
	CoordMode, mouse, Window

	sleep, 200
	;MouseMove 757, 218, 0 ; Zaznacz1
	MouseClick, left, 757, 218, 1

	sleep, 200
	;MouseMove 756, 237, 0 ; Zaznacz2
	MouseClick, left, 756, 237, 1

	sleep, 200
	;MouseMove 758, 256, 0 ; Zaznacz3
	MouseClick, left, 758, 256, 1

	sleep, 500
	;MouseMove 716, 502, 0 ; Przeslij do NZ
	MouseClick, left, 716, 502, 1

	sleep, 1000
	;MouseMove 318, 163, 0 ; JED, RD i TOK - z faktury
	MouseClick, left, 318, 163, 1
	sleep, 2000
	Send, {Enter}
	
	sleep, 1000
	;MouseMove 649, 353, 0 ; Kalendarz
	MouseClick, left, 649, 353, 1
	sleep, 1500
	Send, {Enter}
	
	sleep, 1000
	;MouseMove 624, 424, 0 ; Zatwierdź
	MouseClick, left, 624, 424, 1
	sleep, 4000
	Send, {Enter}
	sleep, 4000
	Send, {Enter}
	sleep, 4000
	Send, {Enter}
return

;CLEAR
>+c::
	sleep, 500
	if WinActive("GOSPODARKA MIENIEM KOMUNALNYM") {	
		;MouseMove 92, 224, 0 ; NR UMOWY
		MouseClick, left, 92, 224, 1
		sleep, 50
		Send {F7}

		;MouseMove 51, 413, 0 ; KOD
		MouseClick, left, 51, 413, 1
	} else if WinActive("FAKTUROWANIE") {
		;MouseMove 97, 217, 0 ; NR REJESTRU
		MouseClick, left, 97, 217, 1
		sleep, 50
		Send {F7}

		;MouseMove 94, 414, 0 ; KOD
		MouseClick, left, 94, 414, 1
	} else {
		MsgBox No windows active
	}
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
