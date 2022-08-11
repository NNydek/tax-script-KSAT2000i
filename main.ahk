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
GUI, Add, Button, x+10 w50 gOnly_21, 21
GUI, Add, Button, x+260 w50 gWatch_Cursor, Watch Cursor

Gui, Add, Checkbox, x11 vFirst_Surname_Pos, Nazwisko Poz. 1
Gui, Add, Checkbox, x+17 vSecond_Surname_Pos, Nazwisko Poz. 2
Gui, Add, Checkbox, x+17 vThird_Surname_Pos, Nazwisko Poz. 3

Gui, Add, Checkbox, x11 vFirst_Footnote, Przypis pierwszy
Gui, Add, Checkbox, x+20 vSecond_Footnote, Przypis drugi

Gui, Add, Checkbox, x11 vGWL_1, WL Poz. 1
Gui, Add, Checkbox, x+20 vGWL_2, WL Poz. 2
Gui, Add, Checkbox, x+20 vGWL_3, WL Poz. 3
Gui, Add, Checkbox, x+20 vGWL_4, WL Poz. 4
Gui, Add, Checkbox, x+20 vGWL_5, WL Poz. 5

Gui, Add, Checkbox, x11 vGWL_6, WL Poz. 6
Gui, Add, Checkbox, x+20 vGWL_7, WL Poz. 7
Gui, Add, Checkbox, x+20 vGWL_8, WL Poz. 8
Gui, Add, Checkbox, x+20 vGWL_9, WL Poz. 9
Gui, Add, Checkbox, x+20 vGWL_10, WL Poz. 10

Gui, Add, Checkbox, x11 vGPWL_1, PWL Poz. 1
Gui, Add, Checkbox, x+13 vGPWL_2, PWL Poz. 2
Gui, Add, Checkbox, x+13 vGPWL_3, PWL Poz. 3
Gui, Add, Checkbox, x+13 vGPWL_4, PWL Poz. 4
Gui, Add, Checkbox, x+13 vGPWL_5, PWL Poz. 5

Gui, Add, Checkbox, x11 vGPWL_6, PWL Poz. 6
Gui, Add, Checkbox, x+13 vGPWL_7, PWL Poz. 7
Gui, Add, Checkbox, x+13 vGPWL_8, PWL Poz. 8
Gui, Add, Checkbox, x+13 vGPWL_9, PWL Poz. 9
Gui, Add, Checkbox, x+13 vGPWL_10, PWL Poz. 10
return


CoordMode, ToolTip, screen
SetTimer, Watch_Cursor, 100 
return

; Labels
; -------------

19_to_21:
	Gui, Submit, NoHide
	SurnamePos := check_surname_boxes(First_Surname_Pos, Second_Surname_Pos, Third_Surname_Pos)
	FootnoteNum := check_footnote_boxes(First_Footnote, Second_Footnote)
	WLPos := check_gwl_boxes(GWL_1, GWL_2, GWL_3, GWL_4, GWL_5, GWL_6, GWL_7, GWL_8, GWL_9, GWL_10)
	PWLPos := check_gpwl_boxes(GPWL_1, GPWL_2, GPWL_3, GPWL_4, GPWL_5)
	if (!SurnamePos || !FootnoteNum || !WLPos || !PWLPos)
		return 
	sleep, 1000
	WinWait, I Love Automation
	WinMinimize
	BtnValue := 1
	LoopVariable := 3
	do_GMK_WL(FootnoteNum, WLPos, BtnValue)
	do_GMK_PWL(LoopVariable, BtnValue, SurnamePos, PWLPos)
	do_FAKTUROWANIE()
	update_excel()
	paste_new_user_code()
return

20_to_21:
	Gui, Submit, NoHide
	SurnamePos := check_surname_boxes(First_Surname_Pos, Second_Surname_Pos, Third_Surname_Pos)
	FootnoteNum := check_footnote_boxes(First_Footnote, Second_Footnote)
	WLPos := check_gwl_boxes(GWL_1, GWL_2, GWL_3, GWL_4, GWL_5, GWL_6, GWL_7, GWL_8, GWL_9, GWL_10)
	PWLPos := check_gpwl_boxes(GPWL_1, GPWL_2, GPWL_3, GPWL_4, GPWL_5)
	if (!SurnamePos || !FootnoteNum || !WLPos || !PWLPos)
		return 
	sleep, 1000
	WinWait, I Love Automation
	WinMinimize
	BtnValue := 2
	LoopVariable := 2
	do_GMK_WL(FootnoteNum, WLPos, BtnValue)
	do_GMK_PWL(LoopVariable, BtnValue, SurnamePos, PWLPos)
	do_FAKTUROWANIE()
	update_excel()
	paste_new_user_code()
return

Only_21:
	Gui, Submit, NoHide
	SurnamePos := check_surname_boxes(First_Surname_Pos, Second_Surname_Pos, Third_Surname_Pos)
	FootnoteNum := check_footnote_boxes(First_Footnote, Second_Footnote)
	WLPos := check_gwl_boxes(GWL_1, GWL_2, GWL_3, GWL_4, GWL_5, GWL_6, GWL_7, GWL_8, GWL_9, GWL_10)
	PWLPos := check_gpwl_boxes(GPWL_1, GPWL_2, GPWL_3, GPWL_4, GPWL_5)
	if (!SurnamePos || !FootnoteNum || !WLPos || !PWLPos)
		return 
	sleep, 1000
	WinWait, I Love Automation
	WinMinimize
	BtnValue := 3
	LoopVariable := 1
	do_GMK_WL(FootnoteNum, WLPos, BtnValue)
	do_GMK_PWL(LoopVariable, BtnValue, SurnamePos, PWLPos)
	do_FAKTUROWANIE()
	update_excel()
	paste_new_user_code()
return

Watch_Cursor:
	while(1){
		CoordMode, mouse, Screen
		MouseGetPos, x_1, y_1, id_1, control_1 ; Coordinates are relative to the desktop (entire screen).

		CoordMode, mouse, Window ; Synonymous with Relative and recommended for clarity.
		MouseGetPos, x_2, y_2, id_2, control_2

		CoordMode, mouse, Client ; Coordinates are relative to the active window's client area
		MouseGetPos, x_3, y_3, id_3, control_3

		ToolTip, Screen :`t`tx %x_1% y %y_1%`nWindow :`tx %x_2% y %y_2%`nClient :`t`tx %x_3% y %y_3%, % A_ScreenWidth-200, % A_ScreenHeight-100
	}
return

GuiClose:
	ExitApp
return


; Functions
; -------------

check_gwl_boxes(a, b, c, d, e, f, g, h, i, j){
	if (!a && !b && !c && !d && !e && !f && !g && !h && !i && !j){
		msgbox, Select at least 1 checkbox (wl)
		return 0
	} else if (a + b + c + d + e + f + g + h + i + j > 1){
		msgbox, Too many arguments (wl)
		return 0
	} else if (a)
		return 1
	else if (b)
		return 2
	else if (c)
		return 3
	else if (d)
		return 4
	else if (e)
		return 5
	else if (f)
		return 6
	else if (g)
		return 7
	else if (h)
		return 8
	else if (i)
		return 9
	else if (j)
		return 10
	else
		msgbox, Error (check_gpwl_boxes else)
	return
}

check_gpwl_boxes(a, b, c, d, e){
	if (!a && !b && !c && !d && !e){
		msgbox, Select at least 1 checkbox (pwl)
		return 0
	} else if (a + b + c + d + e > 1){
		msgbox, Too many arguments (pwl)
		return 0
	} else if (a)
		return 1
	else if (b)
		return 2
	else if (c)
		return 3
	else if (d)
		return 4
	else if (e)
		return 5
	else
		msgbox, Error (check_gpwl_boxes else)
	return
}

check_footnote_boxes(first, second){
	if (!first && !second){
		msgbox, Select at least 1 checkbox (footnote)
		return 0
	} else if (first + second > 1){
		msgbox, Too many arguments (footnote)
		return 0
	} else if (first = 1)
		return 1
	else if (second = 1)
		return 2
	else
		msgbox, Error (check_footnote_boxes else)
	return
}

check_surname_boxes(1_pos, 2_pos, 3_pos){
	if (!1_pos && !2_pos && !3_pos){
		msgbox, Select at least 1 checkbox (surname)
		return 0
	} else if (1_pos + 2_pos + 3_pos > 1){
		msgbox, Too many arguments (surname)
		return 0
	}
	else if (1_pos = 1){
		return 1
	}
	else if (2_pos = 1){
		return 2
	}
	else if (3_pos = 1){
		msgbox, WiP
		return 3
	} else
		msgbox, Error (check_surname_boxes else)
	return
}

do_GMK_WL(FootnoteNum, WLPos, BtnValue){
	sleep, 500
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	CoordMode, mouse, Window
	
	sleep, 1000
	MouseClick, left, 58, 222, 1 ; NUMER UMOWY ZAZN. 1
	Sleep, 1000
	Switch WLPos{
	case 1:
	case 2:
		Send, {Down}
	case 3:
		Send, {Down 2}
	case 4:
		Send, {Down 3}
	case 5:
		Send, {Down 4}
	case 6:
		Send, {Down 5}
	case 7:
		Send, {Down 6}
	case 8:
		Send, {Down 7}
	case 9:
		Send, {Down 8}
	case 10:
		Send, {Down 9}
	default:
		msgbox, Error (do_GMK_WL Switch WLPos)
	}

	sleep, 1000
	MouseClick, left, 748, 385, 1 ; EWIDENCJA
	
	Switch FootnoteNum{
	case 1:
		sleep, 1000
		MouseClick, left, 38, 557, 1 ; PRZYPIS 1		
	case 2:
		sleep, 1000
		MouseClick, left, 38, 577, 1 ; PRZYPIS 2
	default:
		msgbox, Error (do_GMK_WL if(FootnoteNum) else)
	} 
	sleep, 1000
	MouseClick, left, 757, 617, 1 ; OPERACJE
	sleep, 1000
	MouseClick, left, 411, 124, 1 ; ZMIEN DATE DO
	sleep, 100
	Send 31-12-2021
	sleep, 1000
	MouseClick, left, 501, 344, 1 ; ZAPISZ (data)
	sleep, 2000
	Send, {Enter} ; OK
	sleep, 1000
	MouseClick, left, 747, 301, 1 ; ZAPISZ
	if (BtnValue <= 2){
		sleep, 1000
		MouseClick, left, 38, 577, 1 ; PRZYPIS 2
		sleep, 1000
		MouseClick, left, 745, 365, 1 ; KARTOTEKA
		sleep, 1000
		MouseClick, left, 171, 165, 1 ; DOK. FINANSOWE
		sleep, 1000
		MouseClick, left, 651, 249, 1 ; ZAZNACZ2
		sleep, 1000
		MouseClick, left, 604, 377, 1 ; GENERUJ KOREKTE
		sleep, 1000
		MouseClick, left, 415, 416, 1 ; KOREKTA DO ZERA
		sleep, 1000
		MouseClick, left, 484, 470, 1 ; ZATWIERDŹ
		sleep, 1000
		Send, {Enter} ; OK
		sleep, 1000
		MouseClick, left, 650, 230, 1 ; ZAZNACZ1
		sleep, 1000
		MouseClick, left, 604, 377, 1 ; GENERUJ KOREKTE
		sleep, 1000
		MouseClick, left, 415, 416, 1 ; KOREKTA DO ZERA
		sleep, 1000
		MouseClick, left, 484, 470, 1 ; ZATWIERDŹ
		sleep, 2000
		Send, {Enter} ; OK
		sleep, 1000
		MouseClick, left, 777, 93, 1 ; WYJSCIE
	} 
	if (BtnValue = 1){
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
		sleep, 2000
		Send, {Enter} ; OK
		sleep, 1000
		MouseClick, left, 777, 93, 1 ; WYJSCIE
	} else if (BtnValue = 3){
		Switch FootnoteNum{
		case 1:
			sleep, 1000
			MouseClick, left, 38, 557, 1 ; PRZYPIS 1
		case 2:
			sleep, 1000
			MouseClick, left, 38, 577, 1 ; PRZYPIS 2			
		}
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
		sleep, 2000
		Send, {Enter} ; OK
		sleep, 1000
		MouseClick, left, 777, 93, 1 ; WYJSCIE
	} else if (BtnValue != 2)
		msgbox, Error (do_GMK_WL if(FootnoteNum2) else)
	sleep, 1000
	MouseClick, left, 777, 93, 1 ; WYJSCIE
	return
}

do_GMK_PWL(LoopVariable, BtnValue, SurnamePos, PWLPos){
	sleep, 500
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	CoordMode, mouse, Window
	
	sleep, 1000	
	MouseClick, left, 58, 222, 1 ; NUMER UMOWY ZAZN. 1
	sleep, 1000
	Send, {Up 9}
	Sleep, 350
	Send, {Enter}
	sleep, 1000
	Switch PWLPos{
	case 1:
	case 2:
		Send, {Down}
	case 3:
		Send, {Down 2}
	case 4:
		Send, {Down 3}
	case 5:
		Send, {Down 4}
	default:
		msgbox, Error (do_GMK_WL Switch WLPos)
	}
	sleep, 1000
	MouseClick, left, 748, 385, 1 ; EWIDENCJA	

	if(SurnamePos = 2){
		sleep, 1000
		MouseClick, left, 136, 432, 1 ; NAZWISKO 2
		yPos := 433
	} else if (SurnamePos = 3){
		sleep, 1000
		MouseClick, left, 136, 452, 1 ; NAZWISKO 3
		yPos := 452
	} else if (SurnamePos = 1){
		yPos := 413
	}
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
	sleep, 100
	Switch BtnValue{
	case 1: ; 19-21
		Send, {Tab}{Tab}01-01-2019{Tab}31-12-2021
	case 2: ; 20-21
		Send, {Tab}{Tab}01-01-2020{Tab}31-12-2021
	case 3: ; 21
		Send, {Tab}{Tab}01-01-2021{Tab}31-12-2021
	}
	
	sleep, 1000
	MouseClick, left, 747, 301, 1 ; ZAPISZ

	sleep, 1000
	MouseClick, left, 38, 577, 1 ; PRZYPIS 2

	sleep, 1000
	MouseClick, left, 745, 365, 1 ; KARTOTEKA

	sleep, 1000
	MouseClick, left, 697, 293, 1 ; ZAZNACZ

	Loop, %LoopVariable%{
		sleep, 1000
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

		Switch BtnValue{
		case 1: ; 19-21
			if (A_Index = 1)
				Send, 29-02-2020{Tab}29-02-2020{Tab}01-01-2019
			else if (A_Index = 2)
				Send, 30-06-2020{Tab}30-06-2020{Tab}01-01-2020
			else if (A_Index = 3)
				Send, 31-03-2021{Tab}31-03-2021{Tab}01-01-2021
			else msgbox, Error (do_GMK_PWL Case1)
		case 2: ; 20-21
			if (A_Index = 1)
				Send, 30-06-2020{Tab}30-06-2020{Tab}01-01-2020
			else if (A_Index = 2)
				Send, 31-03-2021{Tab}31-03-2021{Tab}01-01-2021
			else msgbox, Error (do_GMK_PWL Case2)
		case 3: ; 21
			Send, 31-03-2021{Tab}31-03-2021{Tab}01-01-2021
		default: msgbox, Error (do_GMK_PWL default)
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

	sleep, 1000
	MouseClick, left, 51, %yPos%, 1 ; KOD
	Send, ^c

	sleep, 1000
	MouseClick, left, 92, 224, 1 ; NR UMOWY
	sleep, 50
	Send {F7}
return
}

do_FAKTUROWANIE(){
	sleep, 500
	if !WinActive("FAKTUROWANIE")
		WinActivate FAKTUROWANIE
	CoordMode, mouse, Window

	sleep, 1000
	MouseClick, left, 94, 414, 1 ; KOD
	sleep, 200
	Send, ^v
	sleep, 1000
	Send, {F8}
	
	sleep, 2000
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
	sleep, 3000
	Send, {Enter}
	
	sleep, 4000
	MouseClick, left, 97, 217, 1 ; NR REJESTRU
	sleep, 200
	Send {F7}
return
}

update_excel(){
	oExcel := ComObjCreate("Excel.Application") ; create Excel Application object
	oExcel.Workbooks.Add ; create a new workbook (oWorkbook := oExcel.Workbooks.Add)
	
	FilePath := "d:\Users\umpawy01\Desktop\Bruttowanie.xlsx" ; example path
	oWorkbook := ComObjGet(FilePath) ; access Workbook object
	
	column := 65 ; ASCII - A
	row := 2
	
	for xlCell in oWorkbook.Sheets(1).UsedRange.Cells{
		c1 := Chr(65) ; ASCII to String - A
		c2 := Chr(67) ; C
		cell = %c1%%row%:%c2%%row% ; e.g. A2
		if(oWorkbook.Sheets(1).Range(cell).Interior.ColorIndex = 43 || oWorkbook.Sheets(1).Range(cell).Interior.ColorIndex = 48){
			row++
			continue
		}
		oWorkbook.Sheets(1).Range(cell).Interior.ColorIndex := 43 ; green
		sleep, 50
		row++
		nextCode = %c1%%row%
		clipboard := Format("{:.0f}", oWorkbook.Sheets(1).Range(nextCode).Value)
		Send, ^c
		break
	}
	return
}

paste_new_user_code(){
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	MouseClick, left, 51, 413, 1 ; KOD
	sleep, 1000
	Send, ^v
	sleep, 500
	Send, {F8}
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
	oExcel := ComObjCreate("Excel.Application") ; create Excel Application object
	oExcel.Workbooks.Add ; create a new workbook (oWorkbook := oExcel.Workbooks.Add)
	
	FilePath := "d:\Users\umpawy01\Desktop\Bruttowanie.xlsx" ; example path
	oWorkbook := ComObjGet(FilePath) ; access Workbook object
	
	column := 65 ; ASCII - A
	row := 2
	
	for xlCell in oWorkbook.Sheets(1).UsedRange.Cells{
		c1 := Chr(65) ; ASCII to String - A
		c2 := Chr(67) ; C
		cell = %c1%%row%:%c2%%row% ; e.g. A2
		if(oWorkbook.Sheets(1).Range(cell).Interior.ColorIndex = 43 || oWorkbook.Sheets(1).Range(cell).Interior.ColorIndex = 48){
			row++
			continue
		}
		oWorkbook.Sheets(1).Range(cell).Interior.ColorIndex := 43 ; green
		sleep, 50
		row++
		nextCode = %c1%%row%
		clipboard := Format("{:.0f}", oWorkbook.Sheets(1).Range(nextCode).Value)
		;msgbox % SubStr(String, 1, InStr(userCode, "."))
		;msgbox % SubStr(String, InStr(userCode, "."))
		
		;string := "asdfasdfasdfasdf - something is written here"
		;msgbox % SubStr(String, 1, InStr(string, "-")) 
		;msgbox % SubStr(String, InStr(string, "-"))
		
		break
	}
	return
}

; Hotkeys
; -------------

;Test
>+o::
	testing()
	;oExcel := ComObjCreate("Excel.Application") ; create Excel Application object
return

;GMK WŁ
>+g::
	;do_GMK_WL(FootnoteNum)
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
