#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent

; GUI Layout
; -------------
Gui, Show, w500 h500, I Love Automation

GUI, Add, Button, x440 y10 w50 gWatch_Cursor, Watch Cursor
GUI, Add, Button, x200 y450 w100 gRun, Run

Gui, Add, Groupbox, x11 y10 Section w200 h40, Date
Gui, Add, Radio, xs+10 	ys+20 v19_to_21, 19-21
Gui, Add, Radio, xs+80 	ys+20 v20_to_21, 20-21
Gui, Add, Radio, xs+150 ys+20 v21_only, 21

Gui, Add, Groupbox, x11 y60 Section w370 h70, PWŁ 
Gui, Add, Radio, xs+10 	ys+20 vPWL_1, Poz. 1
Gui, Add, Radio, xs+80 	ys+20 vPWL_2, Poz. 2
Gui, Add, Radio, xs+150 ys+20 vPWL_3, Poz. 3
Gui, Add, Radio, xs+220 ys+20 vPWL_4, Poz. 4
Gui, Add, Radio, xs+290 ys+20 vPWL_5, Poz. 5
Gui, Add, Radio, xs+10 	ys+50 vPWL_6, Poz. 6
Gui, Add, Radio, xs+80 	ys+50 vPWL_7, Poz. 7
Gui, Add, Radio, xs+150 ys+50 vPWL_8, Poz. 8
Gui, Add, Radio, xs+220 ys+50 vPWL_9, Poz. 9
Gui, Add, Radio, xs+290 ys+50 vPWL_10, Poz. 10

Gui, Add, Groupbox, x11 y140 Section w220 h40, Nazwisko
Gui, Add, Radio, xs+10 	ys+20 vFirst_Surname_Pos, Poz. 1
Gui, Add, Radio, xs+80 	ys+20 vSecond_Surname_Pos, Poz. 2
Gui, Add, Radio, xs+150 ys+20 vThird_Surname_Pos, Poz. 3

Gui, Add, Groupbox, x11 y190 Section w370 h70, WŁ
Gui, Add, Radio, xs+10 	ys+20 vWL_1, Poz. 1
Gui, Add, Radio, xs+80 	ys+20 vWL_2, Poz. 2
Gui, Add, Radio, xs+150 ys+20 vWL_3, Poz. 3
Gui, Add, Radio, xs+220 ys+20 vWL_4, Poz. 4
Gui, Add, Radio, xs+290 ys+20 vWL_5, Poz. 5
Gui, Add, Radio, xs+10 	ys+50 vWL_6, Poz. 6
Gui, Add, Radio, xs+80	ys+50 vWL_7, Poz. 7
Gui, Add, Radio, xs+150 ys+50 vWL_8, Poz. 8
Gui, Add, Radio, xs+220 ys+50 vWL_9, Poz. 9
Gui, Add, Radio, xs+290 ys+50 vWL_10, Poz. 10

Gui, Add, Groupbox, x11 y270 Section w200 h40, Przypis
Gui, Add, Radio, xs+10 ys+20 vFirst_footnote, Pierwszy
Gui, Add, Radio, xs+80 ys+20 vSecond_footnote, Drugi

return

CoordMode, ToolTip, screen
SetTimer, Watch_Cursor, 100 
return

; Labels
; -------------

Run:
	Gui, Submit, NoHide
	DateSelect := check_date_boxes(19_to_21, 20_to_21, 21_only)
	PWLPos := check_gpwl_boxes(PWL_1, PWL_2, PWL_3, PWL_4, PWL_5)
	SurnamePos := check_surname_boxes(First_Surname_Pos, Second_Surname_Pos, Third_Surname_Pos)
	WLPos := check_gwl_boxes(WL_1, WL_2, WL_3, WL_4, WL_5, WL_6, WL_7, WL_8, WL_9, WL_10)
	FootnoteNum := check_footnote_boxes(First_footnote, Second_footnote)
	if (!DateSelect || !SurnamePos || !FootnoteNum || !WLPos || !PWLPos)
		return 
	sleep, 1000
	WinWait, I Love Automation
	WinMinimize
	while(1){
		areParametersRight := check_parameters(PWLPos, SurnamePos, WLPos, FootnoteNum)
		if(!areParametersRight)
			return
		isAlreadyDone(PWLPos)
		do_GMK_WL(FootnoteNum, WLPos, DateSelect)
		do_GMK_PWL(DateSelect, SurnamePos, PWLPos)
		do_FAKTUROWANIE()
		update_excel()
		paste_new_user_code()
		;msgbox, stop
	}
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

check_date_boxes(19_to_21, 20_to_21, 21_only){
	if (!19_to_21 && !20_to_21 && !21_only){
		msgbox, Select at least 1 checkbox (date)
		return 0
	} else if (21_only)
		return 1
	else if (20_to_21)
		return 2
	else if (19_to_21)
		return 3
	else
		msgbox, Error (check_date_boxes else)
	return
}

check_gwl_boxes(a, b, c, d, e, f, g, h, i, j){
	if (!a && !b && !c && !d && !e && !f && !g && !h && !i && !j){
		msgbox, Select at least 1 checkbox (wl)
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
		msgbox, Select at least 1 checkbox (przypis)
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
		msgbox, Select at least 1 checkbox (nazwisko)
		return 0
	} else if (1_pos = 1)
		return 1
	else if (2_pos = 1)
		return 2
	else if (3_pos = 1){
		msgbox, WiP
		return 3
	} else
		msgbox, Error (check_surname_boxes else)
	return
}

do_GMK_WL(FootnoteNum, WLPos, DateSelect){
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
	if (DateSelect >= 2){
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
		check_popup(479, 500)
		Send, {Enter} ; OK
		sleep, 1000
		MouseClick, left, 650, 230, 1 ; ZAZNACZ1
		sleep, 1000
		MouseClick, left, 604, 377, 1 ; GENERUJ KOREKTE
		sleep, 1000
		MouseClick, left, 415, 416, 1 ; KOREKTA DO ZERA
		sleep, 1000
		MouseClick, left, 484, 470, 1 ; ZATWIERDŹ
		sleep, 1000
		check_popup(479, 500)
		Send, {Enter} ; OK
		sleep, 1000
		MouseClick, left, 777, 93, 1 ; WYJSCIE
	} 
	if (DateSelect = 3){
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
		sleep, 1000
		check_popup(479, 500)
		Send, {Enter} ; OK
		sleep, 1000
		MouseClick, left, 777, 93, 1 ; WYJSCIE
	} else if (DateSelect = 1){
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
		sleep, 1000
		check_popup(479, 500)
		Send, {Enter} ; OK
		sleep, 1000
		MouseClick, left, 777, 93, 1 ; WYJSCIE
	} else if (DateSelect != 2)
		msgbox, Error (do_GMK_WL if(FootnoteNum2) else)
	sleep, 1000
	MouseClick, left, 777, 93, 1 ; WYJSCIE
	return
}

do_GMK_PWL(DateSelect, SurnamePos, PWLPos){
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
	Switch DateSelect{
	case 3: ; 19-21
		Send, {Tab}{Tab}01-01-2019{Tab}31-12-2021
	case 2: ; 20-21
		Send, {Tab}{Tab}01-01-2020{Tab}31-12-2021
	case 1: ; 21
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

	Loop, %DateSelect%{
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

		Switch DateSelect{
		case 3: ; 19-21
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
		case 1: ; 21
			Send, 31-03-2021{Tab}31-03-2021{Tab}01-01-2021
		default: msgbox, Error (do_GMK_PWL default)
		}

		sleep, 1000
		MouseClick, left, 608, 127, 1 ; Generuj Dokument
		check_popup(480, 490)
		sleep, 1000
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
	
	check_popup(480, 490)
	sleep, 1000
	Send, {Enter}
	check_popup(480, 490)
	sleep, 1000
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
		c2 := Chr(71) ; 67 - C    71 - G
		cell = %c1%%row%:%c2%%row% ; e.g. A2
		if(oWorkbook.Sheets(1).Range(cell).Interior.Color != 16777215.0000000){
			row++
			continue
		}
		oWorkbook.Sheets(1).Range(cell).Interior.ColorIndex := 43 ; green 48 grey
		sleep, 50
		row++
		nextCode = %c1%%row%
		clipboard := Format("{:.0f}", oWorkbook.Sheets(1).Range(nextCode).Value)
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

isAlreadyDone(PWLPos){
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	sleep, 1000
	MouseClick, left, 58, 222, 1 ; NUMER UMOWY ZAZN. 1
	sleep, 1000
	Send, {Up 9}
	Sleep, 350
	Send, {Enter}
	sleep, 500
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
		msgbox, Error (isAlreadyDone Switch PWLPos)
	}
	sleep 1000
	MouseClick, left, 243, 558, 1 ; DATA DO
	sleep 1000
	MouseMove, 341, 523
	PixelGetColor, color, 341, 523 ; 0x0000FF RED
	if (color != 0x0000FF){
		Send, {Enter}
		clear()
		update_excel()
		paste_new_user_code()
	} else
		Send, {Enter}
	return
}
;389, 133 FAKTUROWANIE
;474, 487
;411, 171
;381, 520
;384, 512
;591, 433

check_parameters(PWLPos, SurnamePos, WLPos, FootnoteNum){
	oExcel := ComObjCreate("Excel.Application") ; create Excel Application object
	oExcel.Workbooks.Add ; create a new workbook (oWorkbook := oExcel.Workbooks.Add)
	
	FilePath := "d:\Users\umpawy01\Desktop\Bruttowanie.xlsx" ; example path
	oWorkbook := ComObjGet(FilePath) ; access Workbook object
	
	column := 68 ; ASCII - D
	row := 2

	for xlCell in oWorkbook.Sheets(1).UsedRange.Cells{
		c1 := Chr(68) ; ASCII to String - D
		c2 := Chr(71) ; 67 - C    71 - G
		cell = %c1%%row%:%c2%%row% ; e.g. D2:G2
		if(oWorkbook.Sheets(1).Range(cell).Interior.Color != 16777215.0000000){
			row++
			continue
		}
		newCell = %c1%%row%
		PWL = %PWLPos%
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != PWL){
			msgbox, WRONG PWL
			return 0
		}
		newCell = E%row%
		SRNPos = %SurnamePos%
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != SRNPos){
			msgbox, WRONG SURNAME
			return 0
		}
		newCell = F%row%
		WL = %WLPos%
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != WL){
			msgbox, WRONG WL
			return 0
		}
		newCell = G%row%
		FNPos = %FootnoteNum%
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != FNPos){
			msgbox, WRONG FOOTNOTE
			return 0
		}
		sleep, 50
		break
	}
	return 1
}

check_new_code(){
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM

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
		if(oWorkbook.Sheets(1).Range(cell).Interior.Color != 16777215.0000000){
			row++
			continue
		}
		oWorkbook.Sheets(1).Range(cell).Interior.ColorIndex := 6 ; yellow
		sleep, 50
		row++
		nextCode = %c1%%row%
		clipboard := Format("{:.0f}", oWorkbook.Sheets(1).Range(nextCode).Value)
		break
	}
	return
}

check_popup(getPixelX, getPixelY){
	Loop, 15{
		sleep 1000
		PixelGetColor, color, getPixelX, getPixelY 
		if(color = 0x0000FF || color = 0xD1B499 || color = 0x00FFFF || color = 0xFF0000 || color = 0x7F7F00)
			return 1
	}
	msgbox, Something went wrong (check_popup)
	return 0
}

testing(){

}

; Hotkeys
; -------------


>+o:: ;Test
	msgbox, pause
return

=:: ;Check new code
	clear()
	check_new_code()
	paste_new_user_code()
return

>+c:: ;CLEAR
	clear()
return

\:: ;Ewidencja
	sleep, 250
	MouseClick, left, 748, 388, 1
return

]:: ;Konto Kontrahenta
	sleep, 250
	MouseClick, left, 747, 427, 1

	sleep, 1000
	MouseClick, left, 441, 394, 1
return

[:: ;Wyjscie z kontrahenta
	sleep, 250
	MouseClick, left, 797, 94, 1
	
	sleep, 250
	MouseClick, left, 236, 399, 1
return

>^x::ExitApp
