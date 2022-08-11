testing(asdf){
	msgbox, testing %asdf%
	return
}

check_parameters(PWLPos, SurnamePos, WLPos, FootnoteNum){
	sleep, 200
	oExcel := ComObjActive("Excel.Application")
	
	FilePath := "d:\Users\umpawy01\Desktop\Bruttowanie.xlsx"
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
		sleep, 100
		newCell = %c1%%row%
		PWL = %PWLPos%
		sleep, 100
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != PWL){
			msgbox,,, Wrong PWL, 2
			correctValue := Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value)
			GuiControl,, PWL_%correctValue%, 1
			return 0
		}
		sleep, 100
		newCell = E%row%
		SRNPos = %SurnamePos%
		sleep, 100
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != SRNPos){
			msgbox,,, Wrong SURNAME, 2
			correctValue := Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value)
			GuiControl,, %correctValue%_Surname_Pos, 1			
			return 0
		}
		sleep, 100
		newCell = F%row%
		WL = %WLPos%
		sleep, 100
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != WL || WL = PWL){
			msgbox,,, Wrong WL, 2
			correctValue := Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value)
			GuiControl,, WL_%correctValue%, 1			
			return 0
		}
		sleep, 100
		newCell = G%row%
		FNPos = %FootnoteNum%
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != FNPos){
			msgbox,,, Wrong FOOTNOTE, 2
			correctValue := Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value)
			GuiControl,, %correctValue%_footnote, 1
			return 0
		}
		sleep, 50
		break
	}
	return 1
}

isAlreadyDone(PWLPos){
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	sleep, 1000
	MouseClick, left, 58, 222, 1 ; NUMER UMOWY ZAZN. 1
	sleep, 1000
	Send, {Up 9}
	Sleep, 1000
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
	MouseClick, left, 290, 558, 1 ; KOD(2)
	sleep 1000
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
		if(FootnoteNum = 2){
			sleep, 1000
			MouseClick, left, 38, 577, 1 ; PRZYPIS 2
		} else{
			sleep, 1000
			MouseClick, left, 38, 557, 1 ; PRZYPIS 1
		}
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
		check_popup(480, 485) ;480 500
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
		check_popup(480, 485) ;480 500
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
		check_popup(480, 485) ;480 500
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
		check_popup(480, 485) ;480 500
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
		sleep, 200

		isDead := is_dead(388, 512)
		if(isDead){
			MsgBox, if isDead true
			sleep, 1000
			Send, {Enter}
			sleep, 1000
			MouseClick, left, 715, 165, 1 ; OPERACJE
			sleep, 1000
			MouseClick, left, 410, 214, 1 ; Generj fakture/dokument
			sleep, 1000
			Send, {Enter}
		}
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
		default: 
			msgbox, Error (do_GMK_PWL default)
		}
		if(isDead){
			sleep, 1000
			MouseClick, left, 608, 127, 1 ; Generuj Dokument
			check_popup(480, 485) ;480 490
			sleep, 1000
			send, {Enter}
			check_popup(480, 475) ;480 490
			sleep, 500
			send, {Enter}
			check_popup(480, 483) ;480 490
			sleep, 500
			send, {Enter}
		} else{
			sleep, 1000
			MouseClick, left, 608, 127, 1 ; Generuj Dokument
			
			check_popup(480, 470) ;480 490
			sleep, 1000
			Send, {Enter}
		}
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
	oExcel := ComObjActive("Excel.Application")
	
	FilePath := "d:\Users\umpawy01\Desktop\Bruttowanie.xlsx"
	oWorkbook := ComObjGet(FilePath) ; access Workbook object
	
	column := 65 ; ASCII - A
	row := 2
	
	for xlCell in oWorkbook.Sheets(1).UsedRange.Cells{
		c1 := Chr(65) ; ASCII to String - A
		c2 := Chr(71) ; 67 - C    71 - G
		cell = %c1%%row%:%c2%%row% ; e.g. A2:G2
		if(oWorkbook.Sheets(1).Range(cell).Interior.Color != 16777215.0000000){
			row++
			continue
		}
		oWorkbook.Sheets(1).Range(cell).Interior.ColorIndex := 43 ; green
		sleep, 50
		row++
		nextCode = %c1%%row%
		if(oWorkbook.Sheets(1).Range(nextCode).Value != "")
			clipboard := Format("{:.0f}", oWorkbook.Sheets(1).Range(nextCode).Value)
		else {
			msgbox, There are no more codes (Pause)
			Pause
		}
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

check_new_code(){
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM

	oExcel := ComObjCreate("Excel.Application")
	
	FilePath := "d:\Users\umpawy01\Desktop\Bruttowanie.xlsx"
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

is_dead(getPixelX, getPixelY){
	sleep 1000
	PixelGetColor, color, getPixelX, getPixelY 
	if(color = 0x0000FF)
		return 1
	else
		return 0
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

group_people(sheet){
	if WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		SendInput, [
	oExcel := ComObjActive("Excel.Application")
	
	FilePath := "d:\Users\umpawy01\Desktop\Bruttowanie.xlsx"
	oWorkbook := ComObjGet(FilePath) ; access Workbook object
	
	
	switch sheet{
	case 1:	
		sheet_date := oWorkbook.sheets(2)
	case 2:
		sheet_date := oWorkbook.sheets(3)
	case 3:
		sheet_date := oWorkbook.sheets(4)
	case 4:
		sheet_date := oWorkbook.sheets(5)
	}
	
	row := 2
	for xlCell in sheet_date.UsedRange.Cells{
		c1 := Chr(65) ; ASCII to String - A
		c2 := Chr(67) ; C
		cellDate = %c1%%row%:%c2%%row%
		if(sheet_date.Range(c1 . row).value != ""){
			row++
			continue
		}
		row := 2
		for xlCell in oWorkbook.Sheets(1).UsedRange.Cells{
			c1 := Chr(65) ; ASCII to String - A
			c2 := Chr(67) ; C
			cellSh1 = %c1%%row%:%c2%%row% ; e.g. A2
			if(oWorkbook.Sheets(1).Range(cellSh1).Interior.Color != 16777215.0000000){
				row++
				continue
			}
			sheet_date.Range(cellDate) := oWorkbook.Sheets(1).Range(cellSh1).value
			break
		}
		break
	}
	sleep, 1000
	if !WinActive("Microsoft Excel - Bruttowanie"){
		WinActivate Microsoft Excel - Bruttowanie
		WinWaitActive Microsoft Excel - Bruttowanie
	}
	sleep, 300
	switch sheet{
	case 1:
		Send, {LCtrl down}{PgDn}{LCtrl up}
	case 2:
		Send, {LCtrl down}{PgDn 2}{LCtrl up}
	case 3:
		Send, {LCtrl down}{PgDn 3}{LCtrl up}
	case 4:
		Send, {LCtrl down}{PgDn 4}{LCtrl up}
	}
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