testing(){
	sleep, 100
	Send, ^c
	sleep, 100
	if (clipboard != "G-PW" . Chr(0321) . "FNPN"){
		msgbox, RODZAJ NALEZNOSCI !!!
	}
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
	Loop{
		if(A_Index = 1)
			continue
		Send, {Down}
	} until (A_Index = WLPos)
	isMunicipality()
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
		if(FootnoteNum = 2){
			sleep, 1000
			MouseClick, left, 650, 230, 1 ; ZAZNACZ1
		} else{
			sleep, 1000
			MouseClick, left, 650, 270, 1 ; ZAZNACZ3
		}
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
	Send, {Up 15}
	Sleep, 350
	Send, {Enter}
	sleep, 1000
	Loop {
		if(A_Index = 1)
			continue
		Send, {Down}
	} until (A_Index = PWLPos)

	sleep, 1000
	MouseClick, left, 748, 385, 1 ; EWIDENCJA	
	sleep, 1000
	MouseClick, left, 125, 413, 1 ; NAZWISKO 1
	Loop{
		if(A_Index = 1)
			continue
		Send, {Down}
	} until (A_Index = SurnamePos)
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
	MouseClick, left, 32, 414, 1 ; KOD 1
	Loop{
		if(A_Index = 1)
			continue
		Send, {Down}
	} until (A_Index = SurnamePos)
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

is_dead(getPixelX, getPixelY){
	sleep 1000
	PixelGetColor, color, getPixelX, getPixelY 
	if(color = 0x0000FF)
		return 1
	else
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