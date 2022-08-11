check_date_boxes(dateArr){
	sum := 0
	Loop
		sum += dateArr[A_Index]
	until (A_Index = dateArr.MaxIndex())
	if(!sum){
		msgbox, Select at least 1 checkbox (date)
		return 0	
	}
	Loop{
		if(dateArr[A_Index])
			return %A_Index%
	} until (A_Index = dateArr.MaxIndex())
	msgbox, Error (check_date_boxes else)
	return
}

check_gpwl_boxes(pwlArr){
	sum := 0
	Loop
		sum += pwlArr[A_Index]
	until (A_Index = pwlArr.MaxIndex())
	if(!sum){
		msgbox, Select at least 1 checkbox (pwl)
		return 0
	}
	Loop {
		if (pwlArr[A_Index])
			return %A_Index%
	} until (A_Index = pwlArr.MaxIndex())
	msgbox, Error (check_gpwl_boxes else)
	return
}

check_surname_boxes(surnameArr){
	sum := 0
	Loop
		sum += surnameArr[A_Index]
	until (A_Index = surnameArr.MaxIndex())
	if(!sum){
		msgbox, Select at least 1 checkbox (surname)
		return 0
	}
	Loop {
		if (surnameArr[A_Index])
			return %A_Index%
	} until (A_Index = surnameArr.MaxIndex())
	msgbox, Error (check_surname_boxes else)
	return
}

check_gwl_boxes(wlArr){
	sum := 0
	Loop
		sum += wlArr[A_Index]
	until (A_Index = wlArr.MaxIndex())
	if(!sum){
		msgbox, Select at least 1 checkbox (wl)
		return 0
	}
	Loop {
		if (wlArr[A_Index])
			return %A_Index%
	} until (A_Index = wlArr.MaxIndex())
	msgbox, Error (check_gwl_boxes else)
	return
} 

check_footnote_boxes(footnoteArr){
	sum := 0
	Loop
		sum += footnoteArr[A_Index]
	until (A_Index = footnoteArr.MaxIndex())
	if(!sum){
		msgbox, Select at least 1 checkbox (footnote)
		return 0
	}
	Loop {
		if (footnoteArr[A_Index])
			return %A_Index%
	} until (A_Index = footnoteArr.MaxIndex())
	msgbox, Error (check_footnote_boxes else)
	return
}

check_parameters(PWLPos, SurnamePos, WLPos, FootnoteNum, DateSelect){
	sleep, 200
	oExcel := ComObjActive("Excel.Application")
	
	FilePath := "d:\Users\umpawy01\Desktop\Bruttowanie.xlsx"
	oWorkbook := ComObjGet(FilePath) ; access Workbook object
	
	column := 68 ; ASCII - D
	row := 2

	for xlCell in oWorkbook.Sheets(1).UsedRange.Cells{
		c1 := Chr(68) ; ASCII to String - D
		c2 := Chr(72) ; 67 - C    71 - H
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
			msgbox,,, Wrong PWL, 1
			correctValue := Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value)
			GuiControl,, PWL_%correctValue%, 1
			return 0
		}
		sleep, 100
		newCell = E%row%
		SRNPos = %SurnamePos%
		sleep, 100
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != SRNPos){
			msgbox,,, Wrong SURNAME, 1
			correctValue := Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value)
			GuiControl,, %correctValue%_Surname_Pos, 1			
			return 0
		}
		sleep, 100
		newCell = F%row%
		WL = %WLPos%
		sleep, 100
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != WL || WL = PWL){
			msgbox,,, Wrong WL, 1
			correctValue := Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value)
			GuiControl,, WL_%correctValue%, 1			
			return 0
		}
		sleep, 100
		newCell = G%row%
		FNPos = %FootnoteNum%
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != FNPos){
			msgbox,,, Wrong FOOTNOTE, 1
			correctValue := Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value)
			GuiControl,, %correctValue%_footnote, 1
			return 0
		}
		sleep, 100
		newCell = H%row%
		DSelect = %DateSelect%
		if(Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value) != DSelect){
			msgbox,,, Wrong DATESELECT, 1
			correctValue := Format("{:.0f}",oWorkbook.Sheets(1).Range(newCell).Value)
			GuiControl,, %correctValue%_dates, 1
			return 0
		}
		sleep, 50
		break
	}
	return 1
}


check_popup(getPixelX, getPixelY){
	Loop, 60{
		sleep 250
		PixelGetColor, color, getPixelX, getPixelY
		if(color = 0x0000FF || color = 0xD1B499 || color = 0x00FFFF || color = 0xFF0000 || color = 0x7F7F00 || color = 0xE0E000 || 0xD1B499)
			return 1
	}
	msgbox, Something went wrong (check_popup)
	return 0
}

check_popup_fakturowanie(getPixelX, getPixelY){
	Loop, 60{
		sleep 250
		PixelGetColor, color, getPixelX, getPixelY
		if(color = 0x000000)
			return 1
	}
	msgbox, Something went wrong (check_popup_fakturowanie)
	return 0
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


isAlreadyDone(PWLPos){
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	sleep, 500
	MouseClick, left, 58, 222, 1 ; NUMER UMOWY ZAZN. 1
	sleep, 500
	Send, {Up 9}
	check_popup(370,519)
	Send, {Enter}
	sleep, 400
	Loop {
		if(A_Index = 1)
			continue
		Send, {Down}
	} until (A_Index = PWLPos)
	sleep, 1000
	MouseClick, left, 290, 558, 1 ; KOD(2)
	sleep, 1000
	PixelGetColor, color, 341, 523 ; 0x0000FF RED
	if (color != 0x0000FF){
		Send, {Enter}
		clear()
		update_excel()
		paste_new_user_code()
	} else
		Send, {Enter}
	isMunicipality()
	sleep, 1000
	MouseClick, left, 58, 222, 1 ; NUMER UMOWY ZAZN. 1
	sleep, 1000
		Send, {Up 15}
	check_popup(370,519)
		Send, {Enter}
	return
}

isMunicipality(){
	sleep, 100
	MouseClick, left, 240, 235
	sleep, 250
	Send, ^c
	sleep, 100
	if (RegExMatch(clipboard,"G-PW" . Chr(0321) . "FN") = 1 || RegExMatch(clipboard,"G-W" . Chr(0321) . "FN") = 1)
		return
	else
		msgbox, RODZAJ NALEZNOSCI !!!
	return
}