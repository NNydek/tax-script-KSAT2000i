Run:
	Gui, Submit, NoHide
	DateSelect := check_date_boxes(19_to_21, 20_to_21, 21_only)
	PWLPos := check_gpwl_boxes(PWL_1, PWL_2, PWL_3, PWL_4, PWL_5)
	SurnamePos := check_surname_boxes(1_Surname_Pos, 2_Surname_Pos, 3_Surname_Pos)
	WLPos := check_gwl_boxes(WL_1, WL_2, WL_3, WL_4, WL_5, WL_6, WL_7, WL_8, WL_9, WL_10)
	FootnoteNum := check_footnote_boxes(1_footnote, 2_footnote)
	if (!DateSelect || !SurnamePos || !FootnoteNum || !WLPos || !PWLPos)
		return 
	sleep, 1000
	WinWait, I Love Automation
	WinMinimize
	while(1){
		areParametersRight := check_parameters(PWLPos, SurnamePos, WLPos, FootnoteNum)
		if(!areParametersRight){
			WinWait, I Love Automation
			WinActivate, I Love Automation
			sleep, 500
			SendInput, {Enter}
			return
		}
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
