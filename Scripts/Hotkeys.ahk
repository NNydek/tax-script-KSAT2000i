>+o:: ;Test
	;testing()
	isMunicipality()
return

Numpad1:: ;19-21
	group_people(1)
return

Numpad2:: ;20-21
	group_people(2)
return

Numpad3:: ;21
	group_people(3)
return

Numpad4:: ;todo
	group_people(4)
return

Numpad7:: ;go back to worksheet 1
	if !WinActive("Microsoft Excel - Bruttowanie")
		return
	Send, {LCtrl down}{PgUp 7}{LCtrl up}
return

Numpad8:: ;Check new code
	if !WinActive("GOSPODARKA MIENIEM KOMUNALNYM")
		WinActivate GOSPODARKA MIENIEM KOMUNALNYM
	clear()
	check_new_code()
	paste_new_user_code()
	sleep, 1000
	SendInput, ]
	check_popup(221, 189)
	Send, {Down 20}
	sleep, 350
	Send, {Enter}
return

>+p:: ;Pause
	Pause
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
