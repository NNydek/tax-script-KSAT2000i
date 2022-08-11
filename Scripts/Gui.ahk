Gui, Show, w500 h500, I Love Automation

GUI, Add, Button, +Default x200 y450 w100 gRun, Run
GUI, Add, Button, -Default Hide x440 y10 w50 gWatch_Cursor, Watch Cursor

Gui, Add, Groupbox, x11 y10 Section w200 h40, Date
Gui, Add, Radio, xs+10 	ys+20 v3_dates, 19-21
Gui, Add, Radio, xs+80 	ys+20 v2_dates, 20-21
Gui, Add, Radio, xs+150 ys+20 v1_dates, 21

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
Gui, Add, Radio, xs+10 	ys+20 v1_Surname_Pos, Poz. 1
Gui, Add, Radio, xs+80 	ys+20 v2_Surname_Pos, Poz. 2
Gui, Add, Radio, xs+150 ys+20 v3_Surname_Pos, Poz. 3

Gui, Add, Groupbox, x11 y190 Section w370 h100, WŁ
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
Gui, Add, Radio, xs+10	ys+80 vWL_11, Poz. 11
Gui, Add, Radio, xs+80 	ys+80 vWL_12, Poz. 12
Gui, Add, Radio, xs+150	ys+80 vWL_13, Poz. 13
Gui, Add, Radio, xs+220 ys+80 vWL_14, Poz. 14
Gui, Add, Radio, xs+290 ys+80 vWL_15, Poz. 15

Gui, Add, Groupbox, x11 y300 Section w200 h40, Przypis
Gui, Add, Radio, xs+10 ys+20 v1_footnote, Pierwszy
Gui, Add, Radio, xs+80 ys+20 v2_footnote, Drugi

return
