#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent

; GUI Layout
; -------------

#Include Scripts\Gui.ahk

CoordMode, ToolTip, screen
SetTimer, Watch_Cursor, 100 
return

; Labels
; -------------

#Include Scripts\Labels.ahk

; Functions
; -------------

#Include Scripts\Functions.ahk

; Hotkeys
; -------------

#Include Scripts\Hotkeys.ahk