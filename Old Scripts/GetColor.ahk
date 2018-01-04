#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;#SingleInstance Force
#Persistent

;^F1::

;***********************Get Color**************************
+^R::Reload      ;Reloads Script
+^P::Pause       ;Pauses  Script
Ins:: Suspend    ;Suspend hotkeys

;Ctrl+Leftmousebutton  = Start stop Color Detection
;Ctrl+RightMousebutton = Copy currently color to clipboard
;***********************************************************

#MaxThreadsPerHotkey 2

^LButton::
Toggle := !Toggle
While Toggle {
	MouseGetPos Xpos, Ypos
	PixelGetColor Colour, %Xpos%, %Ypos%, RGB
	StringTrimLeft Colour, Colour, 2					; Remove 0x
	ToolTip %Colour%
	}
ToolTip
Return

^RButton::
StringTrimLeft Colour, Colour, 2					; Remove 0x
Clipboard = %Colour%
ToolTip Copied to the clipboard:`n%Colour%
Sleep 3000
ToolTip %Colour%
Return
