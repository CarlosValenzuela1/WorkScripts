#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;msgbox Both buttons!

^SPACE::
	IfWinExist, Netas2.0.ahk - Notepad
	{
	WinActivate ; use the window found above
	sleep 100
	send {Ctrl down}s{Ctrl up} ; Save
	sleep 100
	Send {LWin down}r{LWin up}
	sleep 100
	send C:\Users\c760912\Desktop\Netas2.0.ahk	;Run Process of EXE after saving it
	send {ENTER}
	;sleep 800
	;send Y
	}
	else
	    Run, edit C:\Users\c760912\Desktop\Netas2.0.ahk	;Open editor if nothing is open
	
  return