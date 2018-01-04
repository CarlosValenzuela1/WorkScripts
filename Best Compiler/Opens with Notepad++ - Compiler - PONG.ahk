#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;msgbox Both buttons!
SetTitleMatchMode, 2
return

^SPACE::
	IfWinExist, PONG.ahk - Notepad++
	{
	MsgBox passed
	WinActivate ; use the window found above
	sleep 100
	send {Ctrl down}s{Ctrl up} ; Save
	Msgbox saved
	sleep 100
	Send {LWin down}r{LWin up}
	sleep 100
	send C:\Users\%A_UserName%\Desktop\PONG.ahk	;Run Process of EXE after saving it
	send {ENTER}
	;sleep 800
	;send Y
	Msgbox Ran
	}
	else
	{
		msgbox failed
	    Run, C:\Users\%A_UserName%\Documents\Other\Notepad++Portable\Notepad++Portable.exe C:\Users\%A_UserName%\Desktop\PONG.ahk
	    ;Run, edit C:\Users\%A_UserName%\Desktop\PONG.ahk	;Open editor if nothing is open
	}
  return