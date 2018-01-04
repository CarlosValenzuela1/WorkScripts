#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its sCarloserior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

^F1::
	IfWinExist, Queuemanagement V1.7.xlsm - Excel
	{

    	WinActivate ; use the window found above
	sleep 700
	MouseToDutch()	
	sleep 500
	GetMinDutch()

	}

	else
	{
	    Run, edit C:\Users\c760912\Desktop\Queuemanagement V1.7.xlsm	;Open editor if nothing is open
	}
	
return

;========================== WE ARE OUTSIDE OF MAIN PROGRAM =================================;


popup(myvar)
{
	MsgBox, %myvar%
}


MouseToDutch()
{
	SendInput {Ctrl down}{HOME down}
	sendInput {Ctrl up}{Home up}
	sleep 500
	send {Ctrl down}g{ctrl up}
	sleep 100
	send c3
	send {ENTER}
}

GetMinDutch()
{
	send {Ctrl down}c{Ctrl up}
	sleep 500
	;put code here
	minVal:= Clipboard
	sleep 500
	popup(minVal)
}