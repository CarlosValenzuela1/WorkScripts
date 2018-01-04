#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

return 

F10::
InputBox, Var, Teams Automator, Please type the team in lowercase: `n `n •For TAM please type: "tam"`n •For Blah, , 350,400
if (Var = "tam")
{
	send, Infra-TGRC-Technology Asset Management
	sleep 300
	send, {DOWN}
	SEND, {ENTER}
}
else
{
	MsgBox, Incorrect categorization
}

;SendInput, %Var%
return