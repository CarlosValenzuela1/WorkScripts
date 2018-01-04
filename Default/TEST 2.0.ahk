#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

return 

Pause::
InputBox, Var, Categorization Automator, Please type the categorization in lowercase: `n `n •For EUC1 Application please type: "euc1" `n •For Folder Access please type: "folder" `n •For Mobile Iron please type: "mobile" `n •For Printer Drivers please type: "printer" `n •For Software Installation please type: "sw" , , 350,400
if (Var = "euc1")
{
	send, Software
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, IBM i
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Authentication Issues
	sleep 300
	send, {DOWN}
	send, {ENTER}

	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Software
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Application
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Application Platform
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, jde world a9.2 emea
	sleep 300
	send, {DOWN}
	send, {ENTER}
}
else if (Var = "folder")
{
	send, Folder Access
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Files & Folders
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Access Request
	sleep 300
	send, {DOWN}
	send, {ENTER} 

	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Software
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Application
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Application Platform
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, cargill active directory (ad)
	sleep 300
	send, {DOWN}
	send, {ENTER}
}
else if (Var = "mobile")
{
		
	send, Request
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Account
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Reset
	sleep 300
	send, {DOWN}
	send, {ENTER}

	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Platform - Server 
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Hardware
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Applicance
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, MobileIron Virtual Smartphone Platform 
	sleep 300
	send, {DOWN}
	send, {ENTER}
}
else if (Var = "printer")
{
	send, Hardware
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Printer
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Configuration
	sleep 300
	send, {DOWN}
	send, {ENTER}

	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Hardware
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Peripheral
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Printer
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, {ENTER}
}
else if (Var = "sw")
{
	
	send, Software
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Operating System
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Install
	sleep 300
	send, {DOWN}
	send, {ENTER}

	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Software
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Application
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, Desktop
	sleep 300
	send, {TAB}
	sleep 300
	send, {TAB}
	send, {ENTER}
}
else
{
	MsgBox, Incorrect categorization
}

;SendInput, %Var%
return