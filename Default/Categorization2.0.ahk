#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

return 

Pause::
InputBox, Var, Categorization Automator, Please type the categorization: `n `n •For EUC1 Aplication please type: "euc1" `n •For Folder Access please type: "folder" `n •For Mobile Iron please type: "mobile" `n •For Printer Drivers please type: "printer" `n •For Software Installation please type: "sw" `n `n •For Resolution Categorization please type: "resolution"  `n `n •For Netasthra FILTER (Assigned Group Name) type: "agn" , , 400,400
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
else if (Var = "resolution")
{
	send, Performance
	sleep 300
	send, {TAB}
	sleep 400
        send, {TAB}
	sleep 600
        send, Application
	sleep 400
	send, {TAB}
	sleep 400
	send, {TAB}
	send, A
	sleep 800
	send, {DOWN}
	send, {ENTER}

        sleep 200
	send, {TAB}
	sleep 200
	send, {TAB}
	sleep 200
	send, {TAB}
	sleep 200
	send, {TAB}
        sleep 200
	send, {TAB}
	sleep 200
	send, {TAB}
	sleep 200
	send, {TAB}
	sleep 200
	send, {TAB}
        sleep 200
	send, {TAB}
        send, generic
	sleep 300
	send, {ENTER}
}
else if (Var = "agn")
{
MouseClick, left, 523, 661	;Client UNSELECT ALL
sleep 300
MouseClick, left, 519, 701	;Click on Search bar
/*
send, emea-central service desk
sleep 1200
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
send, EMEA-Central Hungary Service Desk
sleep 1200
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
send, EMEA-Central Poland Service Desk
sleep 1200
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
send, EMEA-Central Romania Service Desk
sleep 1200
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
send, EMEA-Central Turkey Service Desk
sleep 1200
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 100
send, EMEA-Central West Africa Service Desk
sleep 1200
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 100
send, EMEA-North Service Desk
sleep 1200
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 100
send, EMEA-East Service Desk
sleep 1200
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 100
send, EMEA-South Service Desk
sleep 1200
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
*/
sleep 800
send, Infra-SD-EMEA-
sleep 3400
MouseClick, left, 520, 742	;Click on first item on list
sleep 10
MouseClick, left, 514, 772	;Click on second item on list
sleep 10
MouseClick, left, 514, 802	;Click on third item on list
sleep 10
MouseClick, left, 514, 832	;Click on fourth item on list
sleep 10
MouseClick, left, 514, 862	;Click on fifth item on list
sleep 10
MouseClick, left, 514, 892	;Click on sixth item on list
sleep 10
MouseClick, left, 514, 922	;Click on seventh item on list
sleep 10
MouseClick, left, 514, 952	;Click on eighth item on list
sleep 10
MouseClick, left, 514, 982	;Click on nineth item on list
sleep 10
MouseClick, left, 514, 1012	;Click on tenth item on list
sleep 10
MouseClick, left, 514, 1042	;Click on eleventh item on list
Click WheelDown			;will scroll down
sleep 10
MouseClick, left, 514, 1060	;Click on twelveth item on list
}
else
{
	MsgBox, Incorrect categorization
}
