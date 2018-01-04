#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

return 

Pause::
InputBox, Var, Beta 3.0, •EUC1 Aplication type: "euc1" `n•Folder Access type: "folder" `n•MobileIron type: "mobile" `n•Printer Drivers type: "printer" `n•Software Installation type: "sw"`n•Resolution Categorization type: "resolution"  `n `n•Netasthra FILTER (Assigned Group Name) type: "agn"  `n•Incident Distribution EXCEL type: "inc"  `n•Service Request EXCEL please type: "service" , , 400,400

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
sleep 2000
MouseClick, left, 519, 701	;Click on Search bar
sleep 2000
send, emea-central service desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-Central Hungary Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-Central Poland Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-Central Romania Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-Central Turkey Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-Central West Africa Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-North Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-East Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-South Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 500
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}

sleep 1000
send, Infra-SD-EMEA-
sleep 4000
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
else if (Var = "inc")
{
CoordMode, Mouse, Screen
MouseClick, right, 1356,417
send {UP}
send {ENTER}
Sleep 1000 ;to wait for inspect code to appear
MouseClick, left, 1516, 1535
send ^c
;WinActivate ; use the window found above
sleep 1000
MouseClick, left, 781, 811
Sleep 1000
MouseClick, left
sleep 1000
StartPos := RegExMatch(ClipBoard, ">(.*)</td>", SubPat)
ClipBoard := SubPat1
send ^v
sleep 1000

;Now to copy the value in Progress
MouseClick, left, 1424, 1541
send ^c
sleep 500
MouseClick, left, 800, 830
sleep 700
MouseClick, left, 800, 830
sleep 700
StartPos := RegExMatch(ClipBoard, ">(.*)</td>", SubPat)
ClipBoard := SubPat1
sleep 100
send ^v

;Now to copy the value in Pending
MouseClick, left, 1413, 1576
send ^c
sleep 500
MouseClick, left, 808, 844
sleep 700
MouseClick, left, 808, 844
sleep 700
StartPos := RegExMatch(ClipBoard, ">(.*)</td>", SubPat)
ClipBoard := SubPat1
sleep 100
send ^v

;Now to copy the value in Resolved
MouseClick, left, 1458, 1607
send ^c
sleep 500
MouseClick, left, 797, 861
sleep 700
MouseClick, left, 797, 861
sleep 700
StartPos := RegExMatch(ClipBoard, ">(.*)</td>", SubPat)
ClipBoard := SubPat1
sleep 100
send ^v
}
else if (Var = "service")
{
CoordMode, Mouse, Screen
MouseClick, right, 1332, 433
send {UP}
send {ENTER}
Sleep 1000 ;to wait for inspect code to appear
MouseClick, left, 1516, 1535
send ^c
;WinActivate ; use the window found above
sleep 1000
MouseClick, left, 775, 878
Sleep 1000
MouseClick, left
sleep 1000
StartPos := RegExMatch(ClipBoard, ">(.*)</td>", SubPat)
ClipBoard := SubPat1
send ^v
sleep 1000

;Now to copy the value in Progress
MouseClick, left, 1454, 1559
send ^c
sleep 500
MouseClick, left, 778, 896
sleep 700
MouseClick, left
sleep 700
StartPos := RegExMatch(ClipBoard, ">(.*)</td>", SubPat)
ClipBoard := SubPat1
sleep 100
send ^v

;Now to copy the value in Pending
MouseClick, left, 1448, 1587
send ^c
sleep 500
MouseClick, left, 779, 914
sleep 700
MouseClick, left
sleep 700
StartPos := RegExMatch(ClipBoard, ">(.*)</td>", SubPat)
ClipBoard := SubPat1
sleep 100
send ^v

;Now to copy the value in Resolved
MouseClick, left, 1398, 1614
send ^c
sleep 500
MouseClick, left, 781, 929
sleep 700
MouseClick, left
sleep 700
StartPos := RegExMatch(ClipBoard, ">(.*)</td>", SubPat)
ClipBoard := SubPat1
sleep 100
send ^v
}
else
{
	MsgBox, Incorrect categorization
}
