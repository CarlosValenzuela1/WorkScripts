#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

return 

Pause::
InputBox, Var, Beta 3.0, •EUC1 Aplication type: "euc1" `n•Folder Access type: "folder" `n•MobileIron type: "mobile" `n•Printer Drivers type: "printer" `n•Software Installation type: "sw"`n•Resolution Categorization type: "resolution"`n•Autofill DS info type: "ds"  `n `n•Netasthra FILTER (Assigned Group Name) type: "agn"  `n•Incident Distribution EXCEL type: "inc"  `n•Service Request EXCEL please type: "service" , , 400,400

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
else if (Var = "ds")
{
	CoordMode, Mouse, Screen
	MouseClick, left, 844, 560
	send ^a
	send ^c
	sleep 100
	StartPos := RegExMatch(ClipBoard, "City: (.*) S", city)
	ClipBoard := city1
	MouseClick, left, 837, 1256
	send {ctrl down}{HOME}{ctrl up}
	send {down 18}
	send {end} 
	^v
	send {,} 
	MouseClick, left, 844, 560
	send ^a
	send ^c
	sleep 100
	StartPos := RegExMatch(ClipBoard, "Country: (.*) G", country)
	ClipBoard := country1
	MouseClick, left, 837, 1256
	send {ctrl down}{HOME}{ctrl up}
	send {down 18}
	send {end} 
	^v
	MouseClick, left, 844, 560
	send ^a
	send ^c
	sleep 100
	StartPos := RegExMatch(ClipBoard, "e Number: (.*)", number)
	ClipBoard := number1
	MouseClick, left, 837, 1256
	send {ctrl down}{HOME}{ctrl up}
	send {down 8}
	send {end} 
	^v
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
	;send {ctrl down}{shift down}i{ctrl up}{shift up}
	MouseClick, right, 1518, 957
	send {UP}
	send {ENTER}
	Sleep 1000 ;to wait for inspect code to appear
	MouseClick, left, 1484, 1533
	send {Right}
	sleep 200
	send {Down}
	sleep 200
	send {Right}
	sleep 200
	send {Down}
	sleep 200
	send {Down}
	sleep 200
	;Add more code
	send ^c
	;WinActivate ; use the window found above
	sleep 1000
	MouseClick, left, 781, 811
	Sleep 500
	MouseClick, left
	sleep 1000
	temp := Clipboard
	sleep 1000
	StartPos := RegExMatch(ClipBoard, "Total:</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""17"" aria-describedby=""list1_Assigned"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""20"" aria-describedby=""list1_In Progress"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""26"" aria-describedby=""list1_Pending"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""324"" aria-describedby=""list1_Resolved"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""387"" aria-describedby=""list1_Total"">387</td></tr></tbody></table>", SubPat)
	sleep 200

	;Replaces Clipoard with first value of table
	ClipBoard := SubPat1
	;MsgBox %temp%
	MsgBox %SubPat1%
	MsgBox %SubPat2%
	MsgBox %SubPat3%
	MsgBox %SubPat4%
	send ^v
	sleep 1000

	;Replaces Clipoard with second value of table
	MouseClick, left, 800, 830
	sleep 700
	MouseClick, left, 800, 830
	sleep 700
	ClipBoard := SubPat2
	sleep 200
	send ^v

	;Replaces Clipoard with third value of table
	MouseClick, left, 808, 844
	sleep 700
	MouseClick, left, 808, 844
	sleep 700
	ClipBoard := SubPat3
	sleep 200
	send ^v

	;Replaces Clipoard with fourth value of table
	MouseClick, left, 797, 861
	sleep 700
	MouseClick, left, 797, 861
	sleep 700
	ClipBoard := SubPat4
	sleep 200
	send ^v
}
else if (Var = "service")
{
	CoordMode, Mouse, Screen
	;send {ctrl down}{shift down}i{ctrl up}{shift up}
	MouseClick, right, 1518, 957
	send {UP}
	send {ENTER}
	Sleep 1000 ;to wait for inspect code to appear
	MouseClick, left, 1484, 1533
	send {Right}
	sleep 200
	send {Down}
	sleep 200
	send {Right}
	sleep 200
	send {Down}
	sleep 200
	send {Down}
	sleep 200
	;Add more code
	send ^c
	;WinActivate ; use the window found above
	sleep 1000
	MouseClick, left, 775, 878
	Sleep 500
	MouseClick, left
	sleep 1000
	temp := Clipboard

	StartPos := RegExMatch(ClipBoard, "Total:</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""15"" aria-describedby=""list1_Assigned"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""83"" aria-describedby=""list1_In Progress"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""111"" aria-describedby=""list1_Pending"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""769"" aria-describedby=""list1_Resolved"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""978"" aria-describedby=""list1_Total"">978</td></tr></tbody></table>", Req)
	sleep 200


	;Replaces Clipoard with first value of table
	ClipBoard := Req1
	send ^v
	sleep 1000

	;Replaces Clipoard with second value of table
	MouseClick, left, 778, 896
	sleep 700
	MouseClick, left, 778, 896
	sleep 700
	ClipBoard := Req2
	sleep 200
	send ^v

	;Replaces Clipoard with third value of table
	MouseClick, left, 779, 914
	sleep 700
	MouseClick, left, 779, 914
	sleep 700
	ClipBoard := Req3
	sleep 200
	send ^v

	;Replaces Clipoard with fourth value of table
	MouseClick, left, 781, 929
	sleep 700
	MouseClick, left, 781, 929
	sleep 700
	ClipBoard := Req4
	sleep 200
	send ^v
}
else
{
	MsgBox, Incorrect categorization
}
