#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

^PrintScreen:: RunWait, %A_WinDir%\system32\SnippingTool.exe

::ccc::C760912{TAB}Start4321{ENTER}

::rrr::C760912{TAB}Rus2000{!}{ENTER}

::mmm::art.z{@}tcs.com{TAB}Mex0030{!}{ENTER}

::general1::
send Name User: 
send {enter}{enter}
send User: 
send {enter}{enter}
send Computer Name: 
send {enter}{enter}
send Does the user speaks in English: No
send {enter}{enter}
send Callback Number: 
send {enter}{enter}
send Application Name: 
send {enter}{enter}
send Error Message / Problem Description: 
send {enter}{enter}
send When the problem started: 
Send, %A_DD%-%A_MM%-%A_YYYY% %A_Hour%:%A_Min%
send {enter}{enter}
send Is the user the only one with the problem: Yes
send {enter}{enter}
send Location: 
send {enter}{enter}
send Troubleshoot: 
return

::uuu::1248463{TAB}Mex0030{!}{ENTER}

::sss::
send, INFRA-SD-EMEA-SPANISH
sleep 300
send, {DOWN}
sleep 300
send, {ENTER}
sleep 300
send, {TAB 2}
send, CARLOS A
sleep 300
send, {DOWN}
sleep 100
send, {ENTER}
return

::ddd::
send %A_MM%{/}%A_DD%{/}%A_YYYY% 
return

;EMEA_ACDADMIN@CARGILL.ngcc.bt.com 
;tcs#1234
;http://drws00.ngcc.bt.com/webadministrator/ 

::yyy::  
	;today = %a_now%
	datestring = %a_now%
	StringSplit, d, datestring, /
	date := d3 d1 d2
	FormatTime, day_of_Week, %date%, ddd
	dayName = %day_of_week%

	FormatTime, day_of_Week, %date%, d
	day = %day_of_week%

	if (dayName = "Sun")
		datestring += -2, days
	else if (dayName = "Mon")
		datestring += -3, days
	else if (dayName = "Tue")
		datestring += -1, days
	else if (dayName = "Wed")
		datestring += -1, days
	else if (dayName = "Thu")
		datestring += -1, days
	else if (dayName = "Fri")
		datestring += -1, days
	else if (dayName = "Sat")
		datestring += -1, days

;FormatTime, today, %today%, MM/dd/yy 
FormatTime, datestring, %datestring%, MM/dd/yy
SendInput %datestring%
return

#r::
run, explorer C:\Users\%A_UserName%\Documents\Reports\Daily\2017
WinActivate, Desktop
return

#m::
run, explorer C:\Users\%A_UserName%\Documents\Reports\Monthly\2017
WinActivate, Desktop
return

#e::
run, explorer C:\Users\%A_UserName%\Desktop\
WinActivate, Desktop
return

::btt::BUD003{@}cargill{.}ngcc{.}bt{.}com{TAB}tcs{#}1234{ENTER}


Shift & WheelUp:: 
	; Scroll to the left
	MouseGetPos,,,id, fcontrol,1
	Loop 8 ; <-- Increase for faster scrolling
		SendMessage, 0x114, 0, 0, %fcontrol%, ahk_id %id% ; 0x114 is WM_HSCROLL and the 0 after it is SB_LINERIGHT.
return
 
Shift & WheelDown:: 
	;Scroll to the right
	MouseGetPos,,,id, fcontrol,1
	Loop 8 ; <-- Increase for faster scrolling
		SendMessage, 0x114, 1, 0, %fcontrol%, ahk_id %id% ;  0x114 is WM_HSCROLL and the 1 after it is SB_LINELEFT.
return



;~Shift::
;400 is the maximum allowed delay (in milliseconds) between presses.
;if (A_PriorHotKey = "~Ctrl" AND A_TimeSincePriorHotkey < 400)
;{
;   Run, Notepad "C:\Users\c760912\Desktop\Scripts\AutoFill2.0.ahk"
;}
;Sleep 0
;KeyWait Ctrl
;return


