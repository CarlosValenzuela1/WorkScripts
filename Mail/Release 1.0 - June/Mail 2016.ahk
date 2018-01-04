#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force
#Persistent
;#NoTrayIcon

;Scrolllock::Pause

;"C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" /c ipm.note




Pause::
	;Gui, Add, Text, x319 y93 w0 h249, Testing
	Gui, -Caption
	Gui, Add, Text, center x130, What type of request do you need?
	;Gui, Add, DropDownList, x29 y30 w364 h90 +Center vchoice gOnChoice, Holiday|Shift Change|Lateness/Break Overpull
	Gui, Add, DropDownList, x29 y30 w364 h90 +Center vchoice gOnChoice, Holiday|Shift Change
	Gui, Add, Button, Default x40 y90 w96 h28 gOK , OK
	Gui, Add, Button, x290 y90 w96 h28 gCancel, Cancel
	Gui, Show, w428 h130, Shift Requestor
	return

	GuiClose:
	;ExitApp
	Reload
Return


OnChoice:
	;Gui, Submit, nohide ;To get value from Dropdown menu
	;MsgBox % choice
return

Cancel:
	;ExitApp
	Reload
return

OK:
	Gui, Submit, nohide
	Gui, Destroy
	if(choice = "Holiday")
	{
		gosub HolidayGui
	} 
	else if (choice = "Shift Change")
	{
		gosub ShiftGui
	}
	/*	
	else if (choice = "Lateness/Break Overpull")
	{
		gosub LatenessGui
	}
	*/
return




HolidayGui:
	Gui, -Caption
	Gui, Add, Text, x110 y10 w326 h57 ,Please select the date for the Holiday Request.
	Gui, Add, DateTime, x100 y30 w240 h28 vdateh1 gDateh1, 
	
	Gui, Add, Text, x30 y60 w450 h57 ,In case of multiple holidays please continue selecting the date on the fields below.
	Gui, Add, DateTime, x100 y80 w240 h28 vdateh2 gDateh2, 
	Gui, Add, DateTime, x100 y130 w240 h28 vdateh3 gDateh3, 
	Gui, Add, DateTime, x100 y180 w240 h28 vdateh4 gDateh4, 
	Gui, Add, DateTime, x100 y230 w240 h28 vdateh5 gDateh5, 
	Gui, Show, w450 h350, Holidays
	Gui, Add, Button, x40 y300 w96 h28 gOKHoliday , OK
	Gui, Add, Button, x290 y300 w96 h28 gCancel, Cancel
return


ShiftGui:
	Gui, -Caption
	Gui, Add, Text, x120 y10 w326 h57 , Please select the date for the Shift Change.
	Gui, Add, DateTime, x100 y30 w240 h28 vdate1 gDate1, 
	Gui, Add, Text, x50 y80 w326 h57 , Type your coworker's LASTNAME and their shift.
	GUI, Add, Edit, x50 y100 w201 h19 r1 Lowercase vperson1 gPerson1, 
	Gui, Add, DropDownList, x271 y100 w96 h100 vshift1 gShift1, 6:00 - 15:00|7:00 - 16:00|8:00 - 17:00|9:00 - 18:00|10:00 - 19:00
	Gui, Add, Text, x75 y150 w316 h48 , Please type your LASTNAME and your CURRENT shift time.
	GUI, Add, Edit, x50 y170 w201 h19 r1 Lowercase vperson2 gPerson2, 
	Gui, Add, DropDownList, x271 y170 w96 h100 vshift2 gShift2, 6:00 - 15:00|7:00 - 16:00|8:00 - 17:00|9:00 - 18:00|10:00 - 19:00
	Gui, Show, w416 h276, Shift Changes
	Gui, Add, Button, x40 y220 w96 h28 gOKShift , OK
	Gui, Add, Button, x290 y220 w96 h28 gCancel, Cancel


return

LatenessGui:
	Msgbox Latenessgui
return





;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;01/20/2017
OKShift:
	;Msgbox % "Time is: " date1 
	;Msgbox % "Year is: " year 
	;MsgBox % "month: " month
	;MsgBox % "day: " day
	DetectHiddenWindows On
	run, "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE" /c ipm.note"
	WinWaitActive, Untitled - Message (HTML) , , 10,
	sendInput ALEXANDER_ENGWERDA@CRGL-THIRDPARTY.COM; ANA_SERCEVSKA@CRGL-THIRDPARTY.COM; ANDRAS_TOTH@CRGL-THIRDPARTY.COM
	send, {tab 3}
	sendInput, Shift Change Request - %month%/%day%/%year%
	send, {tab}
	sendInput, Person: %person1%`nShift %shift2% `nPerson: %person2%`nShift: %shift1%

return

OKHoliday:
	;msgbox date1: %monthh1% day: %dayh1% year: %yearh1%
	;msgbox date2: %monthh2% day: %dayh2% year: %yearh2%
	;msgbox date3: %monthh3% day: %dayh3% year: %yearh3%
	;msgbox date4: %monthh4% day: %dayh4% year: %yearh4%
	;msgbox date5: %monthh5% day: %dayh5% year: %yearh5%
	DetectHiddenWindows On
	run, "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE" /c ipm.note"
	WinWaitActive, Untitled - Message (HTML) , , 10,
	sendInput ALEXANDER_ENGWERDA@CRGL-THIRDPARTY.COM; ANA_SERCEVSKA@CRGL-THIRDPARTY.COM; ANDRAS_TOTH@CRGL-THIRDPARTY.COM
	send, {tab 3}
	sendInput, Holiday Request
	send, {tab}
	if monthh1
		sendInput, Holiday: %monthh1%/%dayh1%/%yearh1%`n
	if monthh2
	{
		;msgbox date2: %monthh2% day: %dayh2% year: %yearh2%
		sendInput, Holiday: %monthh2%/%dayh2%/%yearh2%`n
	}
	if monthh3
		sendInput, Holiday: %monthh3%/%dayh3%/%yearh3%`n
	if monthh4
		sendInput, Holiday: %monthh4%/%dayh4%/%yearh4%`n
	if monthh5
		sendInput, Holiday: %monthh5%/%dayh5%/%yearh5%`n
return

Person1:
	Gui, Submit, nohide
return


Person2:
	Gui, Submit, nohide
return


Shift1:
	Gui, Submit, nohide
return


Shift2:
	Gui, Submit, nohide
return


Date1:
	Gui, Submit, nohide
	;date1 := SubStr( date1, 1, 8 )
	;date1 := SubStr( date1, 1, 8 )
	;newDate1 := StringSplit, d, date1, /
	year := SubStr(date1, 1, 4)
	month := SubStr(date1, 5, 2)
	day := SubStr(date1, 7, 2)
	sleep 1000
return

Dateh1:
	Gui, Submit, nohide
	yearh1 := SubStr(dateh1, 1, 4)
	monthh1 := SubStr(dateh1, 5, 2)
	dayh1 := SubStr(dateh1, 7, 2)
	sleep 1000
return

Dateh2:
	Gui, Submit, nohide
	yearh2 := SubStr(dateh2, 1, 4)
	monthh2 := SubStr(dateh2, 5, 2)
	dayh2 := SubStr(dateh2, 7, 2)
	sleep 1000
return


Dateh3:
	Gui, Submit, nohide
	yearh3 := SubStr(dateh3, 1, 4)
	monthh3 := SubStr(dateh3, 5, 2)
	dayh3 := SubStr(dateh3, 7, 2)
	sleep 1000
return


Dateh4:
	Gui, Submit, nohide
	yearh4 := SubStr(dateh4, 1, 4)
	monthh4 := SubStr(dateh4, 5, 2)
	dayh4 := SubStr(dateh4, 7, 2)
	sleep 1000
return


Dateh5:
	Gui, Submit, nohide
	yearh5 := SubStr(dateh5, 1, 4)
	monthh5 := SubStr(dateh5, 5, 2)
	dayh5 := SubStr(dateh5, 7, 2)
	sleep 1000
return



/*
GUI 1

#NoTrayIcon

Gui, Add, Text, x319 y93 w0 h249 , Text
Gui, Add, DropDownList, x29 y84 w364 h19 +Center, DropDownList
Gui, Add, Button, x40 y151 w96 h28 , OK
Gui, Add, Button, x290 y151 w96 h28 , Cancel
Gui, Add, Text, x58 y16 w307 h19 +Horz +Center, Please select the type of request:
; Generated using SmartGUI Creator for SciTE
Gui, Show, w428 h215, Untitled GUI
return

GuiClose:
ExitApp




GUI 2

#NoTrayIcon

Gui, Add, Text, x43 y7 w326 h57 , Text
Gui, Add, DateTime, x50 y84 w240 h28 , 
Gui, Add, Button, x309 y84 w48 h28 , Button
Gui, Add, DropDownList, x194 y-50 w153 h86 , DropDownList
Gui, Add, ComboBox, x50 y132 w201 h19 , ComboBox
Gui, Add, ComboBox, x271 y132 w96 h9 , ComboBox
Gui, Add, Text, x50 y170 w316 h48 , Text
Gui, Add, ComboBox, x50 y228 w201 h19 , ComboBox
Gui, Add, ComboBox, x271 y228 w96 h19 , ComboBox
; Generated using SmartGUI Creator for SciTE
Gui, Show, w416 h276, Untitled GUI
return

GuiClose:
ExitApp



GUI 3

Gui, Add, Text, x40 y7 w297 h28 , Text
Gui, Add, DateTime, x40 y45 w230 h28 , 
Gui, Add, Button, x290 y46 w48 h28 , +
Gui, Add, Button, x31 y304 w96 h28 , OK
Gui, Add, Button, x242 y304 w96 h28 , Cancel
Gui, Add, DateTime, x40 y93 w230 h28 , 
Gui, Add, Button, x290 y93 w48 h28 , +
Gui, Add, DateTime, x40 y141 w230 h28 , 
Gui, Add, Button, x290 y141 w48 h28 , +
Gui, Add, DateTime, x40 y189 w230 h28 , 
Gui, Add, Button, x290 y189 w48 h28 , +
; Generated using SmartGUI Creator for SciTE
Gui, Show, w381 h377, Untitled GUI
return

GuiClose:
ExitApp


ORIGINAL

Gui, +AlwaysOnTop
	Gui Show, w300 h300, Shift Requestor
	Gui, Add, Text, center x70, What type of request do you need?
	Gui, Add, DropDownList, center w130 x90  vchoice gOnChoice, Holiday|Shift Change|Lateness/Break Overpull

	Gui, Add, Text, center x70, Please select the date for the Holiday

	Gui, Add, DateTime, w100 vMyDateTime
	Gui, Add, Button, x+50 hp gDateRoutine, Go
	
	;;;;;;;;;Dynamic Code START
	Gui, Add, Button, gNew, +
	Gui, Add, Text, Section ym, Field 1
	Gui, Show,, Test
	;;;;;;;;;Dynami Code END


DateRoutine:
	 Gui, Submit, NoHide
	 MyDateTime := SubStr( MyDateTime, 1, 8 )
	 MsgBox, % MyDateTime

Return


	

OnChoice:
	Gui, Submit, nohide
	MsgBox % choice
return



New:
Gui, Add, Text, xs, New Field
Gui, Show, AutoSize, Test
return


*/

