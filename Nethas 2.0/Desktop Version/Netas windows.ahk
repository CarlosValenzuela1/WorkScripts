;=========================================================================================================;
; Netas Script: This script autogenerates the daily reports; Cargill KPI, EUC Backlog, and Backlog Report ;
; Description of each method will be find below.							  ;
; Version: 5.0												  ;
; Author: Carlos Valenzuela										  ;
; Co-author: Fabricio Costeira										  ;
; Company: TATA Consultancy Services									  ;
;=========================================================================================================;


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force
#Persistent
#MaxThreadsPerHotkey 2

Scrolllock::Pause				;It allows script to be pause by pressing the key ScrollLock on keyboard

						;Script will excute when pressing the keys: Ctrl + F1
^F1::
	CoordMode, Mouse, Screen
	mainMethod()

return





mainMethod()
{	
	global
	
	initializeVariables()
	
	;createDirectory()			;Already tested and working

	openNetasthra()

	;==========================================================;
	; Download the Incident Distribution Report from Netasthra ;
	;==========================================================;
	openIncidentDistribution()
	wasFileDownloaded()
	getIncDisfile()	
	CoordMode, Mouse, Screen			;It only works on current window
	/*
	;=========================================================;
	; Download the Service Distribution Report from Netasthra ;
	;=========================================================;
	openServiceDistribution()
	wasFileDownloaded()
	getSerDisfile()	

	;=======================================================;
	; Download the Missrouted Tickets Report from Netasthra ;
	;=======================================================;
	;openMisroutedTickets()
	sleep 5000

	;===============================================================;
	; Download the Missrouted Service Request Report from Netashtra ;
	;===============================================================;
	openMisroutedService()
	sleep 5000

	;=========================================================;
	; Download EUC Incident Distribution Report from Netashtra;
	;=========================================================;
	openEUCIncident()
	wasFileDownloaded()
	sleep 2000
	getIncDisfileEUC()	
	sleep 5000

	;=========================================================;
	; Download EUC Service Distribution Report from Netashtra ;
	;=========================================================;
	openEUCServiceRequest()
	wasFileDownloaded()
	closeBrowser()
	sleep 2000
	getSerDisfileEUC()	
	sleep 5000

	;======================================================;
	; Gather email count from Outlook for the Mail Backlog ;
	;======================================================;	
	extractMailCount()
	closeOutlook()
	
	;=======================================================;
	; Fill the KPI Document to generate first set of Emails ;
	;=======================================================;	
	fillKPI()
	sleep 1000
		*/
	;=======================================================================;
	; Download the Queue Statistics Report from drwr00.ngcc.bt.com/Reports/ ;
	;=======================================================================;	
	queueStatistic()
	
	;==================================================================;
	; Download the Agent Status Report from drwr00.ngcc.bt.com/Reports/;
	;==================================================================;	
	;agentStatus()
	
	
	;====================================================================;
	; Copy the CSV Files from downloads to todays folder in Daily Report ;
	;====================================================================;
	;getCSVfile()


	;==========================================================;
	; Format and Paste the values from Queue Statistics Report ;
	;==========================================================;
	;formatPasteQueueStatistics()


	;==========================================================;
	; Format and Paste the values from Status Agent Report ;
	;==========================================================;
	;formatPasteStatusAgent()

	
}

initializeVariables()				;Global Variables
{
	global incident1 := -1
	global incident2 := -1
	global incident3 := -1
	global incident4 := -1
	global distribution1 := -1
	global distribution2 := -1
	global distribution3 := -1	
	global distribution4 := -1
	global missticket1 := -1
	global missticket2 := -1
	global missdistribution1 := -1
	global missdistribution2 := -1
	global russianMail = -1			;why was this a comment?
	global polishMail := -1
	global frenchMail := -1
	global romanianMail := -1
	global turkishMail := -1
	global hungarianMail := -1
}

closeBrowser()
{
	SetTitleMatchMode, 2
	WinActivate, google chrome
	sleep 1000
	send !{F4}
}

closeOutlook()
{
	SetTitleMatchMode, 2
	WinActivate, Outlook
	sleep 1000
	send !{F4}
}

formatPasteStatusAgent()
{
	SetTitleMatchMode, 2
	WinActivate, Agent State

	yesterday1 = %A_DD%
	IfEqual, yesterday1, Monday
		yesterday1 -= 03			;Because the yesterday of Monday is Friday                
	Else                                       
		yesterday1 -= 01
	yesterday1 := 0 . yesterday1		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(yesterday1)
	if (length <= 2 & length > 0)
		;MsgBox Less than or equal to 2
		rellenador = 1
	else
	StringTrimLeft, yesterday1, yesterday1, 1 

	send {LWin down}{UP down}{UP up}{LWin up} 
	send ^{Home}{Alt down}at{Alt up}
	send {right 3}{alt down}{down}{alt up}
	send {tab 7}%yesterday1%{ENTER}
	sleep 100
	send {right}{alt down}{down}{alt up}
	sleep 100
	send {tab 8}
	send {space}
	send {down 2}{space}
	send {down 3}{space}{enter}
	sleep 100
	send {right 4}
	send {alt down}{down}{alt up}{tab 7}
	sleep 100
	send initialized{ENTER}
	sleep 100
	send {left 5}{down}{Shift down}{right 5}{ctrl down}{down}{ctrl up}{shift up}
	sleep 100
	send ^c
	sleep 100
	send {Shift down}{F11}{Shift up}
	sleep 100
	send ^v
	sleep 100
	send ^{Home}{right 3}{shift down}{right}{shift up}
	sleep 100
	send {shift down}{ctrl down}{down}{shift up}{ctrl up}
	sleep 100
	send ^-c{enter}
	send ^{Home}
	send +^{right}+^{down}
	sleep 100
	send !a!m{enter}{tab 6}{enter}{enter}
	send ^{Home}+^{right}+^{down}^c
	sleep 2000

	SetTitleMatchMode, 2
	WinActivate, kpi
	sleep 1000
	send {F5}
	sendRaw 'States Source'!A1
	send {ENTER}
	sleep 1000
	send ^{down}
	sleep 1000
	send {down}
	sleep 2000
	send ^v
	sleep 1200
	send {PgDn}
	sleep 1000
	send {F5}
	sendRaw 'EMEA Charts'!A76
	send {ENTER}{PgDn}
	sleep 500
	send {PgDn}
	sleep 500
	send {F5}
	sendRaw Mailing!A1
	send {ENTER}{PgDn}{PgDn}
	sleep 500
	CoordMode, Mouse, Relative
	MouseClick left, 984, 432
	
	



}

formatPasteQueueStatistics()
{
	SetTitleMatchMode, 2
	WinActivate, Queue Statistics
	send {LWin down}{UP down}{UP up}{LWin up} 
	send ^{Home}{Alt down}at{Alt up}
	send {alt down}{down}{alt up}
	send {tab 7}EMEA{enter}
	send {RIGHT 18}{alt down}{down}{alt up}{tab 8}{SPACE}{DOWN}{SPACE}{DOWN}{SPACE}{ENTER}
	send ^{Home}{down}
	send {Shift down}{RCtrl down}{Right 4}{Down}{RCtrl up}{Shift up}
	sleep 200
	send ^c
	sleep 2000
	send {alt down}{F4}{alt up}
	send {alt down}n{alt up}
	send {alt down}y{alt up}
	sleep 200

	SetTitleMatchMode, 2
	WinActivate, kpi
	sleep 100
	send {F5}
	send Source{!}A1
	send {ENTER}
	send {ctrl down}{down 4}{ctrl up}{down}
	sleep 200
	send ^v
	
}

agentStatus()
{
	run, chrome.exe
	sleep 500
	send {LWin down}{UP down}{UP up}{LWin up} 
	sleep 500
	MouseClick left, 1982, 8
	sleep 100
	MouseClick left, 1810, 285
	sleep 200
	MouseClick left, 1373, 1143
	sleep 200
	send ^t
	send  https://drwr00.ngcc.bt.com/Reports/Pages/Folder.aspx?ItemPath=/CARGILL{ENTER}
	sleep 2000
	send BUD003@cargill.ngcc.bt.com
	send {TAB up}{Tab down}    
	send tcs{#}1234
	send {ENTER}
	sleep 4000
	send {Tab 9}
	send {ENTER}
	sleep 500



	MouseMove, 1073, 264
	waitColorChange()
	send {TAB 12}				;Move to Queue Statistics
	send {ENTER}
	waitColorChange()
	sleep 5000
	yesterday = %A_DD%
	IfEqual, yesterday, Monday
		yesterday -= 03			;Because the yesterday of Monday is Friday                
	Else                                       
		yesterday -= 01
	yesterday := 0 . yesterday		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(yesterday)
	if (length <= 2 & length > 0)
		;MsgBox Less than or equal to 2
		insta = 1
	else
	StringTrimLeft, yesterday, yesterday, 1 
	send {TAB 8}
	sleep 200
	Send, %A_MM%/%yesterday%/%A_YYYY%{TAB}

	sleep 3000
	send {TAB 2}%A_MM%/%yesterday%/%A_YYYY%{TAB}

	sleep 3000
	send {TAB 4}{space}
	sleep 3000
	send {TAB 6}{space}
	send 60
	send {ENTER}

	sleep 5000
	send {TAB 6}
	send 60
	send {ENTER}
	sleep 5000
	send {TAB 4}{SPACE}
	send 60
	send {ENTER}

	sleep 3000
	send {TAB 9}
	send {ENTER}
	sleep 10000
	send {TAB 18}
	sleep 1000
	send {ENTER}
	sleep 1000
	send {DOWN}
	sleep 200
	send {ENTER}

}



queueStatistic()
{
	run, chrome.exe
	sleep 500
	send {LWin down}{UP down}{UP up}{LWin up} 
	sleep 500
	MouseClick left, 1982, 8
	sleep 100
	MouseClick left, 1810, 285
	sleep 200
	MouseClick left, 1373, 1143
	sleep 200
	send ^t
	send  https://drwr00.ngcc.bt.com/Reports/Pages/Folder.aspx?ItemPath=/CARGILL{ENTER}
	sleep 2000
	send BUD003@cargill.ngcc.bt.com
	send {TAB up}{Tab down}    
	send tcs{#}1234
	send {ENTER}
	sleep 1000
	send {TAB}
	send {Tab 11}
	send {ENTER}
	sleep 500
	;MsgBox testing movemont of mouse
	MouseMove, 1073, 264
	waitColorChange()
	send {TAB 12}				;Move to Queue Statistics
	send {ENTER}
	waitColorChange()
	sleep 5000
	yesterday = %A_DD%
	IfEqual, yesterday, Monday
		yesterday -= 03			;Because the yesterday of Monday is Friday                
	Else                                       
		yesterday -= 01
	yesterday := 0 . yesterday		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(yesterday)
	if (length <= 2 & length > 0)
		;MsgBox Less than or equal to 2
		insta = 1
	else
	StringTrimLeft, yesterday, yesterday, 1 
	send {TAB 8}
	sleep 200
	Send, %A_MM%/%yesterday%/%A_YYYY%
	send {TAB}
	sleep 3000
	send {TAB 2}%A_MM%/%yesterday%/%A_YYYY%{TAB}
	sleep 3000
	send {TAB 4}{DOWN}{ENTER}
	sleep 3000
	send {TAB 5}
	send 60
	send {ENTER}
	sleep 5000
	send {TAB 6}
	send 60
	send {ENTER}
	sleep 5000
	send {TAB 7}
	send 60
	send {ENTER}
	sleep 3000
	send {TAB 10}
	send {ENTER}
	sleep 10000
	send {TAB 18}
	sleep 1000
	send {ENTER}
	sleep 1000
	send {DOWN}
	sleep 200
	send {ENTER}
	sleep 4000
	send ^{F4}
}


openNetasthra()
{

	run, chrome.exe http://itreporting.cargill.com/Netasthra/loginform
	send {LWin down}{UP down}{UP up}{LWin up}						;Maximizing window
	MouseMove 1553, 413
	waitColorChange()
	sleep 3000
	send {TAB}
	
	send, c760912{tab}Pear1234{enter}
	waitColorChange()
	sleep 2500
	CoordMode, Mouse, Screen			;It only works on current window
}

createDirectory()
{
	send {LWin down}e{LWin up} 
	sleep 600
	send {LWin down}{UP down}{UP up}{LWin up}		;Maximizing window
	sleep 600
	send {LAlt down}d{LAlt up}
	sleep 600
	csvMonth = %A_MMMM%
	csvDay = %A_DD%
	csvDay := 0 . csvDay		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(csvDay)	;If variable has 3 char means number > 9 for example, day 024, instead of 24
	if (length <= 2 & length > 0)
		;MsgBox Warning! Lenght of current day is less than 0 or greater than or equal to 3
		testo = 1
	else
		StringTrimLeft, csvDay, csvDay, 1

	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\
	send {ENTER}
	sleep 600
	send {Ctrl down}{Shift down}n{Ctrl up}{Shift up}
	sleep 500
	send %csvDay%_%csvMonth%
	send {ENTER}
	sleep 600
	send {LAlt down}d{LAlt up}
	sleep 600
	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%
	send {ENTER}
	sleep 600
	send {Ctrl down}{Shift down}n{Ctrl up}{Shift up}
	sleep 600
	send EUC Report - %csvDay%_%csvMonth%
	send {ENTER}
	send {Alt down}{F4 down}{Alt up}{F4 up}
	sleep 900
}


getIncDisfile()
{
	send {LWin down}e{LWin up} 
	sleep 500
	send {LWin down}{UP down}{UP up}{LWin up}		;Maximizing window
	sleep 500
	send {LAlt down}d{LAlt up}
	sleep 500
	send C:\Users\c760912\Downloads
	sleep 800
	send {ENTER}
	sleep 500
	send {F3}
	sleep 500
	sendRaw Incident_Distribution
	sleep 500
	send {Enter}
	sleep 500
	send {TAB 3}
	sleep 1000
	send ^a
	sleep 3000
	send ^x
	sleep 3000
	send {LAlt down}d{LAlt up}
	csvMonth = %A_MMMM%
	csvDay = %A_DD%
	shortDay := 0 . csvDay		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(shortDay)
	if (length <= 2 & length > 0)
		;MsgBox Less than or equal to 2
		testo = 1
	else
		StringTrimLeft, shortDay, shortDay, 1
	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%
	sleep 900
	send {Enter}
	sleep 500
	MouseClick left, 522, 1240
	sleep 3000
	send ^v
	sleep 900
	send {Alt down}{F4 down}{Alt up}{F4 up}
	sleep 500
}

getSerDisfile()
{
	send {LWin down}e{LWin up} 
	sleep 500
	send {LWin down}{UP down}{UP up}{LWin up}		;Maximizing window
	sleep 500
	send {LAlt down}d{LAlt up}
	sleep 500
	send C:\Users\c760912\Downloads
	sleep 500
	send {ENTER}
	sleep 500
	send {F3}
	sleep 500
	sendRaw Service_Request
	sleep 500
	send {Enter}
	sleep 500
	send {TAB 3}
	sleep 1000
	send ^a
	sleep 3000
	send ^x
	sleep 3000
	send {LAlt down}d{LAlt up}
	csvMonth = %A_MMMM%
	csvDay = %A_DD%
	shortDay := 0 . csvDay		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(shortDay)
	if (length <= 2 & length > 0)
		;MsgBox Less than or equal to 2
		testi = 1
	else
		StringTrimLeft, shortDay, shortDay, 1

	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%
	sleep 500
	send {Enter}
	sleep 500
	MouseClick left, 522, 1240
	sleep 3000
	send ^v
	sleep 500
	send {Alt down}{F4 down}{Alt up}{F4 up}
	sleep 500
}


getIncDisfileEUC()
{
	send {LWin down}e{LWin up} 
	sleep 500
	send {LWin down}{UP down}{UP up}{LWin up}		;Maximizing window
	sleep 500
	send {LAlt down}d{LAlt up}
	sleep 500
	send C:\Users\c760912\Downloads
	sleep 500
	send {ENTER}
	sleep 500
	send {F3}
	sleep 500
	sendRaw Incident_Distribution
	sleep 500
	send {Enter}
	sleep 500
	send {TAB 3}
	sleep 500
	send ^a
	sleep 1000
	send ^x
	sleep 1000
	send {LAlt down}d{LAlt up}
	csvMonth = %A_MMMM%
	csvDay = %A_DD%
	shortDay := 0 . csvDay		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(shortDay)
	if (length <= 2 & length > 0)
		;MsgBox Less than or equal to 2
		testp = 1
	else
		StringTrimLeft, shortDay, shortDay, 1
	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%\EUC Report - %csvDay%_%csvMonth%
	sleep 500
	send {Enter}
	sleep 500
	MouseClick left, 522, 1240
	sleep 500
	send ^v
	sleep 500
	send {Alt down}{F4 down}{Alt up}{F4 up}
	sleep 500
}

getSerDisfileEUC()
{
	send {LWin down}e{LWin up} 
	sleep 500
	send {LWin down}{UP down}{UP up}{LWin up}		;Maximizing window
	sleep 500
	send {LAlt down}d{LAlt up}
	sleep 500
	send C:\Users\c760912\Downloads
	sleep 500
	send {ENTER}
	sleep 500
	send {F3}
	sleep 500
	sendRaw Service_Request
	sleep 500
	send {Enter}
	sleep 500
	send {TAB 3}
	sleep 500
	send ^a
	sleep 1000
	send ^x
	sleep 1000
	send {LAlt down}d{LAlt up}
	csvMonth = %A_MMMM%
	csvDay = %A_DD%
	shortDay := 0 . csvDay		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(shortDay)
	if (length <= 2 & length > 0)
		;MsgBox Less than or equal to 2
		magickarp = 1
	else
		StringTrimLeft, shortDay, shortDay, 1

	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%\EUC Report - %csvDay%_%csvMonth%
	sleep 500
	send {Enter}
	sleep 500
	MouseClick left, 522, 1240
	sleep 500
	send ^v
	sleep 500
	send {Alt down}{F4 down}{Alt up}{F4 up}
	sleep 500
}



getCSVfile()
{
	send {LWin down}e{LWin up} 
	sleep 200
	send {LWin down}{UP down}{UP up}{LWin up}		;Maximizing window
	sleep 200
	send {LAlt down}d{LAlt up}
	sleep 200
	send C:\Users\c760912\Downloads
	sleep 200
	send {ENTER}
	send {TAB}
	sendRaw *.csv
	send {Enter}
	sleep 1000
	send {TAB 3}
	send ^a
	send ^x
	send {LAlt down}d{LAlt up}
	
	csvMonth = %A_MMMM%
	csvDay = %A_DD%
	shortDay := 0 . csvDay		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(shortDay)
	if (length <= 2 & length > 0)
		;MsgBox Less than or equal to 2
		ma = 1
	else
		StringTrimLeft, shortDay, shortDay, 1

	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%
	send {Enter}
	sleep 100
	send {Tab 5}
	sleep 200
	send ^v
	sleep 2000
	send {ENTER}
	;send {Alt down}{F4 down}{Alt up}{F4 up}
	;sleep 900
}


waitColorChange()
{
	CoordMode, Mouse, Relative				;It only works on current window
	found := 0						;Initializes the variable found
	loopTimes := 0						;Initializes the variable loopTimes
	MouseGetPos Xpos, Ypos					;Obtains mouse position at that time
	PixelGetColor firstColor, %Xpos%, %Ypos%, RGB		;Gets the color of the pixel where mouse is at
	secondColor := 1					;Initializes the variable secondColor
	while(found == 0)				;Check if the color has been found, if not then it continues
	{
		sleep 400					;Waits 400 miliseconds for animations that haven't finished
		if(secondColor != firstColor && loopTimes != 0)
			found := 1
	else
	{
		MouseGetPos Xpos, Ypos				;Gets position of mouse
		PixelGetColor secondColor, %Xpos%, %Ypos%, RGB	;gets the color of the pixel where mouse is at
		loopTimes := (loopTimes + 1)			;Increases the times it has checked for a color change
	}
	cliboard := 						;NOT SURE IF ITS NEEDED
	}
	CoordMode, Mouse, Screen				;Returns coordinates to both monitors
}


openIncidentDistribution()
{
	sleep 600
	MouseMove, 1442, 132				;Hovers over Reports
	sleep 600
	MouseMove, 1467, 193				;Hovers over Volumetric Reports
	sleep 600
	MouseMove, 1631, 242				;Hovers over Incident
	sleep 600
	MouseClick, left, 1786, 245			;Clicks in Incident Distribution
	sleep 600
	sleep 300					;Waits 300 miliseconds to move the mouse to wait for color change
	MouseMove, 1616, 779				;Hovers over loading image
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	MouseClick, left, 1208, 170 			;It clicks on the settings wheel
	sleep 600
	MouseClick, left, 1685, 822 			;It clicks on the dropdown menu
	sleep 600
	MouseClick, left, 1648, 867 			;Select SD OPEN REPORT
	sleep 600
	MouseClick, left, 1578, 890 			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	waitColorChange()				;Wait for color change
	MouseClick, left, 2070, 253			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 1613, 676			;Clicks on the date field
	sleep 500					;Waits for popup to appear
	MouseClick, left, 1622, 660			;Clicks on next month
	sleep 500					;waits for month to change on screen

	PixelSearch, Xpos, Ypos, 413, 706, 610, 879, 0xF40F0F, 3, Fast RGB
	sleep 300					;Look for the first red color (Day of today)
	CoordMode, Mouse, Relative			;Changes focus to active window (To use the prev coor to select today)
	MouseClick, left, Xpos, Ypos			;Clicks on day of today
	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 1585, 1119			;Click in Done
	sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 1118, 170			;Clicks on Table icon (Upper-left side on Netasthra)
	sleep 2000 Maybe remove this			;Wait 10 seconds for table to appear
	MouseClick, left, 1534, 1358			;Clicks on position that is empty to be able to use next shortcuts
	sleep 300
	send ^a						;Select All text on webpage
	sleep 700
	send ^c						;Copy All text on webpage
	sleep 400
	StartPos := RegExMatch(ClipBoard, "Total:\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})", Inc)
							;Find the values after "Total: " and stores them in Inc array
	sleep 500
	MouseClick, left, 1969, 1380			;Scroll horizontally to the right
	sleep 500
	MouseClick, left, 1992, 390 			;Clicking on the Total cell of the Incident Distribution table.
	MouseMove, 1583, 938				;Moving mouse to activate color change
	waitColorChange()
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 1000
	MouseClick, left, 1131, 220			;Click on the Eye (add/remove columns)
	sleep 7000 					;Maybe delete this
	send {TAB 3}
	sleep 1000
	send category					;Types CATEGORY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send type					;Types TYPE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send item					;Types ITEM
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Created By					;Types CREATED BY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Solved Date				;Types SOLVED DATE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Closed Date				;Types CLOSED DATE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Source					;Types SOURCE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Country					;Types COUNTRY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Location					;Types LOCATION
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300	
	MouseClick, left, 1519, 1058
	MouseClick, left, 1090, 219			;It clicks on download file

	global incident1 = Inc1
	global incident2 := Inc2
	global incident3 := Inc3
	global incident4 := Inc4
}

openServiceDistribution()
{
	SetTitleMatchMode, 2
	WinActivate, Google Chrome
	sleep 600
	MouseMove, 1442, 140				;Goes to Service Request distribution option
	sleep 600
	MouseMove, 1467, 193
	sleep 600
	MouseMove, 1609, 190
	sleep 600
        MouseMove, 1645, 345	
	sleep 600
	MouseClick, left, 1808, 351
	sleep 300
	MouseMove, 1616, 779
	waitColorChange()
	waitColorChange()
	MouseClick, left, 1208, 170 			;It clicks on the settings wheel
	sleep 600
	MouseClick, left, 1685, 822 			;It clicks on the dropdown menu
	sleep 600
	MouseClick, left, 1648, 867 			;Select SD OPEN REPORT
	sleep 600
	MouseClick, left, 1578, 890 			;Click in Okay option.
	sleep 100
	waitColorChange()
	MouseClick, left, 2070, 253
	sleep 500
	send {TAB 4}
	sleep 100
	send {down}
	MouseClick, left, 1613, 676
	sleep 500
	MouseClick, left, 1622, 660
	sleep 500
	
	PixelSearch, Xpos, Ypos, 413, 706, 610, 879, 0xF40F0F, 3, Fast RGB
	sleep 300
	CoordMode, Mouse, Relative
	MouseClick, left, Xpos, Ypos
	CoordMode, Mouse, Screen
	MouseClick, left, 1585, 1119
	sleep 1000
	waitColorChange()
	MouseClick, left, 1118, 170
	sleep 4000 Maybe remove this
	MouseClick, left, 1534, 1358
	sleep 300
	send ^a
	sleep 700
	send ^c
	sleep 400
	StartPos := RegExMatch(ClipBoard, "Total:\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})", Dis)
	
	MouseClick, left, 1969, 1380
	MouseClick, left, 1918, 417 ;Clicking on the Total cell of the Incident Distribution table.
	MouseMove, 1583, 938
	waitColorChange()
	waitColorChange() ;Maybe it requires two color changes if speed is fast
	sleep 1000
	MouseClick, left, 1131, 220
	sleep 7000 
	send {TAB 3}
	sleep 1000
	send category					;Types CATEGORY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send type					;Types TYPE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send item					;Types ITEM
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Created By					;Types CREATED BY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Solved Date				;Types SOLVED DATE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Closed Date				;Types CLOSED DATE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Source					;Types SOURCE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Country					;Types COUNTRY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Location					;Types LOCATION
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300	
	MouseClick, left, 1519, 1058
	MouseClick, left, 1090, 219
	;MsgBox Dis: %Dis1% Dis: %Dis2% Dis: %Dis3% Dis: %Dis4%
	global distribution1 := Dis1
	global distribution2 := Dis2
	global distribution3 := Dis3
	global distribution4 := Dis4	
}




openMisroutedTickets()
{
	SetTitleMatchMode, 2
	WinActivate, Google Chrome
	sleep 600
	MouseMove, 1442, 132 			;Hover mouse on Reports tab
	sleep 600
	MouseMove, 1428, 161 			;Select option Operational Reports
	sleep 600
	MouseMove, 1626, 161 			;Move mouse to the first option of the Operational Reports menu (change)
	sleep 600
	MouseMove, 1614, 215			;Hover mouse on Incident
	sleep 600
	MouseMove, 1614, 215			;Hover mouse on first option of Incident (Incident Aging)
	sleep 600
	MouseMove, 1825, 221
	sleep 600
	MouseClick, left, 1791, 448		;Click in Misrouted Incidents option.
	sleep 600
	MouseMove, 1616, 779			;Hover mouse on loading image
	waitColorChange()
	waitColorChange()
	MouseClick, left, 1208, 170 		;It clicks on the settings wheel
	sleep 600
	MouseClick, left, 1685, 822 		;It clicks on the dropdown menu
	sleep 600
	MouseClick, left, 1624, 866 		;HERE I AM SELECTING Misrouted report
	sleep 600
	MouseClick, left, 1578, 890 		;I clicking in Okay option.
	sleep 100
	waitColorChange()
	MouseClick, left, 2070, 253		;Click in filter option
	sleep 500
	MouseClick, left, 1613, 676		;Click on Date field
	sleep 500
	MouseClick, left, 1622, 660		;Go to the next month
	sleep 500
						;Following code clicks on first day of the month
	
	PixelSearch, Xpos, Ypos, 413, 706, 612, 738, 0xA7CCE9, 3, Fast RGB		;Coordinates ARE relative in SPY
	sleep 300
	CoordMode, Mouse, Relative		;Changing Relative to click on same coordinates received.
	MouseClick, left, Xpos, Ypos
	CoordMode, Mouse, Screen		;Return coordinate to both desktops, not only active window
	MouseClick, left, 1585, 1119		;Click in Done Button
	sleep 1000
	waitColorChange()
	MouseClick, left, 1118, 170		;Click in the Table icon
	sleep 10000 				;Sleep 10 seconds since waitColorChange is very volatile
	
	MouseClick, left, 1534, 1358		;Click anywhere on the screen (NOT ON TABLE)
	sleep 300
	send ^a					;Select all the text found
	sleep 700
	send ^c					;Copy all the text to Clipboard
	sleep 400
	
	shortDay = %A_DD%
	
	longDay = %A_DDDD%			;todayName will have the literal day. i.e., Monday, Tuesday, ec
	IfEqual, longDay, Monday
		shortDay -= 03			;Because the yesterday of Monday is Friday
	Else
		shortDay -= 01
	shortDay := 0 . shortDay		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(shortDay)
	if (length <= 2 & length > 0)
		;MsgBox Less than or equal to 2
		sdlk = 1
	else
		StringTrimLeft, shortDay, shortDay, 1
						
	StartPos := RegExMatch(ClipBoard,shortDay "\s([0-9]{1,})\s([0-9]{1,})", misst) ;finds values of table for yesterday
	;MsgBox first: %miss1% second: %miss2%
	
	global missticket1 := misst1
	global missticket2 := misst2
}

openMisroutedService()
{
	SetTitleMatchMode, 2
	WinActivate, Google Chrome
	sleep 600
	MouseMove, 1442, 132 			;Hover mouse on Reports tab
	sleep 600
	MouseMove, 1428, 161 			;Select option Operational Reports
	sleep 600
	MouseMove, 1626, 163 			;Move mouse to the first option of the Operational Reports menu (change)
	sleep 600
	MouseMove, 1612, 359			;Hover mouse on Service Requests
	sleep 600
	MouseMove, 1787, 346			;Hover mouse on first option of Service Request (Mean time to...)
	sleep 600
	MouseClick, left, 1782, 381		;Click in Misrouted Service Requests option.
	sleep 600
	MouseMove, 1616, 779			;Hover mouse on loading image
	waitColorChange()
	waitColorChange()
	MouseClick, left, 1208, 170 		;It clicks on the settings wheel
	sleep 600
	MouseClick, left, 1685, 822 		;It clicks on the dropdown menu
	sleep 600
	MouseClick, left, 1624, 866 		;HERE I AM SELECTING Misrouted report
	sleep 600
	MouseClick, left, 1578, 890 		;I clicking in Okay option.
	sleep 100
	waitColorChange()
	MouseClick, left, 2070, 253		;Click in filter option
	sleep 500
	MouseClick, left, 1613, 676		;Click on Date field
	sleep 500
	MouseClick, left, 1622, 660		;Go to the next month
	sleep 500
						;Following code clicks on first day of the month
	
	PixelSearch, Xpos, Ypos, 413, 706, 612, 738, 0xA7CCE9, 3, Fast RGB		;Coordinates ARE relative in SPY
	sleep 300
	CoordMode, Mouse, Relative		;Changing Relative to click on same coordinates received.
	MouseClick, left, Xpos, Ypos
	CoordMode, Mouse, Screen		;Return coordinate to both desktops, not only active window
	MouseClick, left, 1585, 1119		;Click in Done Button
	sleep 1000
	waitColorChange()
	MouseClick, left, 1118, 170		;Click in the Table icon
	sleep 10000 				;Sleep 10 seconds since waitColorChange is very volatile
	
	MouseClick, left, 1534, 1358		;Click anywhere on the screen (NOT ON TABLE)
	sleep 300
	send ^a					;Select all the text found
	sleep 700
	send ^c					;Copy all the text to Clipboard
	sleep 400
	
	shortDay = %A_DD%
	
	longDay = %A_DDDD%			;todayName will have the literal day. i.e., Monday, Tuesday, ec
	IfEqual, longDay, Monday
		shortDay -= 03			;Because the yesterday of Monday is Friday
	Else
		shortDay -= 01
	shortDay := 0 . shortDay		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(shortDay)
	if (length <= 2 & length > 0)
		;MsgBox Less than or equal to 2
		proba = 1
	else
		StringTrimLeft, shortDay, shortDay, 1
						
	StartPos := RegExMatch(ClipBoard,shortDay "\s([0-9]{1,})\s([0-9]{1,})", misss) ;finds values of table for yesterday
	;MsgBox first: %miss1% second: %miss2%
	global missdistribution1 := misss1
	global missdistribution2 := misss2
}

openEUCIncident()
{
	SetTitleMatchMode, 2
	WinActivate, Google Chrome
	sleep 600
	MouseMove, 1442, 132				;Hovers over Reports
	sleep 600
	MouseMove, 1467, 193				;Hovers over Volumetric Reports
	sleep 600
	MouseMove, 1631, 242				;Hovers over Incident
	sleep 600
	MouseClick, left, 1786, 245			;Clicks in Incident Distribution
	sleep 300					;Waits 300 miliseconds to move the mouse to wait for color change
	MouseMove, 1616, 779				;Hovers over loading image
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	MouseClick, left, 1208, 170 			;It clicks on the settings wheel
	sleep 600
	MouseClick, left, 1685, 822 			;It clicks on the dropdown menu
	sleep 600
	MouseClick, left, 1597, 887 			;Select SD OPEN REPORT
	sleep 600
	MouseClick, left, 1578, 890 			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	waitColorChange()				;Wait for color change
	MouseClick, left, 2070, 253			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 1613, 676			;Clicks on the date field
	sleep 500					;Waits for popup to appear
	MouseClick, left, 1622, 660			;Clicks on next month
	sleep 500					;waits for month to change on screen

	PixelSearch, Xpos, Ypos, 413, 706, 610, 879, 0xF40F0F, 3, Fast RGB
	sleep 300					;Look for the first red color (Day of today)
	CoordMode, Mouse, Relative			;Changes focus to active window (To use the prev coor to select today)
	MouseClick, left, Xpos, Ypos			;Clicks on day of today
	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 1585, 1119			;Click in Done
	sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 2000, 318			;Clicks on Total count number (Skip table)
	MouseMove, 1583, 938				;Moving mouse to activate color change
	waitColorChange()
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 1000
	MouseClick, left, 1131, 220			;Click on the Eye (add/remove columns)
	sleep 3000 					;Maybe delete this
	send {TAB 3}
	sleep 1000
	send category					;Types CATEGORY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send type					;Types TYPE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send item					;Types ITEM
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Created By					;Types CREATED BY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Solved Date				;Types SOLVED DATE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Closed Date				;Types CLOSED DATE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Source					;Types SOURCE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300
	Send OWNER GROUP NAME				;Types OWNER GROUP NAME
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300
	Send Business Unit				;Types BUSINESS UNIT
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Group Change Count				;Types GROUP CHANGE COUNT
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300						
	Send Resolution Target Date			;Types RESOLUTION TARGET DATE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300						
	Send Country					;Types COUNTRY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Location					;Types LOCATION
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300	
	MouseClick, left, 1519, 1058
	MouseClick, left, 1090, 219			;It clicks on download file
}

openEUCServiceRequest()
{
	SetTitleMatchMode, 2
	WinActivate, Google Chrome
	sleep 600
	MouseMove, 1442, 140				;Goes to Service Request distribution option
	sleep 600
	MouseMove, 1467, 193
	sleep 600
	MouseMove, 1609, 190
	sleep 600
        MouseMove, 1645, 345	
	sleep 600
	MouseClick, left, 1808, 351
	sleep 300
	MouseMove, 1616, 779
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	MouseClick, left, 1208, 170 			;It clicks on the settings wheel
	sleep 600
	MouseClick, left, 1685, 822 			;It clicks on the dropdown menu
	sleep 600
	MouseClick, left, 1616, 888 			;Select EUC OPEN REPORT
	sleep 600
	MouseClick, left, 1578, 890 			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	waitColorChange()				;Wait for color change
	MouseClick, left, 2070, 253			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 1613, 676			;Clicks on the date field
	sleep 500					;Waits for popup to appear
	MouseClick, left, 1622, 660			;Clicks on next month
	sleep 500					;waits for month to change on screen

	PixelSearch, Xpos, Ypos, 413, 706, 610, 879, 0xF40F0F, 3, Fast RGB
	sleep 300					;Look for the first red color (Day of today)
	CoordMode, Mouse, Relative			;Changes focus to active window (To use the prev coor to select today)
	MouseClick, left, Xpos, Ypos			;Clicks on day of today
	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 1585, 1119			;Click in Done
	sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 2000, 318			;Clicks on Total count number (Skip table)
	MouseMove, 1583, 938				;Moving mouse to activate color change
	waitColorChange()
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 1000
	MouseClick, left, 1131, 220			;Click on the Eye (add/remove columns)
	sleep 3000 					;Maybe delete this
	send {TAB 3}
	sleep 1000
	send category					;Types CATEGORY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send type					;Types TYPE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send item					;Types ITEM
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Created By					;Types CREATED BY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Solved Date				;Types SOLVED DATE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Closed Date				;Types CLOSED DATE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Source					;Types SOURCE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300
	Send OWNER GROUP NAME				;Types OWNER GROUP NAME
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300
	Send Business Unit				;Types BUSINESS UNIT
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Group Change Count				;Types GROUP CHANGE COUNT
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300						
	Send Resolution Target Date			;Types RESOLUTION TARGET DATE
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300						
	Send Country					;Types COUNTRY
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300			
	Send Location					;Types LOCATION
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300	
	MouseClick, left, 1519, 1058
	MouseClick, left, 1090, 219			;It clicks on download file
}


wasFileDownloaded()
{
	MouseMove 1171, 1654
	waitColorChange()
	sleep 2000
	MouseClick left, 2081, 1657			;Click downloads notification in Chrome
	sleep 2000
}


extractMailCount()
{
	SetTitleMatchMode, 2
	IfWinExist, Outlook
	{
		WinActivate
		sleep 500
	}
	else
	{
		Run, outlook.exe
		sleep 2000
		MouseMove 646, 891
		waitColorChange()
		sleep 500	

	}
	send {LWin down}{up down}{up up}{LWin up}
	sleep 500

	getRussianMail()
	getPolishMail()
	getFrenchMail()
	getRomanianMail()
	getTurkishMail()
	getHungarianMail()
	sleep 1000
}


getValueFromWindow()		;This method reads data on the window containing the total number of highligted tickets
{
	sleep 100
	valor := -1
	sleep 100
	WinGetText, valor, Numero de Correos
	sleep 100
	temp := RegExMatch(valor, "([0-9]{1,})", temp)
	sleep 100
	send, {enter}
	sleep 100
	return temp1	
}

		
getRussianMail()
{
	
	MouseClick left, 118, 185
	sleep 3000
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 118, 185
	Sleep 1000	
	send ^{a}
	sleep 200
	send !{F11}
	send {LCtrl down}{HOME down}{LCtrl up}{HOME up}
	send {F5}
	global russianMail = getValueFromWindow()

	
}


getPolishMail()
{
	sleep 1000
	MouseClick left, 112, 212
	sleep 3000
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 112, 212
	Sleep 1000	
	send ^{a}
	sleep 200
	send !{F11}
	send {LCtrl down}{HOME down}{LCtrl up}{HOME up}
	send {F5}
	global polishMail = getValueFromWindow()
}


getFrenchMail()
{
	MouseClick left, 106, 234
	sleep 3000
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 106, 234
	Sleep 1000	
	send ^{a}
	sleep 200
	send !{F11}
	send {LCtrl down}{HOME down}{LCtrl up}{HOME up}
	send {F5}
	global frenchMail = getValueFromWindow()
}


getRomanianMail()
{
	MouseClick left, 108, 257
	sleep 3000
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 108, 257
	Sleep 1000	
	send ^{a}
	sleep 200
	send !{F11}
	send {LCtrl down}{HOME down}{LCtrl up}{HOME up}
	send {F5}
	global romanianMail = getValueFromWindow()
}


getTurkishMail()
{
	MouseClick left, 118, 277
	sleep 3000
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 118, 277	
	Sleep 1000
	send ^{a}
	sleep 200
	send !{F11}
	send {LCtrl down}{HOME down}{LCtrl up}{HOME up}
	send {F5}
	global turkishMail = getValueFromWindow()
}


getHungarianMail()
{
	MouseClick left, 107, 302
	sleep 3000
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 107, 302
	Sleep 1000	
	send ^{a}
	sleep 200
	send !{F11}
	send {LCtrl down}{HOME down}{LCtrl up}{HOME up}
	send {F5}
	global hungarianMail = getValueFromWindow()
}


setFilterProperties()
{
	;Enter to Advanced Settings Menu on Outlook 2013 to add the filter to the context menu
	send !{3}
	send {f}
	send {ctrl down}{tab}{tab}{ctrl up}
	send !{i}
	send {up}
	send {enter}
	send !{a}
	send {enter}
	send {tab 2}
	send {Enter}
	send {tab 8}
	send {enter}
	
	;Enter to Advanced Settings Menu to select the previously created filter LastVerb.cfg
	send !{3}
	send {f}
	send {ctrl down}{tab}{tab}{ctrl up}
	send !{i}
	send {up 2}
	send {Right}
	send {down}
	send {enter}

	send {tab 2}
	send {End}
	send {tab 2}
	send {Enter}
	send {y}
	send {tab 8}
	send {enter}
}

fillKPI()
{
	global
	SetTitleMatchMode, 2
	WinActivate, kpi
	MouseClick, left, 789, 1174
	send {DELETE}
	sleep 2000
	send %incident1%
	sleep 2000
	send {Down}
	send %incident2%
	sleep 2000
	send {Down}
	send %incident3%	
	sleep 2000
	send {Down}
	send %incident4%
	sleep 2000
	send {Down}
	send %distribution1%
	sleep 2000
	send {Down}
	send %distribution2%
	sleep 2000
	send {Down}
	send %distribution3%	
	sleep 2000
	send {Down}
	send %distribution4%
	sleep 2000
	send {Up 28}
	send %missticket1%
	sleep 2000
	send {Down}
	send %missticket2%
	sleep 2000
	send {Down 2}
	send %missdistribution1%
	sleep 2000
	send {Down}
	send %missdistribution2%
	sleep 2000
	send {down 26}
	send %russianMail%
	sleep 2500
	send {down}
	send %polishMail%
	sleep 2500
	send {down}
	send %frenchMail%
	sleep 2500
	send {down}
	send %romanianMail%
	sleep 2500
	send {down}
	send %turkishMail%	
	sleep 2500
	send {down}
	send %hungarianMail%
	sleep 1000


	
	;MsgBox Incident1: %incident1%`n Incident2: %incident2%`n Incident3: %incident3%`n Incident4: %incident4%`n Incidentdistribution1: %distribution1%`n Incidentdistribution2: %distribution2%`n Incidentdistribution3: %distribution3%`n Incidentdistribution4: %distribution4%`n Missrouted Ticket1: %missticket1%`n Missrouted Ticket2: %missticket2%`n Missrouted Distribution1: %missdistribution1%`n Missrouted Distribution2: %missdistribution2%`n Russian Mail: %russianMail%`n Polish Mail: %PolishMail%`n French Mail: %frenchMail%`n Romanian Mail: %romanianMail%`n Turkish Mail: %turkishMail%`n Hungarian Mail: %hungarianMail%
}