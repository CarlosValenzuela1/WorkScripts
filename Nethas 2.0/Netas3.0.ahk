








































#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force
#Persistent
#MaxThreadsPerHotkey 2
Scrolllock::Pause



^F1::
	CoordMode, Mouse, Screen
	mainMethod()
return





extractMailCount()
{
	SetTitleMatchMode, 2
	IfWinExist, Outlook
	{
		WinActivate
		sleep 100
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
	sleep 1000
	getPolishMail()
	sleep 1000
	getFrenchMail()
	sleep 1000
	getRomanianMail()
	sleep 1000
	getTurkishMail()
	sleep 1000
	getHungarianMail()
	sleep 1000
}



getValueFromWindow()
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
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 118, 185
	Sleep 500	
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
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 112, 212
	Sleep 500	
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
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 106, 234
	Sleep 500	
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
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 108, 257
	Sleep 500	
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
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 118, 277	
	Sleep 500
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
	send {ctrl down}e{ctrl up}
	send {tab 2}
	send {SPACE}
	send {enter}
	setFilterProperties()
	MouseClick left, 107, 302
	Sleep 500	
	send ^{a}
	sleep 200
	send !{F11}
	send {LCtrl down}{HOME down}{LCtrl up}{HOME up}
	send {F5}
	global hungarianMail = getValueFromWindow()
}







































setFilterProperties()
{
	;Enter to Advanced Settings Menu to add the filter to the context menu
	send !{3}
	send {f}
	sleep 200
	send {ctrl down}{tab}{tab}{ctrl up}
	sleep 200
	send !{i}
	sleep 200
	send {up}
	sleep 200
	send {enter}
	sleep 200
	send !{a}
	sleep 200
	send {enter}
	sleep 200
	send {tab 2}
	sleep 200
	send {Enter}
	sleep 200
	send {tab 8}
	sleep 200
	send {enter}
	
	;Enter to Advanced Settings Menu to add the filter to the context menu
	send !{3}
	send {f}
	sleep 200
	send {ctrl down}{tab}{tab}{ctrl up}
	sleep 200
	send !{i}
	sleep 200
	send {up 2}
	sleep 200
	send {Right}
	sleep 200
	send {down}
	sleep 200
	send {enter}
	sleep 200

	send {tab 2}
	sleep 200
	send {End}
	sleep 200
	send {tab 2}
	sleep 200
	send {Enter}
	sleep 200
	send {y}
	sleep 200
	send {tab 8}
	sleep 200
	send {enter}
}



mainMethod()
{	
	global

	initializeVariables()
	
	

	createDirectory()			;Already tested and working

	;getCSVfile()

	;getIncDisfile()
	
	;getSerDisfile()

	

	;url = http://itreporting.cargill.com/Netasthra/loginform
	;run % "chrome.exe" ( winExist("ahk_class Chrome_WidgetWin_1") ? " --new-window " : " " ) url
	run, chrome.exe http://itreporting.cargill.com/Netasthra/loginform
	send {LWin down}{UP down}{UP up}{LWin up}		;Maximizing window
	MouseMove 1553, 413
	waitColorChange()
	sleep 3000
	send {TAB}
	
	send, c760912{tab}Pear1234{enter}
	waitColorChange()
	sleep 2500
	CoordMode, Mouse, Screen			;It only works on current window
	openIncidentDistribution()
	wasFileDownloaded()

	getIncDisfile()	

	CoordMode, Mouse, Screen			;It only works on current window
	openServiceDistribution()
	wasFileDownloaded()

	getSerDisfile()	

	openMisroutedTickets()
	sleep 5000
	openMisroutedService()
	sleep 5000
	


	openEUCIncident()
	wasFileDownloaded()
	sleep 2000
	getIncDisfileEUC()	
	sleep 5000

	openEUCServiceRequest()
	wasFileDownloaded()
	sleep 2000
	getSerDisfileEUC()	
	sleep 5000
	
	

	extractMailCount()



	SetTitleMatchMode, 2
	WinActivate, kpi
	

	MouseClick, left, 759, 1168
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


	
	MsgBox Incident1: %incident1%`n Incident2: %incident2%`n Incident3: %incident3%`n Incident4: %incident4%`n Incidentdistribution1: %distribution1%`n Incidentdistribution2: %distribution2%`n Incidentdistribution3: %distribution3%`n Incidentdistribution4: %distribution4%`n Missrouted Ticket1: %missticket1%`n Missrouted Ticket2: %missticket2%`n Missrouted Distribution1: %missdistribution1%`n Missrouted Distribution2: %missdistribution2%`n Russian Mail: %russianMail%`n Polish Mail: %PolishMail%`n French Mail: %frenchMail%`n Romanian Mail: %romanianMail%`n Turkish Mail: %turkishMail%`n Hungarian Mail: %hungarianMail%

		
}

initializeVariables()
{
	;==================;
	; Global Variables ;
	;==================;
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
	;global russianMail = -1
	global polishMail := -1
	global frenchMail := -1
	global romanianMail := -1
	global turkishMail := -1
	global hungarianMail := -1
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

	length := StrLen(csvDay)
	if (length <= 2 & length > 0)
		MsgBox Less than or equal to
	else
		StringTrimLeft, csvDay, csvDay, 1

	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\
	send {ENTER}
	sleep 600
	send {Ctrl down}{Shift down}n{Ctrl up}{Shift up}
	sleep 500
	;MsgBox %csvDay% %csvMonth%
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
	sleep 2000
	send {LWin down}{UP down}{UP up}{LWin up}		;Maximizing window
	sleep 2000
	send {LAlt down}d{LAlt up}
	sleep 2000
	send C:\Users\c760912\Downloads
	sleep 2000
	send {ENTER}
	sleep 2000
	send {TAB}
	sleep 2000
	sendRaw Incident_Distribution
	sleep 2000
	send {Enter}
	sleep 2000
	send {TAB 3}
	sleep 2000
	send ^a
	sleep 2000
	send ^x
	sleep 2000
	send {LAlt down}d{LAlt up}
	sleep 2000
	csvMonth = %A_MMMM%
	csvDay = %A_DD%
	shortDay := 0 . csvDay		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(shortDay)
	if (length <= 2 & length > 0)
		MsgBox Less than or equal to 2
	else
		StringTrimLeft, shortDay, shortDay, 1
	sleep 2000
	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%
	sleep 2000
	send {Enter}
	sleep 2000
	send {Tab 5}
	sleep 2000
	send ^v
	sleep 2000
	send {Alt down}{F4 down}{Alt up}{F4 up}
	sleep 2000
}

getSerDisfile()
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
	sleep 300
	sendRaw Service_Request
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
		MsgBox Less than or equal to 2
	else
		StringTrimLeft, shortDay, shortDay, 1

	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%
	send {Enter}
	sleep 100
	send {Tab 5}
	sleep 400
	send ^v
	sleep 400
	send {Alt down}{F4 down}{Alt up}{F4 up}
	sleep 900
}


getIncDisfileEUC()
{
	send {LWin down}e{LWin up} 
	sleep 2000
	send {LWin down}{UP down}{UP up}{LWin up}		;Maximizing window
	sleep 2000
	send {LAlt down}d{LAlt up}
	sleep 2000
	send C:\Users\c760912\Downloads
	sleep 2000
	send {ENTER}
	sleep 2000
	send {TAB}
	sleep 2000
	sendRaw Incident_Distribution
	sleep 2000
	send {Enter}
	sleep 1000
	send {TAB 3}
	sleep 2000
	send ^a
	sleep 2000
	send ^x
	sleep 2000
	send {LAlt down}d{LAlt up}
	sleep 2000
	csvMonth = %A_MMMM%
	csvDay = %A_DD%
	shortDay := 0 . csvDay		;Puting a zero in front of the day of Yesterday to match table

	length := StrLen(shortDay)
	if (length <= 2 & length > 0)
		MsgBox Less than or equal to 2
	else
		StringTrimLeft, shortDay, shortDay, 1
	sleep 2000
	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%\EUC Report - %csvDay%_%csvMonth% 
	sleep 2000
	send {Enter}
	sleep 2000
	send {Tab 5}
	sleep 2000
	send ^v
	sleep 2000
	send {Alt down}{F4 down}{Alt up}{F4 up}
	sleep 2000
}

getSerDisfileEUC()
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
	sleep 300
	sendRaw Service_Request
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
		MsgBox Less than or equal to 2
	else
		StringTrimLeft, shortDay, shortDay, 1

	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%\EUC Report - %csvDay%_%csvMonth% 
	send {Enter}
	sleep 100
	send {Tab 5}
	sleep 400
	send ^v
	sleep 400
	send {Alt down}{F4 down}{Alt up}{F4 up}
	sleep 900
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
		MsgBox Less than or equal to 2
	else
		StringTrimLeft, shortDay, shortDay, 1

	send C:\Users\c760912\Documents\Daily Report\%csvMonth%\%csvDay%_%csvMonth%
	send {Enter}
	sleep 100
	send {Tab 5}
	sleep 200
	send ^v
	send {Alt down}{F4 down}{Alt up}{F4 up}
	sleep 900
}





















































wasFileDownloaded()
{
	MouseMove 1171, 1654
	waitColorChange()
	sleep 2000
	MouseClick left, 2081, 1657			;Click downloads notification in Chrome
	sleep 2000
}






queueStatistic()
{
	/*
	run, chrome.exe https://drwr00.ngcc.bt.com/Reports/Pages/Folder.aspx?ItemPath=/CARGILL
	send {ENTER}
	waitColorChange()
	send BUD003@cargill.ngcc.bt.comw
	send {TAB up}{Tab down}
	send tcs{#}1234
	send {ENTER}
	
 	send {TAB 12}	            				;Move to Queue Reports					
        sleep 1000
	send {ENTER}						
    	sleep 1000 
	send {TAB 12} 						;Move to Queue Statistics
	sleep 1000
	send {ENTER}
	MouseMove 1853, 212     				;Positioning mouse for color change
	waitColorChange()
	send {TAB 8}    					;Move to "From" field
		
 	*/
	
	yesterday = %A_DD%
	yesterday =  
	MsgBox Current value of the date: %yesterday%
	
}

















































































waitColorChange()
{
CoordMode, Mouse, Relative					;It only works on current window
found := 0							;Initializes the variable found
loopTimes := 0							;Initializes the variable loopTimes
MouseGetPos Xpos, Ypos						;Obtains mouse position at that time
PixelGetColor firstColor, %Xpos%, %Ypos%, RGB			;Gets the color of the pixel where mouse is at
secondColor := 1						;Initializes the variable secondColor
while(found == 0)						;Check if the color has been found, if not then it continues
{
	sleep 400						;Waits 400 miliseconds for animations that haven't finished
	if(secondColor != firstColor && loopTimes != 0)
		found := 1
	else
	{
		MouseGetPos Xpos, Ypos				;Gets position of mouse
		PixelGetColor secondColor, %Xpos%, %Ypos%, RGB	;gets the color of the pixel where mouse is at
		loopTimes := (loopTimes + 1)			;Increases the times it has checked for a color change
	}
	;MsgBox ending of while found: %found% loopTimes: %loopTimes% firstColor: %firstColor% secondColor: %secondColor%
cliboard := 							;NOT SURE IF ITS NEEDED
}
CoordMode, Mouse, Screen					;Returns coordinates to both monitors
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
	sleep 4000 Maybe remove this			;Wait 10 seconds for table to appear
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
	MouseClick, left, 1918, 417 			;Clicking on the Total cell of the Incident Distribution table.
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
	;MsgBox Inc: %Inc1% Inc: %Inc2% Inc: %Inc3% Inc: %Inc4%

	global incident1 = Inc1
	global incident2 := Inc2
	global incident3 := Inc3
	global incident4 := Inc4

	/*
	;Activating Excel to copy all of the values into it.
	SetTitleMatchMode, 2
	WinActivate, Cargill_kpi ;Will match any window name containing 'Notepad'
	send {ctrl down}{PgDn down}{ctrl up}{PgDn up}
	*/
	
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

	/*
	;Activating Excel to copy all of the values into it.
	SetTitleMatchMode, 2
	WinActivate, Cargill_kpi ;Will match any window name containing 'Notepad'
	send {ctrl down}{PgDn down}{ctrl up}{PgDn up}
	*/
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
		MsgBox Less than or equal to 2
	else
		StringTrimLeft, shortDay, shortDay, 1
						
	StartPos := RegExMatch(ClipBoard,shortDay "\s([0-9]{1,})\s([0-9]{1,})", misst) ;finds values of table for yesterday
	;MsgBox first: %miss1% second: %miss2%
	
	global missticket1 := misst1
	global missticket2 := misst2
						;Activating Excel to copy all of the values into it.
	/*
	SetTitleMatchMode, 2
	WinActivate, Cargill_kpi 		;Will match any window name containing 'Notepad'
	send {ctrl down}{PgDn down}{ctrl up}{PgDn up}
	*/
	
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
		MsgBox Less than or equal to 2
	else
		StringTrimLeft, shortDay, shortDay, 1
						
	StartPos := RegExMatch(ClipBoard,shortDay "\s([0-9]{1,})\s([0-9]{1,})", misss) ;finds values of table for yesterday
	;MsgBox first: %miss1% second: %miss2%
	global missdistribution1 := misss1
	global missdistribution2 := misss2
	/*					;Activating Excel to copy all of the values into it.
	SetTitleMatchMode, 2
	WinActivate, Cargill_kpi 		;Will match any window name containing 'Notepad'
	send {ctrl down}{PgDn down}{ctrl up}{PgDn up}
	*/
	
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

	/*
	;Activating Excel to copy all of the values into it.
	SetTitleMatchMode, 2
	WinActivate, Cargill_kpi ;Will match any window name containing 'Notepad'
	send {ctrl down}{PgDn down}{ctrl up}{PgDn up}
	*/
	
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

	/*
	;Activating Excel to copy all of the values into it.
	SetTitleMatchMode, 2
	WinActivate, Cargill_kpi ;Will match any window name containing 'Notepad'
	send {ctrl down}{PgDn down}{ctrl up}{PgDn up}
	*/
	
}

