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
	openIncidentDistribution()
	;openServiceDistribution()
	;openMisroutedTickets()
	;openEUCIncident()
	;openEUCServiceRequest()
	;openMisroutedService()
	;queueStatistic()
return


queueStatistic()
{
	/*
	MouseMove 1504, 282
	send ^l
	send https://drwr00.ngcc.bt.com/Reports/Pages/Folder.aspx?ItemPath=/CARGILL
	send {ENTER}
	waitColorChange()
	send BUD003@cargill.ngcc.bt.com
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
	
	MouseMove, 1442, 132				;Hovers over Reports
	MouseMove, 1467, 193				;Hovers over Volumetric Reports
	MouseMove, 1631, 242				;Hovers over Incident
	MouseClick, left, 1786, 245			;Clicks in Incident Distribution
	sleep 300					;Waits 300 miliseconds to move the mouse to wait for color change
	MouseMove, 1616, 779				;Hovers over loading image
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	MouseClick, left, 1208, 170 			;It clicks on the settings wheel
	MouseClick, left, 1685, 822 			;It clicks on the dropdown menu
	MouseClick, left, 1648, 867 			;Select SD OPEN REPORT
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
	sleep 10000 Maybe remove this			;Wait 10 seconds for table to appear
	MouseClick, left, 1534, 1358			;Clicks on position that is empty to be able to use next shortcuts
	sleep 300
	send ^a						;Select All text on webpage
	sleep 700
	send ^c						;Copy All text on webpage
	sleep 400
	StartPos := RegExMatch(ClipBoard, "Total:\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})", Inc)
							;Find the values after "Total: " and stores them in Inc array
	
	MouseClick, left, 1969, 1380			;Scroll horizontally to the right
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

	MsgBox Inc: %Inc1% Inc: %Inc2% Inc: %Inc3% Inc: %Inc4%

	/*
	;Activating Excel to copy all of the values into it.
	SetTitleMatchMode, 2
	WinActivate, Cargill_kpi ;Will match any window name containing 'Notepad'
	send {ctrl down}{PgDn down}{ctrl up}{PgDn up}
	*/
	
}

openServiceDistribution()
{
	
	MouseMove, 1442, 140				;Goes to Service Request distribution option
	MouseMove, 1467, 193
	MouseMove, 1609, 190
        MouseMove, 1645, 345	
	MouseClick, left, 1808, 351
	sleep 300
	MouseMove, 1616, 779
	waitColorChange()
	waitColorChange()
	MouseClick, left, 1208, 170 			;It clicks on the settings wheel
	MouseClick, left, 1685, 822 			;It clicks on the dropdown menu
	MouseClick, left, 1648, 867 			;Select SD OPEN REPORT
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
	sleep 10000 Maybe remove this
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
	MsgBox Dis: %Dis1% Dis: %Dis2% Dis: %Dis3% Dis: %Dis4%
	
	/*
	;Activating Excel to copy all of the values into it.
	SetTitleMatchMode, 2
	WinActivate, Cargill_kpi ;Will match any window name containing 'Notepad'
	send {ctrl down}{PgDn down}{ctrl up}{PgDn up}
	*/
}




openMisroutedTickets()
{
		
	MouseMove, 1442, 132 			;Hover mouse on Reports tab
	MouseMove, 1428, 161 			;Select option Operational Reports
	MouseMove, 1626, 161 			;Move mouse to the first option of the Operational Reports menu (change)
	MouseMove, 1614, 215			;Hover mouse on Incident
	MouseMove, 1614, 215			;Hover mouse on first option of Incident (Incident Aging)
	MouseMove, 1825, 221
	MouseClick, left, 1791, 448		;Click in Misrouted Incidents option.
	MouseMove, 1616, 779			;Hover mouse on loading image
	waitColorChange()
	waitColorChange()
	MouseClick, left, 1208, 170 		;It clicks on the settings wheel
	MouseClick, left, 1685, 822 		;It clicks on the dropdown menu
	MouseClick, left, 1624, 866 		;HERE I AM SELECTING Misrouted report
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
						
	StartPos := RegExMatch(ClipBoard,shortDay "\s([0-9]{1,})\s([0-9]{1,})", misst) ;finds values of table for yesterday
	MsgBox first: %misst1% second: %misst2%
	
						;Activating Excel to copy all of the values into it.
	SetTitleMatchMode, 2
	WinActivate, Cargill_kpi 		;Will match any window name containing 'Notepad'
	send {ctrl down}{PgDn down}{ctrl up}{PgDn up}
	
}

openEUCIncident()
{
	
	MouseMove, 1442, 132				;Hovers over Reports
	MouseMove, 1467, 193				;Hovers over Volumetric Reports
	MouseMove, 1631, 242				;Hovers over Incident
	MouseClick, left, 1786, 245			;Clicks in Incident Distribution
	sleep 300					;Waits 300 miliseconds to move the mouse to wait for color change
	MouseMove, 1616, 779				;Hovers over loading image
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	MouseClick, left, 1208, 170 			;It clicks on the settings wheel
	MouseClick, left, 1685, 822 			;It clicks on the dropdown menu
	MouseClick, left, 1597, 887 			;Select SD OPEN REPORT
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
	
	MouseMove, 1442, 140				;Goes to Service Request distribution option
	MouseMove, 1467, 193
	MouseMove, 1609, 190
        MouseMove, 1645, 345	
	MouseClick, left, 1808, 351
	sleep 300
	MouseMove, 1616, 779
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	MouseClick, left, 1208, 170 			;It clicks on the settings wheel
	MouseClick, left, 1685, 822 			;It clicks on the dropdown menu
	MouseClick, left, 1616, 888 			;Select EUC OPEN REPORT
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

openMisroutedService()
{
		
	MouseMove, 1442, 132 			;Hover mouse on Reports tab
	MouseMove, 1428, 161 			;Select option Operational Reports
	MouseMove, 1626, 163 			;Move mouse to the first option of the Operational Reports menu (change)
	MouseMove, 1612, 359			;Hover mouse on Service Requests
	MouseMove, 1787, 346			;Hover mouse on first option of Service Request (Mean time to...)
	MouseClick, left, 1782, 381		;Click in Misrouted Service Requests option.
	
	MouseMove, 1616, 779			;Hover mouse on loading image
	waitColorChange()
	waitColorChange()
	MouseClick, left, 1208, 170 		;It clicks on the settings wheel
	MouseClick, left, 1685, 822 		;It clicks on the dropdown menu
	MouseClick, left, 1624, 866 		;HERE I AM SELECTING Misrouted report
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
						
	StartPos := RegExMatch(ClipBoard,shortDay "\s([0-9]{1,})\s([0-9]{1,})", misss) ;finds values of table for yesterday
	MsgBox first: %misss1% second: %misss2%
	
						;Activating Excel to copy all of the values into it.
	SetTitleMatchMode, 2
	WinActivate, Cargill_kpi 		;Will match any window name containing 'Notepad'
	send {ctrl down}{PgDn down}{ctrl up}{PgDn up}
	
}