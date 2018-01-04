;=========================================================================================================;
; Netas Script: This script gathers specific data from Netasthra:   					  ; 
; • Incident Distribution										  ;
; • Service Request Distribution 									  ;
; • Misrouted Tickets 									  		  ;
; • Misrouted Service Request	 									  ;
;=========================================================================================================;
; Description of each method will be find below.							  ;
; Version: 1.0												  ;
; Author: Carlos Valenzuela										  ;
; Company: TATA Consultancy Services									  ;
;=========================================================================================================;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force
#Persistent
#MaxThreadsPerHotkey 2


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
return 


Scrolllock::Pause
PrintScreen::Suspend
Insert::Reload

Pause::
InputBox, Var, Reports Script, • SD Incident Distribution "inc" `n• SD Service Request "sr" `n• Misrouted Incidents "miss1" `n• Missrouted Service Request "miss2" `n• EUC Incident Distrubition "eucinc"`n• EUC Service Request "eucsr"`n• Agent State "state"`n• Queue Statistics "statistics"  `n• Queues by Agent "qagent"  `n• Chat Reports "chat"  `n• ScoreCard "reported"  `n• Test "test" , , 250, 350

if (Var = "inc")
{
	InputBox, fecha, Date, Please type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	CoordMode, Mouse, Screen
	mainMethodIncidentSD()
} 
else if (Var = "sr")
{
	InputBox, fecha, Date, Please type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	CoordMode, Mouse, Screen
	mainMethodServiceSD()
}
else if (Var = "miss1")
{
	InputBox, fecha, Date, Please type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	CoordMode, Mouse, Screen
	mainMethodMiss1()
}
else if (Var = "miss2")
{
	 InputBox, fecha, Date, Please type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	CoordMode, Mouse, Screen
	mainMethodMiss2()
}
else if (Var = "eucinc")
{
	InputBox, fecha, Date, Please type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	CoordMode, Mouse, Screen
	mainMethodIncidentEUC()
}
else if (Var = "eucsr")
{
	InputBox, fecha, Date, Please type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	CoordMode, Mouse, Screen
	mainMethodServiceEUC()
}
else if (Var = "state")
{
	InputBox, fecha, Date, START type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	InputBox, ayer, Date, END type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	CoordMode, Mouse, Screen
	mainMethodState()
}
else if (Var = "statistics")
{
	InputBox, fecha, Date, START type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	InputBox, ayer, Date, END type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	CoordMode, Mouse, Screen
	mainMethodStatistics()
}
else if (Var = "qagent")
{
	InputBox, fecha, Date, START type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	InputBox, ayer, Date, END type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	CoordMode, Mouse, Screen
	mainMethodqagent()
}
else if (Var = "chat")
{
	InputBox, fecha, Date, START type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	mainMethodChat()
}
else if (Var = "reported")
{
	InputBox, fecha, Date, START type the month of the report you would like to get: "January" , , 450, 125
	mainMethodReported()

}
else if (Var = "test")
{
	send hello
	Input, SingleKey, L1, {right}
	send how
	Input, SingleKey, L1, {right}
	send are
	Input, SingleKey, L1, {right}
	send you?
}
else
{
	MsgBox, Incorrect categorization
}






filterDate()
{
	selectMonth()
	countWhites()
	selectDay()
}

selectMonth()
{
	global fecha
	loop = 0
	while(loop=0)
{
	CoordMode, Mouse, Screen
	clipboard = 
	monthValue = 0
	datestring = %fecha%
	StringSplit, d, datestring, /
	date := d1
	FormatTime, typedMonth, %date%, MM
	typedMonth = %date%
	MouseClick, left, 2844, 405
	MouseClick, left, 2844, 405
	send ^c
	ClipWait
	monthRequested = %clipboard%
	
	if (monthRequested = "January")  
		monthValue = 01
	else if (monthRequested = "February")  
		monthValue = 02
	else if (monthRequested = "March")  
		monthValue = 03
	else if (monthRequested = "April")  
		monthValue = 04
	else if (monthRequested = "May")  
		monthValue = 05
	else if (monthRequested = "June")  
		monthValue = 06
	else if (monthRequested = "July")  
		monthValue = 07
	else if (monthRequested = "August")  
		monthValue = 08
	else if (monthRequested = "September")  
		monthValue = 09
	else if (monthRequested = "October")  
		monthValue = 10
	else if (monthRequested = "November")  
		monthValue = 11
	else if (monthRequested = "December")  
		monthValue = 12
	else
		monthValue = -1

	if(monthValue < typedMonth)
	{
		;MsgBox less monthRequested: %monthValue% typedMonth: %typedMonth%
		MouseClick, left, 2927, 404
	}
	else if (monthValue > typedMonth)
	{
		;MsgBox more monthRequested: %monthValue% typedMonth: %typedMonth%
		MouseClick, left, 2792, 403
	}
	else if (monthValue = typedMonth)
	{
		;MsgBox same monthRequested: %monthValue% typedMonth: %typedMonth%
		loop = 1
	}
}	;END OF LOOP
}

selectDay()
{
	global
	CoordMode, Mouse, Relative
	datestring = %fecha%
	StringSplit, d, datestring, /
	date := d3 d1 d2
	FormatTime, day_of_Week, %date%, ddd
	dayName = %day_of_week%

	FormatTime, day_of_Week, %date%, d
	day = %day_of_week%

	week := ((whites + day)/7)-.14
	week := floor(week)

	if (dayName = "Sun")
		Xpos = 850
	else if (dayName = "Mon")
		Xpos = 880
	else if (dayName = "Tue")
		Xpos = 910
	else if (dayName = "Wed")
		Xpos = 940
	else if (dayName = "Thu")
		Xpos = 970
	else if (dayName = "Fri")
		Xpos = 1000
	else if (dayName = "Sat")
		Xpos = 1030


	if (week = 0)
		Ypos = 460
	else if (week = 1)
		Ypos = 490
	else if (week = 2)
		Ypos = 520
	else if (week = 3)
		Ypos = 550
	else if (week = 4)
		Ypos = 580

	MouseClick, left, %Xpos%, %Ypos%
	
	CoordMode, Mouse, Screen				;Returns coordinates to both monitors
}

countWhites()
{
	global
	CoordMode, Mouse, Relative				;It only works on current window
	whites := 0
	color = 0xFFFFFF

	MouseMove 850, 479
	sleep 100
	MouseGetPos Xpos, Ypos					;Obtains mouse position at that time

	PixelGetColor firstColor, %Xpos%, %Ypos%, RGB		;Gets the color of the pixel where mouse is at
	if (firstColor == color)
		whites := (whites + 1)
	
	MouseMove 880, 479
	MouseGetPos Xpos, Ypos					;Obtains mouse position at that time
	PixelGetColor firstColor, %Xpos%, %Ypos%, RGB		;Gets the color of the pixel where mouse is at
	if (firstColor = color)
		whites := (whites + 1)

	MouseMove 910, 479
	MouseGetPos Xpos, Ypos					;Obtains mouse position at that time
	PixelGetColor firstColor, %Xpos%, %Ypos%, RGB		;Gets the color of the pixel where mouse is at
	if (firstColor = color)
		whites := (whites + 1)

	MouseMove 940, 479
	MouseGetPos Xpos, Ypos					;Obtains mouse position at that time
	PixelGetColor firstColor, %Xpos%, %Ypos%, RGB		;Gets the color of the pixel where mouse is at
	if (firstColor = color)t
		whites := (whites + 1)

	MouseMove 970, 479
	MouseGetPos Xpos, Ypos					;Obtains mouse position at that time
	PixelGetColor firstColor, %Xpos%, %Ypos%, RGB		;Gets the color of the pixel where mouse is at
	if (firstColor = color)
		whites := (whites + 1)

	MouseMove 1000, 479
	MouseGetPos Xpos, Ypos					;Obtains mouse position at that time
	PixelGetColor firstColor, %Xpos%, %Ypos%, RGB		;Gets the color of the pixel where mouse is at
	if (firstColor = color)
		whites := (whites + 1)

	MouseMove 1030, 479
	MouseGetPos Xpos, Ypos					;Obtains mouse position at that time
	PixelGetColor firstColor, %Xpos%, %Ypos%, RGB		;Gets the color of the pixel where mouse is at
	if (firstColor = color)
		whites := (whites + 1)

	;MsgBox number of whites: %whites%
	;CoordMode, Mouse, Screen				;Returns coordinates to both monitors
}


initializeVariables()				;Global Variables
{
	global whites := 0
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
}

mainMethodIncidentSD()
{	
	global
	initializeVariables()
	openNetasthra()
	openIncidentDistribution()
	writeToFileInc()
}

mainMethodServiceSD()
{	
	global
	initializeVariables()
	openNetasthra()
	openServiceDistribution()
	writeToFileSer()	
}

mainMethodIncidentEUC()
{	
	global
	initializeVariables()
	openNetasthra()
	openIncidentDistributionEUC()	
}

mainMethodServiceEUC()
{	
	global
	initializeVariables()
	openNetasthra()
	openServiceDistributionEUC()	
}

mainMethodMiss1()
{	
	global
	initializeVariables()
	openNetasthra()
	openMiss1()
	writeToFileInc()
}

mainMethodMiss2()
{	
	global
	initializeVariables()
	openNetasthra()
	openMiss2()
	writeToFileSer()	
}

openNetasthra()
{
	run, chrome.exe http://itreporting.cargill.com/Netasthra/loginform
	sleep 50
	send {LWin down}{UP down}{UP up}{LWin up}						;Maximizing window
	MouseMove 2574, 457
	;waitColorChange()
	waitColorChange()
	sleep 3000
	send {TAB}
	
	send, c760912{tab}Hola12345{enter}
	waitColorChange()
	sleep 2500
	CoordMode, Mouse, Screen			;It only works on current window
}

openBT()
{
	run, chrome.exe https://drwr00.ngcc.bt.com/Reports/Pages/Folder.aspx?ItemPath=/CARGILL
	CoordMode, Mouse, Screen			;It only works on current window
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
	MouseMove, 2517, 138				;Hovers over Reports
	sleep 200
	MouseMove, 2535, 191				;Hovers over Volumetric Reports
	sleep 200
	MouseMove, 2783, 187				;Hovers over Change
	sleep 200
	MouseMove, 2791, 244				;Hovers over Incident
	sleep 200
	MouseClick, left, 2901, 249			;Clicks in Incident Distribution
	sleep 900					;Waits 900 miliseconds to move the mouse to wait for color change
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	MouseClick, left, 2081, 171 			;It clicks on the settings wheel
	sleep 300
	MouseClick, left, 2991, 557 			;It clicks on the dropdown menu
	sleep 300
	MouseClick, left, 2964, 609 			;Select SD OPEN REPORT
	sleep 300
	MouseClick, left, 2884, 640  			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange()				;Wait for color change
	MouseClick, left, 3806, 255			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 2951, 419			;Clicks on the date field
	sleep 500					;Waits for popup to appear
	
	tempMove()
	filterDate()



	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	MouseMove, 2886, 518				;Hovers over loading image
	;sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 1995, 174			;Clicks on Table icon (Upper-left side on Netasthra)
	sleep 2000
	MouseClick, left, 2495, 851			;Clicks on position that is empty to be able to use next shortcuts
	sleep 300
	send ^a						;Select All text on webpage
	sleep 300
	send ^c						;Copy All text on webpage
	sleep 400
	StartPos := RegExMatch(ClipBoard, "Total:\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})", Inc)
							;Find the values after "Total: " and stores them in Inc array
	
	MouseClick, left, 1989, 174			;Clicking on the graph icon to go back to the table
	Mouseclick, left, 3740, 272			; Click in Total Count

	MouseMove, 2886, 518				;Hovers over loading image
	sleep 400
	;waitColorChange()
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 500
	MouseClick, left, 2001, 226			;Click on the Eye (add/remove columns)
	sleep 300 					;Maybe delete this
	send {TAB 3}
	sleep 1000
	send category					;Types CATEGORY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send type					;Types TYPE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send item					;Types ITEM
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Created By					;Types CREATED BY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Solved Date				;Types SOLVED DATE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Closed Date				;Types CLOSED DATE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Source					;Types SOURCE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Country					;Types COUNTRY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Location					;Types LOCATION
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200	
	MouseClick, left, 2827, 799
	MouseClick, left, 1965, 223			;It clicks on download file

	global incident1 = Inc1
	global incident2 := Inc2
	global incident3 := Inc3
	global incident4 := Inc4


}

openServiceDistribution()
{
	MouseMove, 2517, 138				;Hovers over Reports
	sleep 200
	MouseMove, 2535, 191				;Hovers over Volumetric Reports
	sleep 200
	MouseMove, 2783, 187				;Hovers over Change
	sleep 200
	MouseMove, 2782, 354				;Hovers over Service Request
	sleep 200
	MouseClick, left, 2903, 356			;Clicks in Service Request
	sleep 900					;Waits 900 miliseconds to move the mouse to wait for color change
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	MouseClick, left, 2081, 171 			;It clicks on the settings wheel
	sleep 300
	MouseClick, left, 2991, 557 			;It clicks on the dropdown menu
	sleep 300
	MouseClick, left, 2869, 612 			;Select SD OPEN REPORT
	sleep 300
	MouseClick, left, 2884, 640  			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange()				;Wait for color change
	MouseClick, left, 3806, 255			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 2951, 419			;Clicks on the date field
	sleep 500					;Waits for popup to appear
	

	tempMove()
	filterDate()



	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	MouseMove, 2886, 518				;Hovers over loading image
	;sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 1995, 174			;Clicks on Table icon (Upper-left side on Netasthra)
	sleep 2000
	MouseClick, left, 2495, 851			;Clicks on position that is empty to be able to use next shortcuts
	sleep 300
	send ^a						;Select All text on webpage
	sleep 300
	send ^c						;Copy All text on webpage
	sleep 400
	StartPos := RegExMatch(ClipBoard, "Total:\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})", Dis)
							;Find the values after "Total: " and stores them in Inc array
	
	MouseClick, left, 1989, 174			;Clicking on the graph icon to go back to the table
	Mouseclick, left, 3740, 272			; Click in Total Count

	MouseMove, 2886, 518				;Hovers over loading image
	sleep 300
	;waitColorChange()
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 500
	MouseClick, left, 2001, 226			;Click on the Eye (add/remove columns)
	sleep 300 					;Maybe delete this
	send {TAB 3}
	sleep 1000
	send category					;Types CATEGORY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send type					;Types TYPE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send item					;Types ITEM
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Created By					;Types CREATED BY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Solved Date				;Types SOLVED DATE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Closed Date				;Types CLOSED DATE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Source					;Types SOURCE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Country					;Types COUNTRY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Location					;Types LOCATION
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200	
	MouseClick, left, 2827, 799
	MouseClick, left, 1965, 223			;It clicks on download file

	global distribution1 = Dis1
	global distribution2 := Dis2
	global distribution3 := Dis3
	global distribution4 := Dis4
}

openIncidentDistributionEUC()
{
	MouseMove, 2517, 138				;Hovers over Reports
	sleep 200
	MouseMove, 2535, 191				;Hovers over Volumetric Reports
	sleep 200
	MouseMove, 2783, 187				;Hovers over Change
	sleep 200
	MouseMove, 2791, 244				;Hovers over Incident
	sleep 200
	MouseClick, left, 2901, 249			;Clicks in Incident Distribution
	sleep 900					;Waits 900 miliseconds to move the mouse to wait for color change
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	MouseClick, left, 2081, 171 			;It clicks on the settings wheel
	sleep 300
	MouseClick, left, 2991, 557 			;It clicks on the dropdown menu
	sleep 300
	MouseClick, left, 2915, 634			;Select EUC OPEN REPORT
	sleep 300
	MouseClick, left, 2884, 640  			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange()				;Wait for color change
	MouseClick, left, 3806, 255			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 2951, 419			;Clicks on the date field
	sleep 500					;Waits for popup to appear


	tempMove()
	filterDate()



	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	;sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 3742, 271			;Clicks on Table icon (Upper-left side on Netasthra)
	sleep 200
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 500
	MouseClick, left, 2001, 226			;Click on the Eye (add/remove columns)
	sleep 300 					;Maybe delete this
	send {TAB 3}
	sleep 1000
	send category					;Types CATEGORY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send type					;Types TYPE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200	
	Send Owner Group Name				;Types Owner Group Name
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200					
	Send item					;Types ITEM
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Created By					;Types CREATED BY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Solved Date				;Types SOLVED DATE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Resolution Target Date			;Types Resolution Target Date
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Closed Date				;Types CLOSED DATE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Source					;Types SOURCE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200	
	Send Region					;Types Region
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200					
	Send Country					;Types COUNTRY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200
	Send Business Unit				;Types Business Unit
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200		
	Send Group Change Count				;Types Group Change Count
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200					
	Send Location					;Types LOCATION
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 300	
	MouseClick, left, 2827, 799
	MouseClick, left, 1965, 223			;It clicks on download file

}

openServiceDistributionEUC()
{
	MouseMove, 2517, 138				;Hovers over Reports
	sleep 200
	MouseMove, 2535, 191				;Hovers over Volumetric Reports
	sleep 200
	MouseMove, 2783, 187				;Hovers over Change
	sleep 200
	MouseMove, 2782, 354				;Hovers over Service Request
	sleep 200
	MouseClick, left, 2903, 356			;Clicks in Service Request
	sleep 900					;Waits 900 miliseconds to move the mouse to wait for color change
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	MouseClick, left, 2081, 171 			;It clicks on the settings wheel
	sleep 300
	MouseClick, left, 2991, 557 			;It clicks on the dropdown menu
	sleep 300
	MouseClick, left, 2958, 634 			;Select EUC OPEN REPORT
	sleep 300
	MouseClick, left, 2884, 640  			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange()				;Wait for color change
	MouseClick, left, 3806, 255			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 2951, 419			;Clicks on the date field
	sleep 500					;Waits for popup to appear


	tempMove()
	filterDate()


	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	;sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 3742, 271			;Clicks on Table icon (Upper-left side on Netasthra)
	sleep 200
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 500
	MouseClick, left, 2001, 226			;Click on the Eye (add/remove columns)
	sleep 300 					;Maybe delete this
	send {TAB 3}
	sleep 1000
	send category					;Types CATEGORY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send type					;Types TYPE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200	
	Send Owner Group Name				;Types Owner Group Name
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200					
	Send item					;Types ITEM
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Created By					;Types CREATED BY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Solved Date				;Types SOLVED DATE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Resolution Target Date			;Types Resolution Target Date
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Closed Date				;Types CLOSED DATE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200			
	Send Source					;Types SOURCE
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200	
	Send Region					;Types Region
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200					
	Send Country					;Types COUNTRY
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200
	Send Business Unit				;Types Business Unit
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200		
	Send Group Change Count				;Types Group Change Count
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 200					
	Send Location					;Types LOCATION
	sleep 200
	send {Tab}{Space}
	sleep 200
	send {Shift down}{Tab}{Shift up}
	sleep 300	
	MouseClick, left, 2827, 799
	MouseClick, left, 1965, 223			;It clicks on download file

}

openMiss2()
{
	MouseMove, 2517, 138				;Hovers over Reports
	sleep 200
	MouseMove, 2533, 169				;Hovers over Volumetric Reports
	sleep 200
	MouseMove, 2753, 169				;Hovers over Change
	sleep 200
	MouseMove, 2747, 346				;Hovers over Incident
	sleep 200
	MouseMove, 2916, 354
	sleep 200
	MouseClick, left, 2913, 380			;Clicks in Incident Distribution
	sleep 900					;Waits 900 miliseconds to move the mouse to wait for color change
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	
	MouseClick, left, 3806, 255			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	MouseClick, left, 2951, 419			;Clicks on the date field
	sleep 500					;Waits for popup to appear


	tempMove()
	filterDate()

	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	;sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 1995, 174			;Clicks on Table icon (Upper-left side on Netasthra)	
}

openMiss1()
{
MouseMove, 2517, 138					;Hovers over Reports
	sleep 200
	MouseMove, 2533, 169				;Hovers over Volumetric Reports
	sleep 200
	MouseMove, 2753, 169				;Hovers over Change
	sleep 200
	MouseMove, 2770, 219				;Hovers over Incident
	sleep 200
	MouseMove, 2919, 227
	sleep 200
	MouseClick, left, 2906, 455			;Clicks in Incident Distribution
	sleep 900					;Waits 900 miliseconds to move the mouse to wait for color change
	MouseMove, 2886, 518				;Hovers over loading image
	waitColorChange()				;waits for the first change color
	waitColorChange()				;waits for the secod change color
	
	MouseClick, left, 3806, 255			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	MouseClick, left, 2951, 419			;Clicks on the date field
	sleep 500					;Waits for popup to appear


	tempMove()
	filterDate()

	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	;sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 1995, 174			;Clicks on Table icon (Upper-left side on Netasthra)	
}

mainMethodState()
{
	openBT()

	navigateState()
}

navigateState()
{
	global
	sleep 2000
	send {TAB 9}{ENTER}
	sleep 2000
	send {TAB 13}{ENTER}
	sleep 3500
	send {TAB 8}
	send %fecha%{TAB}
	sleep 3500
	send {TAB 2}
	send %ayer%{TAB}
	sleep 3500
	send {TAB 4}{SPACE}{TAB 6}{SPACE}{ENTER}
	sleep 3500
	send {TAB 6}{SPACE}{TAB 4}{SPACE}{ENTER}
	sleep 3500
	send {TAB 9}{ENTER}
	sleep 5000
	MouseMove 2814, 212
	waitColorChange()
	sleep 200
	send {TAB 17}{SPACE}{DOWN}{ENTER}
}


mainMethodStatistics()
{
	openBT()

	navigateStatistics()
}

navigateStatistics()
{
	global
	sleep 2000
	send {TAB 12}{ENTER}
	sleep 2000
	send {TAB 12}{ENTER}
	sleep 3500
	send {TAB 8}
	send %fecha%{TAB}
	sleep 3500
	send {TAB 2}
	send %ayer%{TAB}
	sleep 3500
	send {TAB 3}{DOWN}{TAB}
	sleep 3500
	send {TAB 4}{SPACE}{DOWN}{ENTER}
	sleep 3500
	send {TAB 5}60{TAB}
	sleep 3500
	send {TAB 6}60{TAB}
	sleep 3500
	send {TAB 7}61{TAB}
	sleep 3500
	send {TAB 10}{ENTER}
	sleep 5000
	MouseMove 2336, 335
	waitColorChange()
	sleep 300
	send {TAB 18}{SPACE}{DOWN}{ENTER}
}

mainMethodqagent()
{
	openBT()

	navigateqagent()
}

navigateqagent()
{
	global
	sleep 2000
	send {TAB 12}{ENTER}
	sleep 2000
	send {TAB 15}{ENTER}
	sleep 3500
	send {TAB 8}
	send %fecha%{TAB}
	sleep 3500
	send {TAB 2}
	send %ayer%{TAB}
	sleep 3500
	send {TAB 3}{DOWN}
	sleep 3500
	send {TAB 4}60{TAB}
	sleep 3500
	send {TAB 5}{down}
	sleep 3500
	send {TAB 7}{ENTER}
	sleep 5000
	MouseMove 2323, 276
	waitColorChange()
	sleep 300
	send {TAB 15}{SPACE}{DOWN}{ENTER}
}

mainMethodChat()
{
global
	CoordMode, Mouse, Screen
	MouseClick, left, 1933, 212			;Click in Applications
	waitClick()
	MouseMove, 2212, 209				;Hover over Quick Links
	waitClick()
	MouseClick, left, 2216, 286			;Click in AR System Report Console
	waitClick()
	MouseClick, left, 2358, 182			;Click in first item of list
	send {PgDn 32}
	MouseClick, left, 2186, 227			;Click Anupam Agent Driven Abandon Chat Report
	MouseClick, left, 3705, 320			;Click in Show Additional Filters
	waitClick()
	MouseClick, left, 2189, 568
	MouseClick, left, 2244, 527
	sendraw 'Assigned To'   !=  " "  AND  'Region'  = "EMEA"  AND  'Resolution Status'  =  "Abandoned"  AND 'Request Live agent Time'  >= "
	send %fecha% 12:00:00 AM"
	MouseClick, left, 1941, 317
	MouseMove, 2027, 440
	waitClick()
	MouseClick, left, 1969, 346
	MouseClick, left, 2996, 795
	waitClick()
	send ^j
	waitClick()
	send ^{F4}



	MouseClick, left, 2119, 248
	send {PgUp 4}
	MouseClick, left, 2188, 191			;Click in Alex Agent Driven Resolved Chat
	MouseClick, left, 3705, 320			;Click in Show Additional Filters
	waitClick()
	MouseClick, left, 2244, 527
	sendraw 'Assigned To'   !=  " "  AND  'Region'  = "EMEA"  AND  'Resolution Status'  =  "Resolved via Agent" 

 AND 'Request Live agent Time'  >= "
	send %fecha% 12:00:00 AM"
	MouseClick, left, 1941, 317
	MouseMove, 2027, 440
	waitClick()
	MouseClick, left, 1969, 346
	MouseClick, left, 2996, 795
	waitClick()
	send ^j
	waitClick()
	send ^{F4}



	;MouseClick, left, 3831, 271			;Temporarily
	MouseClick, left, 2237, 204			;Alex User and System Driven Abandoned Chat
	MouseClick, left, 3705, 320			;Click in Show Additional Filters
	waitClick()
	MouseClick, left, 2189, 568
	MouseClick, left, 2244, 527
	sendraw 'Region' = "EMEA" AND 'Resolution Status' = "Abandoned" AND 'Request Live agent Time' >= "
	send %fecha% 12:00:00 AM"
	MouseClick, left, 1941, 317
	MouseMove, 2027, 440
	waitClick()
	MouseClick, left, 1969, 346
	MouseClick, left, 2996, 795
	waitClick()
	send ^j
	waitClick()
	send ^{F4}

	waitClick()
	MouseClick, left, 2077, 145
	MouseClick, left, 2043, 162
	MouseClick, left, 2120, 241
	waitClick()
	MouseClick, left, 3705, 320			;Click in Show Additional Filters
	waitClick()
	MouseClick, left, 2189, 568
	MouseClick, left, 2244, 527
	sendraw 'Request Live agent Time'  >= "
	send %fecha% 12:00:00 AM"
	MouseClick, left, 1941, 317
	MouseMove, 2027, 440
	waitClick()
	MouseClick, left, 1969, 346
	MouseClick, left, 2996, 795
	waitClick()
	send ^j
	waitClick()
	send ^{F4}

}

mainMethodReported()
{
	global
	CoordMode, Mouse, Screen
	
	;Set File | CSV | Available Fields | etc
	MouseClick, left, 2107, 540
	MouseClick, left, 2107, 540
	MouseClick, left, 2107, 540
	MouseClick, left, 2107, 540	;Click scroll down 7 times to see Reported Date+ field
	MouseClick, left, 2107, 540
	MouseClick, left, 2107, 540
	MouseClick, left, 2107, 540
	MouseClick, left, 2019, 453	;Click on Reported Date+ (To select it)
	MouseClick, left, 2199, 361	;Click in Add Field
	MouseClick, left, 2531, 391	;Click in is equal (To set it to >= )
	MouseClick, left, 2583, 456	;Set it to  >=
	MouseClick, left, 2822, 395	;Click in Calendar icon
	MouseClick, left, 2972, 432	;Set Calendar to January (Default)
	Mouseclick, left, 2900, 452

	;Selecting Month
	MouseClick, left, 2972, 432	
	If(fecha = "January")
	{
		Mouseclick, left, 2900, 452
		sleep 100
		MouseClick, left, 2838, 481
	}	
	else if(fecha = "February")
	{
		Mouseclick, left, 2900, 467
		sleep 100
		MouseClick, left, 2931, 484
	}
	else if(fecha = "March")
	{
		Mouseclick, left, 2900, 480
		sleep 100
		MouseClick, left, 2931, 484
	}
	else if(fecha = "April")
	{
		Mouseclick, left, 2900, 497
		sleep 100
		MouseClick, left, 3015, 483
	}
	else if(fecha = "May")
	{
		Mouseclick, left, 2900, 515
		sleep 100
		MouseClick, left, 2874, 483
	}
	else if(fecha = "June")
	{
		Mouseclick, left, 2900, 528
		sleep 100
		MouseClick, left, 2968, 480
	}
	else if(fecha = "July")
	{
		Mouseclick, left, 2900, 545
		MouseClick, left, 3017, 486
	}
	else if(fecha = "August")
	{
		Mouseclick, left, 2900, 559
		sleep 100
		MouseClick, left, 2900, 482
	}
	else if(fecha = "September")
	{
		Mouseclick, left, 2900, 578
		sleep 100
		MouseClick, left, 2994, 479
	}
	else if(fecha = "October")
	{	
		Mouseclick, left, 2900, 592
		sleep 100
		MouseClick, left, 2845, 481
	}
	else if(fecha = "November")
	{
		Mouseclick, left, 2900, 611
		sleep 100
		MouseClick, left, 2932, 481
	}
	else if(fecha = "December")
	{	
		Mouseclick, left, 2900, 624
		sleep 100
		MouseClick, left, 2993, 478
	}
	else
		Msgbox The month provided is incorrect
	;ACEPT DATE
	MouseClick, left, 2898, 641
	
	;------------------------------
	MouseClick, left, 2019, 453	;Click on Reported Date+ (To select it)
	MouseClick, left, 2199, 361	;Click in Add Field
	MouseClick, left, 2528, 429	;Click in is equal (To set it to >= )
	MouseClick, left, 2581, 526	;Set it to  <=
	MouseClick, left, 2823, 431	;Click in Calendar icon
	MouseClick, left, 2972, 471	;Set Calendar to January (Default)
	Mouseclick, left, 2900, 491

	;Selecting Month
	MouseClick, left, 2972, 470	
	If(fecha = "January")
	{
		Mouseclick, left, 2900, 491
		sleep 100
		MouseClick, left, 2900, 581
	}	
	else if(fecha = "February")
	{
		Mouseclick, left, 2900, 503
		sleep 100
		MouseClick, left, 2900, 581
	}
	else if(fecha = "March")
	{
		Mouseclick, left, 2900, 522
		sleep 100
		MouseClick, left, 2993, 581
	}
	else if(fecha = "April")
	{
		Mouseclick, left, 2900, 539
		sleep 100
		MouseClick, left, 2839, 595
	}
	else if(fecha = "May")
	{
		Mouseclick, left, 2900, 554
		sleep 100
		MouseClick, left, 2934, 581
	}
	else if(fecha = "June")
	{
		Mouseclick, left, 2900, 567
		sleep 100
		MouseClick, left, 2996, 581
	}
	else if(fecha = "July")
	{
		Mouseclick, left, 2900, 587
		sleep 100
		MouseClick, left, 2874, 598
	}
	else if(fecha = "August")
	{
		Mouseclick, left, 2900, 601
		sleep 100
		MouseClick, left, 2966, 581
	}
	else if(fecha = "September")
	{
		Mouseclick, left, 2900, 618
		sleep 100
		MouseClick, left, 3018, 581
	}
	else if(fecha = "October")
	{	
		Mouseclick, left, 2900, 632
		sleep 100
		MouseClick, left, 2905, 584
	}
	else if(fecha = "November")
	{
		Mouseclick, left, 2900, 646
		sleep 100
		MouseClick, left, 2966, 579
	}
	else if(fecha = "December")
	{	
		Mouseclick, left, 2900, 663
		sleep 100
		MouseClick, left, 2840, 598
	}
	else
		Msgbox The month provided is incorrect
	;ACEPT DATE
	MouseClick, left, 2898, 681	
	
	MouseClick, left, 2081, 633	;Click in Destination
	MouseClick, left, 2020, 664	;Click in File
	MouseClick, left, 2235, 629	;Click in Format
	MouseClick, left, 2165, 661	;Click in CSV
	MouseClick, left, 2190, 571	;Click in Advanced




	;=========DOWNLOAD DUTCH REPORT============;
	MouseClick, left, 2336, 524
	Sendraw ('Submitter*' = "I793776") OR 
	sendraw ('Submitter*' = "S972800") OR 
	sendraw ('Submitter*' = "T867312") OR 
	sendraw ('Submitter*' = "M239565") OR 
	sendraw ('Submitter*' = "K984621") OR 
	sendraw ('Submitter*' = "M396541")
	send {tab 5}
	send ^a
	send 01_Dutch_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	MouseMove, 2024, 1060
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD FRENCH REPORT============;	
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Submitter*' = "A885930") OR  
	sendraw ('Submitter*' = "A079482") OR 
	sendraw ('Submitter*' = "F340923") OR 
	sendraw ('Submitter*' = "V510056")
	send {tab 5}
	send ^a
	send 01_French_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	MouseMove, 2024, 1060
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD GERMAN REPORT============;	
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Submitter*' = "G045601") OR 
	sendraw ('Submitter*' = "N069127") OR 
	sendraw ('Submitter*' = "I947801") OR 
	sendraw ('Submitter*' = "F743921")
	send {tab 5}
	send ^a
	send 01_German_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	MouseMove, 2024, 1060
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD POLISH REPORT============;	IT BREAKS ON POLISH

	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Submitter*' = "A522461") OR 
	sendraw ('Submitter*' = "A697516") OR 
	sendraw ('Submitter*' = "A899684") OR 
	sendraw ('Submitter*' = "P634176")
	send {tab 5}
	send ^a
	send 01_Polish_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315
	
	;=========DOWNLOAD RUSSIAN REPORT============;	
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Submitter*' = "A132708") OR 
	sendraw ('Submitter*' = "B885997") OR 
	sendraw ('Submitter*' = "F646521") OR 
	sendraw ('Submitter*' = "I871455") OR 
	sendraw ('Submitter*' = "J069431") OR 
	sendraw ('Submitter*' = "M679429")
	send {tab 5}
	send ^a
	send 01_Russian_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD MIXED REPORT============;	
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Submitter*' = "C446029") OR 
	sendraw ('Submitter*' = "R321861") OR 
	sendraw ('Submitter*' = "E402152") OR 
	sendraw ('Submitter*' = "M886061") OR 
	sendraw ('Submitter*' = "D079533") OR 
	sendraw ('Submitter*' = "U318158") OR 
	sendraw ('Submitter*' = "H797054") OR 
	sendraw ('Submitter*' = "S901062") 
	send {tab 5}
	send ^a
	send 01_Mixed_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315
	
	MouseClick, left, 2041, 466	;Click in Reported Source
	MouseClick, left, 2200, 363	;Click in Add Field
	MouseClick, left, 2945, 452	;Click to scroll down
	MouseClick, left, 2823, 444	;Click in Dropdown menu (WEB|MAIL)
	MouseClick, left, 2843, 607	;Click in WEB
	

	;=========DOWNLOAD DUTCH GROUP REPORT============;	HAS TO BE WEB ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Dutch")
	send {tab 5}
	send ^a
	send 02_HDCC_Dutch_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD GERMAN REPORT============;		HAS TO BE WEB ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-German") 
	send {tab 5}
	send ^a
	send 02_HDCC_German_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD FRENCH REPORT============;		HAS TO BE WEB ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-French") 
	send {tab 5}
	send ^a
	send 02_HDCC_French_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD RUSSIAN REPORT============;		HAS TO BE WEB ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Russian") 
	send {tab 5}
	send ^a
	send 02_HDCC_Russian_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315
	
	;=========DOWNLOAD HUN-ROM REPORT============;		HAS TO BE WEB ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Hungarian") OR
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Romanian")
	send {tab 5}
	send ^a
	send 02_HDCC_HUG-ROM_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD POLISH REPORT============;		HAS TO BE WEB ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Polish") 
	send {tab 5}
	send ^a
	send 02_HDCC_Polish_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD MIXED REPORT============;		HAS TO BE WEB ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Italian") OR 
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Spanish") OR 
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Portuguese")
	send {tab 5}
	send ^a
	send 02_HDCC_Mixed_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD TURKISH REPORT============;		HAS TO BE WEB ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Turkish")
	send {tab 5}
	send ^a
	send 02_HDCC_Turkish_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	MouseClick, left, 2821, 448	;Click in dropdown menu
	sleep 50
	MouseClick, left, 2854, 479	;Click in EMAIL

	;=========DOWNLOAD FRENCH REPORT============;		HAS TO BE MAIL ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-French")
	send {tab 5}
	send ^a
	send 03_MAIL_French_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD RUSSIAN REPORT============;		HAS TO BE MAIL ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Russian")
	send {tab 5}
	send ^a
	send 03_MAIL_Russian_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD HUN-ROM REPORT============;		HAS TO BE MAIL ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Hungarian") OR
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Romanian")
	send {tab 5}
	send ^a
	send 03_MAIL_Hun-Rom_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD POLISH REPORT============;		HAS TO BE MAIL ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Polish")
	send {tab 5}
	send ^a
	send 03_MAIL_Polish_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD MIXED REPORT============;		HAS TO BE MAIL ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Italian") OR
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Spanish") OR
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Portuguese")
	send {tab 5}
	send ^a
	send 03_MAIL_Mixed_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315

	;=========DOWNLOAD TURKISH REPORT============;		HAS TO BE MAIL ONLY
	MouseClick, left, 2336, 524
	send ^a
	send {del}
	Sendraw ('Owner Group+'  = "Infra-SD-EMEA-Turkish")
	send {tab 5}
	send ^a
	send 03_MAIL_Turkish_Submitted_%fecha%.csv
	MouseClick, left, 1955, 312
	waitClick()
	send ^j
	waitClick()
	Send {Shift down}{LCtrl down}{Tab}{LCtrl up}{Shift up}
	MouseClick, left, 3729,315


}




tempMove()
{
	CoordMode, Mouse, Screen
	MouseClick, left, 2929, 411
}

waitClick()
{
	Input, SingleKey, L1, {right}
}

writeToFileInc()
{
	global
	FileAppend, %incident1%`n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, %incident2%`n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, %incident3%`n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, %incident4%`n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
}

writeToFileSer()
{
	global
	FileAppend, %distribution1%`n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, %distribution2%`n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, %distribution3%`n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, %distribution4%`n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
}