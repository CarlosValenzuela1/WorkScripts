;=========================================================================================================;
; COM Objects Tutorial								  			  ; 
;=========================================================================================================;
; Description of each method will be find below.							  ;
; Version: 1.0												  ;
; Author: Carlos Valenzuela										  ;
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



Pause::
InputBox, Var, Reports Script, • SD Incident Distribution "inc" `n• SD Service Request "sr" `n• Misrouted Incidents "miss1" `n• Missrouted Service Request "miss2" `n• EUC Incident Distrubition "eucinc"`n• EUC Service Request "eucsr"`n• Agent State "state"`n• Queue Statistics "statistics"  `n• Queues by Agent "qagent"  `n• Test "test" , , 250, 300

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
else if (Var = "test")
{
	;InputBox, fecha, Date, START type the date of the report you would like to get: "MM/DD/YYYY" , , 450, 125
	;filterDate()
	Msgbox Hello Csaba
	
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
	waitColorChange()
	waitColorChange()
	sleep 3000
	send {TAB}
	
	send, c760912{tab}Hola1234{enter}
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
	sleep 600
	MouseClick, left, 2991, 557 			;It clicks on the dropdown menu
	sleep 600
	MouseClick, left, 2964, 609 			;Select SD OPEN REPORT
	sleep 600
	MouseClick, left, 2884, 640  			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	waitColorChange()				;Wait for color change
	MouseClick, left, 3806, 255			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 2951, 419			;Clicks on the date field
	sleep 500					;Waits for popup to appear

	filterDate()



	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 1995, 174			;Clicks on Table icon (Upper-left side on Netasthra)
	sleep 2000
	MouseClick, left, 2495, 851			;Clicks on position that is empty to be able to use next shortcuts
	sleep 300
	send ^a						;Select All text on webpage
	sleep 700
	send ^c						;Copy All text on webpage
	sleep 400
	StartPos := RegExMatch(ClipBoard, "Total:\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})", Inc)
							;Find the values after "Total: " and stores them in Inc array
	
	MouseClick, left, 3114, 388			;Clicking on the Total cell of the Incident Distribution table.
	MouseMove, 2887, 522				;Moving mouse to activate color change
	waitColorChange()
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 1000
	MouseClick, left, 2001, 226			;Click on the Eye (add/remove columns)
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
	sleep 600
	MouseClick, left, 2991, 557 			;It clicks on the dropdown menu
	sleep 600
	MouseClick, left, 2869, 612 			;Select SD OPEN REPORT
	sleep 600
	MouseClick, left, 2884, 640  			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	MouseMove, 2888, 522
	waitColorChange()				;Wait for color change
	MouseClick, left, 3806, 255			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 2951, 419			;Clicks on the date field
	sleep 500					;Waits for popup to appear

	filterDate()



	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 1995, 174			;Clicks on Table icon (Upper-left side on Netasthra)
	sleep 2000
	MouseClick, left, 2495, 851			;Clicks on position that is empty to be able to use next shortcuts
	sleep 300
	send ^a						;Select All text on webpage
	sleep 700
	send ^c						;Copy All text on webpage
	sleep 400
	StartPos := RegExMatch(ClipBoard, "Total:\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})\s([0-9]{1,})", Dis)
							;Find the values after "Total: " and stores them in Inc array
	
	MouseClick, left, 3114, 388			;Clicking on the Total cell of the Incident Distribution table.
	MouseMove, 2887, 522				;Moving mouse to activate color change
	waitColorChange()
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 1000
	MouseClick, left, 2001, 226			;Click on the Eye (add/remove columns)
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
	sleep 600
	MouseClick, left, 2991, 557 			;It clicks on the dropdown menu
	sleep 600
	MouseClick, left, 2915, 634			;Select EUC OPEN REPORT
	sleep 600
	MouseClick, left, 2884, 640  			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	waitColorChange()				;Wait for color change
	MouseClick, left, 3806, 255			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 2951, 419			;Clicks on the date field
	sleep 500					;Waits for popup to appear

	filterDate()



	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 3742, 271			;Clicks on Table icon (Upper-left side on Netasthra)
	
	MouseMove, 2887, 522				;Moving mouse to activate color change
	waitColorChange()
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 1000
	MouseClick, left, 2001, 226			;Click on the Eye (add/remove columns)
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
	Send Owner Group Name				;Types Owner Group Name
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
	Send Resolution Target Date			;Types Resolution Target Date
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
	Send Region					;Types Region
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
	Send Business Unit				;Types Business Unit
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300		
	Send Group Change Count				;Types Group Change Count
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
	sleep 600
	MouseClick, left, 2991, 557 			;It clicks on the dropdown menu
	sleep 600
	MouseClick, left, 2958, 634 			;Select EUC OPEN REPORT
	sleep 600
	MouseClick, left, 2884, 640  			;Click in Okay option.
	sleep 100					;Wait for animation to finish
	waitColorChange()				;Wait for color change
	MouseClick, left, 3806, 255			;Clicks on Filter options
	sleep 500					;Waits till filter option opens
	send {TAB 4}					;Types TAB four times
	sleep 100					;Wait for changes to be reflected on screen
	send {down}					;selects the dropdown option (SYMBOL <= or =>)
	MouseClick, left, 2951, 419			;Clicks on the date field
	sleep 500					;Waits for popup to appear

	filterDate()


	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	sleep 1000					;Wait 1 second and activate the color change
	waitColorChange()
	MouseClick, left, 3744, 271			;Clicks on Table icon (Upper-left side on Netasthra)
	
	MouseMove, 2887, 522				;Moving mouse to activate color change
	waitColorChange()
	waitColorChange() 				;Maybe it requires two color changes if speed is fast
	sleep 1000
	MouseClick, left, 2001, 226			;Click on the Eye (add/remove columns)
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
	Send Owner Group Name				;Types Owner Group Name
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
	Send Resolution Target Date			;Types Resolution Target Date
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
	Send Region					;Types Region
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
	Send Business Unit				;Types Business Unit
	sleep 300
	send {Tab}{Space}
	sleep 300
	send {Shift down}{Tab}{Shift up}
	sleep 300		
	Send Group Change Count				;Types Group Change Count
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

	filterDate()

	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	sleep 1000					;Wait 1 second and activate the color change
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

	filterDate()

	CoordMode, Mouse, Screen			;Returns focus to both monitors
	MouseClick, left, 2890, 869			;Click in Done
	sleep 1000					;Wait 1 second and activate the color change
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
	send {TAB 12}{ENTER}
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
	send {TAB 4}{SPACE}{DOWN}{ENTER}
	sleep 3500
	send {TAB 5}60{TAB}
	sleep 3500
	send {TAB 6}60{TAB}
	sleep 3500
	send {TAB 7}60{TAB}
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

writeToFileInc()
{
	global
	FileAppend, Incident Assigned: %incident1% `n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, Incident In Progress: %incident2% `n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, Incident Pending: %incident3% `n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, Incident Resolved: %incident4% `n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
}

writeToFileSer()
{
	global
	FileAppend, Service Request Assigned: %distribution1% `n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, Service Request In Progress: %distribution2% `n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, Service Request Pending: %distribution3% `n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
	FileAppend, Service Request Resolved: %distribution4% `n, C:\Users\%A_UserName%\Desktop\KPI %A_MM%-%A_DD%-%A_YYYY%.txt
}