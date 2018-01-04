#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


^F1::
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
return

