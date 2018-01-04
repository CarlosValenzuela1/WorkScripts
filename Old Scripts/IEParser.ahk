#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

^F1::
;;;;;;;;;;;;;;;;;;TO COPY VALUES FROM THE INCIDENT DISTRIBUTION TABLE TO EXCELL
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
StartPos := RegExMatch(ClipBoard, "Total:</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""17"" aria-describedby=""list1_Assigned"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""20"" aria-describedby=""list1_In Progress"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""26"" aria-describedby=""list1_Pending"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""324"" aria-describedby=""list1_Resolved"">([0-9]{1,})</td><td role=""gridcell"" style=""text-align:right;"" class=""link-tag"" title=""387"" aria-describedby=""list1_Total"">387</td></tr></tbody></table>", SubPat)
sleep 200
ClipBoard := SubPat1
;MsgBox %temp%
;MsgBox %SubPat1%
;MsgBox %SubPat2%
;MsgBox %SubPat3%
;MsgBox %SubPat4%
send ^v
sleep 1000

MouseClick, left, 800, 830
sleep 700
MouseClick, left, 800, 830
sleep 700
ClipBoard := SubPat2
sleep 200
send ^v


MouseClick, left, 808, 844
sleep 700
MouseClick, left, 808, 844
sleep 700
ClipBoard := SubPat3
sleep 200
send ^v


MouseClick, left, 797, 861
sleep 700
MouseClick, left, 797, 861
sleep 700
ClipBoard := SubPat4
sleep 200
send ^v


/*
EVERYTHING HERE IS JUST A COMMENT
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
*/
return
;===============================================SECOND 

PART================================================;
^F2::
;;;;;;;;;;;;;;;;;;TO COPY VALUES FROM THE SERVICE REQUEST TABLE TO EXCELL
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

return
 

^F3::
send {ctrl down}l{ctrl up}
send ^c
url=http://itreporting.cargill.com/Netasthra/loginform.action
Gui,Add,Edit,w800 h500 -Wrap,% URLDownloadToVar(url)
Gui,show,
return


URLDownloadToVar(url){
	hObject:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
	hObject.Open("GET",url)
	hObject.Send()
	return hObject.ResponseText
}
