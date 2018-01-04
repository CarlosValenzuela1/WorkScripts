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

;Replaces Clipoard with first value of table
ClipBoard := SubPat1
;MsgBox %temp%
;MsgBox %SubPat1%
;MsgBox %SubPat2%
;MsgBox %SubPat3%
;MsgBox %SubPat4%
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

return

;===============================================SECOND PART================================================;
^F2::
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


return
