Pause::Suspend
return

#NumPad1::  			;Press Windows key + 1 on the Numpad to activate script
MouseClick, left, 532, 529 ;Impact
sleep 2000
MouseClick, left, 475, 592 ;Minor/Localized
sleep 2000
MouseClick, left, 535, 556 ;Urgency
sleep 2000
MouseClick, left, 454, 622 ;Low
sleep 2000
MouseClick, left, 532, 631 ; Reported Source
sleep 2000
MouseClick, left, 444, 745 ; Phone
sleep 2000
MouseClick, left, 409, 672 ; Click on Assigned Group
Send, infra-sd-emea-spanish ; Type Spanish team
sleep 2000
send {down}
send {Enter}
MouseClick, left, 405, 701 ;Click on Assignee field
sleep 2000
send, Carlo
sleep 2000
send {down}
send {Enter}
MouseClick, left, 950, 222 ;Date/System
sleep 2000
MouseClick, left, 954, 578 ;Click on Owner group field
sleep 2000
Send, infra-sd-emea-spanish ; Type Spanish team
sleep 2000
send {down}
send {Enter}
MouseClick, left, 892, 604 
sleep 2000
send, Carlo
sleep 2000
send {down}
send {Enter}
MouseClick, left, 764, 221 ;Categorization Tab
MouseClick, left, 874, 503
send, generic
sleep 2000
send {Enter}
MouseClick, left, 933, 284
MouseClick, left, 874, 503
send, generic
sleep 2000
send {Enter}
MouseClick, left, 949, 310
send, no
sleep 2000
send {down}
send {Enter}
MouseClick, left, 883, 331 ; Tier 2 - Resolution Categorization
send, no
sleep 2000
send {down}
send {Enter}
MouseClick, left, 874, 364
send, no
sleep 2000
send {down}
send {Enter}
MouseClick, left, 533, 772 ; Status
sleep 2000
MouseClick, left, 529, 830 ; Status to In Progress
MouseClick, left, 468, 278 ; Click on customer
StringUpper Clipboard, clipboard ;Convert Clipboard content to UpperCase
Send %Clipboard% ;Paste Clipboard
sleep 2000
send {down}
send {Enter}
MouseClick, left, 995, 286 ; Go to Operational Categorization
MouseClick, left, 930, 315 ; Click on Tier 1
send, b			; for Business Application
sleep 2000
send {down}
send {Enter}
MouseClick, left, 881, 339
send, b			; for Business Unit
sleep 2000
send {down}
send {Enter}
CoordMode, Mouse, Screen
Click 522, 286		;copy Telefone number
Click 522, 286
Click 522, 286
sleep 2000
send ^c
Click 1166, 243
sleep 2000
send {SPACE}
send ^v
Click 187, 324		;Copy City to put on notes
Click 187, 324		
send ^c
Click 1131, 412		;Paste City
send {SPACE}
send ^v
send {,}
send {SPACE}
Click 186, 354		;Copy Country to put on notes
Click 186, 354
send ^c
Click 1159, 410
send ^v
Click 1074, 348
Send {SPACE}
FormatTime, CurrentDateTime,, dd-MM-yy HH:mm
SendInput %CurrentDateTime%
Click 1572, 996
Send ^a
Send ^c
Click 2237, 342
sleep 2000
Click 2416, 559
send ^v
Return