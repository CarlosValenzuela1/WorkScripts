#NumPad1:: 
MouseClick, left, 532, 529 ;Impact
sleep 400
MouseClick, left, 475, 592 ;Minor/Localized
sleep 400
MouseClick, left, 535, 556 ;Urgency
sleep 400
MouseClick, left, 454, 622 ;Low
sleep 400
MouseClick, left, 532, 631 ; Reported Source
sleep 400
MouseClick, left, 444, 745 ; Phone
sleep 400
MouseClick, left, 409, 672 ; Click on Assigned Group
Send, infra-sd-emea-spanish ; Type Spanish team
sleep, 1200
send {down}
send {Enter}
MouseClick, left, 405, 701 ;Click on Assignee field
sleep, 400
send, Carlo
sleep, 1200
send {down}
send {Enter}
MouseClick, left, 950, 222 ;Date/System
sleep, 400
MouseClick, left, 954, 578 ;Click on Owner group field
sleep, 400
Send, infra-sd-emea-spanish ; Type Spanish team
sleep, 1200
send {down}
send {Enter}
MouseClick, left, 892, 604 
sleep, 400
send, Carlo
sleep, 1200
send {down}
send {Enter}
MouseClick, left, 764, 221 ;Categorization Tab
MouseClick, left, 874, 503
send, generic
sleep, 400
send {Enter}
MouseClick, left, 933, 284
MouseClick, left, 874, 503
send, generic
sleep, 400
send {Enter}
MouseClick, left, 949, 310
send, no
sleep, 1200
send {down}
send {Enter}
MouseClick, left, 883, 331 ; Tier 2 - Resolution Categorization
send, no
sleep, 1200
send {down}
send {Enter}
MouseClick, left, 874, 364
send, no
sleep, 1200
send {down}
send {Enter}
MouseClick, left, 533, 772 ; Status
sleep 100
MouseClick, left, 529, 830 ; Status to In Progress
MouseClick, left, 468, 278 ; Click on customer
StringUpper Clipboard, clipboard ;Convert Clipboard content to UpperCase
Send %Clipboard% ;Paste Clipboard
sleep, 1200
send {down}
send {Enter}
MouseClick, left, 995, 286 ; Go to Operational Categorization
MouseClick, left, 930, 315 ; Click on Tier 1
send, b			; for Business Application
sleep, 1200
send {down}
send {Enter}
MouseClick, left, 881, 339
send, b			; for Business Unit
sleep, 1200
send {down}
send {Enter}
MouseClick, left, 438, 293
MouseClick, left, 438, 293
^c
MouseClick, left, 183, 242
^v
Return