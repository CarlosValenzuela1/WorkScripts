#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its sCarloserior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

^F1::
;MouseClick, left, 531, 669	;relative
MouseClick, left, 523, 661	;Client UNSELECT ALL
sleep 2000
MouseClick, left, 519, 701	;Click on Search bar
sleep 2000
send, emea-central service desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-Central Hungary Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-Central Poland Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-Central Romania Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-Central Turkey Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-Central West Africa Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-North Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-East Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 100
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}
sleep 500
send, EMEA-South Service Desk
sleep 2001
MouseClick, left, 520, 742	;Click on first item on list
sleep 500
MouseClick, left, 519, 701	;Click on Search bar
send {ctrl down}a{ctrl up}

sleep 1000
send, Infra-SD-EMEA-
sleep 4000
MouseClick, left, 520, 742	;Click on first item on list
sleep 10
MouseClick, left, 514, 772	;Click on second item on list
sleep 10
MouseClick, left, 514, 802	;Click on third item on list
sleep 10
MouseClick, left, 514, 832	;Click on fourth item on list
sleep 10
MouseClick, left, 514, 862	;Click on fifth item on list
sleep 10
MouseClick, left, 514, 892	;Click on sixth item on list
sleep 10
MouseClick, left, 514, 922	;Click on seventh item on list
sleep 10
MouseClick, left, 514, 952	;Click on eighth item on list
sleep 10
MouseClick, left, 514, 982	;Click on nineth item on list
sleep 10
MouseClick, left, 514, 1012	;Click on tenth item on list
sleep 10
MouseClick, left, 514, 1042	;Click on eleventh item on list
Click WheelDown			;will scroll down
sleep 10
MouseClick, left, 514, 1060	;Click on twelveth item on list
return