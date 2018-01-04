#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
^!I::
Send ^c
ClipWait, 2
google_url:="http://translate.google.com/#en/es/" clipboard
Run %google_url%
sleep 200
send ^c
send {alt}{tab}
sleep 200
send ^v
Return

Escape:: 
Pause ;ExitApp 
Return