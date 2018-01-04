SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force
#Persistent
#MaxThreadsPerHotkey 2
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.

Scrolllock::Pause
PrintScreen::Suspend
Insert::Reload


$^c::
Clip := []
clipboard =
sleep 200
SendInput, {ctrl down}c{ctrl up}
clipwait, 1
Text := Clipboard
loop, parse, text, `n, `r
{
    if (A_LoopField = "") or RegExMatch(A_LoopField, "^@")
        continue
    thisline++
    Clip[thisline] := A_LoopField
}
thisline := ""
return

#v::
loop{
line++
if !Clip[line] {
    line := 1
    break
}

clipboard := Clip[line]
send ^v
send {;}{SPACE}
sleep 50
}
return