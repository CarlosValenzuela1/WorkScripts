#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Open Snipping Tool when typing Ctrl + Printscn
^PrintScreen:: RunWait, %A_WinDir%\system32\SnippingTool.exe

~LButton & RButton::
{
; MsgBox, testing that
MyInt = 0
WinActivate, Snipping Tool
sleep 500
send !F
send a
sleep 500
if 0 = %MyInt%
{
	send  Capture%MyInt%
}
else
{
MyInt = 1
	send Capture%MyInt%
}

  return

}
