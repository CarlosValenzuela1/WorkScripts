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
wb := ComObjCreate("InternetExplorer.Application") ;create a IE instance
wb.Visible := True
wb.Navigate("http://itreporting.cargill.com/Netasthra/loginform")


while wb.ReadyState != 4 || wb.document.readystate != "complete" || wb.busy
    sleep 10


msgbox test
wb.Document.All.userName.Value := "C760912"
wb.Document.All.password.Value := "Hola1234"
;wb.Document.All.password.focus()

;send {ENTER}



wb.Document.GetElementById("Incident Distribution").Click()

;wb.quit

return
