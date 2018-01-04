#Escape:: 
Pause ;ExitApp 
Return

#8::
MouseClick, left, 564, 347   ; Open Notes window
sleep 400
MouseClick, left, 683, 472  ; Click anywhere on the notes
sleep 100
send ^a			    ; Select all from notes
sleep 100
send ^c			    ; Copy all from notes
sleep 100		  
send ^t			    ; Open new Chrome Tab
sleep 200
send translate.google.com   ; Enter URL to Google Translate
sleep 50
send {ENTER}
sleep 4000
MouseClick, left, 239, 296  ; Click on left box to put text to be translated
sleep 500
send ^v			    ; Paste text to be translated
MouseClick, Left, 914, 214  ;Press Button "Translate"
sleep 500
MouseClick, left, 1251, 316 ; First Click
sleep 500
MouseClick, left, 1251, 316 ; Second Click (To highlight all text)
sleep 500
send ^c
sleep 50
send ^1
MouseClick, left, 683, 472  ; Click anywhere on the notes
send ^{HOME}		    ; Go to first line of notes
send {ENTER}{ENTER}	    ; Make space at the first line
send ^{HOME}		    ; After the spacing go back to first line
send ENGLISH:
send {ENTER}
send ^v
send {ENTER}{ENTER}
SendRaw ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
send {ENTER}{ENTER}
send SPANISH: 