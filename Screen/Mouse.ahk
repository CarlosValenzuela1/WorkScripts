MButton::
   If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500) {
      Sleep 100   ; wait for right-click menu, fine tune for your PC
      Send tcs{#}1234  ; close it

   }

Return