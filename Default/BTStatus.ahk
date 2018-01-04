#3::
Gui, Color, black
Gui +AlwayOnTop
Gui -Caption
Gui, Show, x0 y0 w%A_ScreenWidth% h%A_ScreenHeight%
return
;--- switch ON/OFF with F12 ---
F12::
V++
M:=mod(V,2)
if M=1
Gui, Show,minimize
else
Gui, Show,
return

esc::exitapp