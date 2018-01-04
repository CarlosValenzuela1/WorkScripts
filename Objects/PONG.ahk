#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance force
;#Warn All
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

PlayGame:
x1 := A_ScreenWidth/2
y1 := A_ScreenHeight/2
Ball_ImgW = 60
Ball_ImgH = 60
ScrnW := A_ScreenWidth-Ball_ImgW
ScrnH := A_ScreenHeight-Ball_ImgH

p1_pad_W = 13
p1_pad_H = 93
p1_pad_h_mid := p1_pad_H/2
p2_pad_W = 13
p2_pad_H = 93
p2_pad_h_mid := p2_pad_H/2

p1_pad_x := 20+p1_pad_w/2
p1_pad_y := y1-p1_pad_H/2

p2_pad_x := A_ScreenWidth-(20+p1_pad_w/2)
p2_pad_y := y1-p2_pad_H/2

P1Color = Red
P2Color = Blue

P1Scorexpos := x1-200
P2Scorexpos := x1+100

p1_Score = 0
p2_Score = 0

ScoreToWin = 10

Random, GoSouth, 1, 2
Random, GoWest, 1, 2


ball_speeder = 20
p1_speeder = 30
p2_speeder = 30

Gosub, Filip


loop 
{
WinMove, the_ball,, %x1%, %y1%

Random, p2_bored, 1, 10
if (GoWest =1 and p2_bored <> 1)
{
 
 if (p2_pad_y+p2_pad_h_mid > y1 and y1 > A_ScreenHeight*.1)
 {
   p2_pad_y := p2_pad_y-(1*p2_speeder)
   WinMove, p2_paddle,, %p2_pad_x%, %p2_pad_y%
 }
 else if (p2_pad_y+p2_pad_h_mid < y1 and y1 < A_ScreenHeight*.9)
 {
   p2_pad_y := p2_pad_y+(1*p2_speeder)
   WinMove, p2_paddle,, %p2_pad_x%, %p2_pad_y%
 }

}


if (p1_pad_y < y1 and p1_pad_y+p1_pad_H > y1)

{
;msgbox inbetween p1 horizontally x1: %x1%, value: %p1_pad_x%+%p1_padw%

 if(p1_pad_x+p1_pad_W > x1)
 {
   ;msgbox HIT PLAYER 1, BOUNCE!!
   GoWest = 1
 }
}


if (p2_pad_y < y1 and p2_pad_y+p2_pad_H > y1)
{
 ;msgbox inbetween p2 horizontally

 if(p2_pad_x < x1+Ball_ImgW)
 {
   ;msgbox HIT PLAYER 2, BOUNCE!!
   GoWest = 0
 }
}


if (x1 < 2)
{
 GoWest = 1
 p2_Score++
}
if (x1 > ScrnW)
{
 GoWest = 0
 p1_Score++
}


if (y1 < 2)
 GoSouth = 1
if (y1 > ScrnH)
 GoSouth = 0

if (GoWest = 1)
 x1 := x1+(1*ball_speeder)
else
 x1 := x1-(1*ball_speeder)

if (GoSouth = 1)
 y1 := y1+(1*ball_speeder)
else
 y1 := y1-(1*ball_speeder)
}
 
return


Filip:
Gui, 1:Color, red,
Gui, 1:Add, Picture, X0 Y0 H%Ball_ImgH% W%Ball_ImgW%, C:\Users\%A_UserName%\Desktop\Filip.jpg
Gui, 1:-Caption +ToolWindow +Disabled +AlwaysOntop +LastFound
Gui, 1:Show, X%x1% Y%y1% W%Ball_ImgW% H%Ball_ImgH%, the_ball
return


~Up::
	y1 := y1-10
	WinMove, the_ball,, %x1%, %y1%
return 

~Down::
	y1 := y1+10
	WinMove, the_ball,, %x1%, %y1%
return 

~Left::
	x1 := x1-10
	WinMove, the_ball,, %x1%, %y1%
return 

~Right::
	x1 := x1+10
	WinMove, the_ball,, %x1%, %y1%
return 
