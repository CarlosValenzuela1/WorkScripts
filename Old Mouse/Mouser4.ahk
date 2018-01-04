#Persistent
#SingleInstance Force
#InstallKeybdHook
;#NoTrayIcon

~LCtrl & RCtrl::
if toggle := !toggle
{
	Hotkey, Up, Up, On
	Hotkey, Right, Right, On
	Hotkey, Left, Left, On
	Hotkey, Down, Down, On
	Hotkey, Shift, Shift, On
	Hotkey, Enter, Enter, On
}
else
{
	Hotkey, Up, Up, Off
	Hotkey, Right, Right, Off
	Hotkey, Left, Left, Off
	Hotkey, Down, Down, Off
	Hotkey, Shift, Shift, Off
	Hotkey, Enter, Enter, Off
}
Return

Up:
Loop
{
	GetKeyState, PlaceU, Up, P
	GetKeyState, Place2, Right, P
	GetKeyState, Place3, Left, P
	GetKeyState, PlaceLC, LControl, P
	if PlaceU = D &
	{
		if Place2 = D
			MouseMove, 6, -6, 0, R
		else if Place3 = D
			MouseMove, -6, -6, 0, R
		else
			MouseMove, 0, -6, 0, R
	}
	else if PlaceU = U
		Break
}
Return

Right:
Loop
{
	GetKeyState, PlaceR, Right, P
	GetKeyState, Place4, Up, P
	GetKeyState, Place5, Down, P
	if PlaceR = D
	{
		if Place4 = D
			MouseMove, 6, -6, 0, R
		else if Place5 = D
			MouseMove, 6, 6, 0, R
		else
			MouseMove, 6, 0, 0, R
	}
	else if PlaceR = U
		Break
}
Return

Left:
Loop
{
	GetKeyState, PlaceL, Left, P
	GetKeyState, Place6, Up, P
	GetKeyState, Place7, Down, P
	if PlaceL = D
	{
		if Place6 = D
			MouseMove, -6, -6, 0, R
		else if Place7 = D
			MouseMove, -6, 6, 0, R
		else
			MouseMove, -6, 0, 0, R
	}
	else if PlaceL = U
		Break
}
Return

Down:
Loop
{
	GetKeyState, PlaceD, Down, P
	GetKeyState, Place8, Right, P
	GetKeyState, Place9, Left, P
	if PlaceD = D
	{
		if Place8 = D
			MouseMove, 6, 6, 0, R
		else if Place9 = D
			MouseMove, -6, 6, 0, R
		else
			MouseMove, 0, 6, 0, R
	}
	else if PlaceD = U
		Break
}
Return

Shift:
MouseClick, Left,,,,, D
KeyWait, Shift
MouseClick, Left,,,,, U
Return

Enter:
MouseClick, Right,,,,, D
KeyWait, Enter
MouseClick, Right,,,,, U
Return