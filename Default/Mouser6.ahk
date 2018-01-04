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
	Hotkey, RShift, RShift, On
	Hotkey, RControl, RControl, On
}
else
{
	Hotkey, Up, Up, Off
	Hotkey, Right, Right, Off
	Hotkey, Left, Left, Off
	Hotkey, Down, Down, Off
	Hotkey, RShift, RShift, Off
	Hotkey, RControl, RControl, Off
}
Return

~LControl & UP::
Loop
{
	GetKeyState, PlaceUSX, Up, P
	GetKeyState, Place01X, Right, P
	GetKeyState, Place02X, Left, P
	if PlaceUSX = D
	{
		if Place01X = D
			MouseMove, 45, -45, 0, R
		else if Place02X = D
			MouseMove, -45, -45, 0, R
		else
			MouseMove, 0, -45, 0, R
	}
	else if PlaceUSX = U
		Break
}
Return

~LShift & UP::
Loop
{
	GetKeyState, PlaceUS, Up, P
	GetKeyState, Place01, Right, P
	GetKeyState, Place02, Left, P
	if PlaceUS = D
	{
		if Place01 = D
			MouseMove, 3, -3, 0, R
		else if Place02 = D
			MouseMove, -3, -3, 0, R
		else
			MouseMove, 0, -3, 0, R
	}
	else if PlaceUS = U
		Break
}
Return

Up:
Loop
{
	GetKeyState, PlaceU, Up, P
	GetKeyState, Place2, Right, P
	GetKeyState, Place3, Left, P
	if PlaceU = D
	{
		if Place2 = D
			MouseMove, 20, -20, 0, R
		else if Place3 = D
			MouseMove, -20, -20, 0, R
		else
			MouseMove, 0, -20, 0, R
	}
	else if PlaceU = U
		Break
}
Return

~LControl & Right::
Loop
{
	GetKeyState, PlaceRSX, Right, P
	GetKeyState, Place04X, Up, P
	GetKeyState, Place05X, Down, P
	if PlaceRSX = D
	{
		if Place04X = D
			MouseMove, 45, -45, 0, R
		else if Place05X = D
			MouseMove, 45, 45, 0, R
		else
			MouseMove, 45, 0, 0, R
	}
	else if PlaceRSX = U
		Break
}
Return

~LShift & Right::
Loop
{
	GetKeyState, PlaceRS, Right, P
	GetKeyState, Place04, Up, P
	GetKeyState, Place05, Down, P
	if PlaceRS = D
	{
		if Place04 = D
			MouseMove, 3, -3, 0, R
		else if Place05 = D
			MouseMove, 3, 3, 0, R
		else
			MouseMove, 3, 0, 0, R
	}
	else if PlaceRS = U
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
			MouseMove, 20, -20, 0, R
		else if Place5 = D
			MouseMove, 20, 20, 0, R
		else
			MouseMove, 20, 0, 0, R
	}
	else if PlaceR = U
		Break
}
Return

~LControl & Left::
Loop
{
	GetKeyState, PlaceLSX, Left, P
	GetKeyState, Place06X, Up, P
	GetKeyState, Place07X, Down, P
	if PlaceLSX = D
	{
		if Place06X = D
			MouseMove, -45, -45, 0, R
		else if Place07X = D
			MouseMove, -45, 45, 0, R
		else
			MouseMove, -45, 0, 0, R
	}
	else if PlaceLSX = U
		Break
}

~LShift & Left::
Loop
{
	GetKeyState, PlaceLS, Left, P
	GetKeyState, Place06, Up, P
	GetKeyState, Place07, Down, P
	if PlaceLS = D
	{
		if Place06 = D
			MouseMove, -3, -3, 0, R
		else if Place07 = D
			MouseMove, -3, 3, 0, R
		else
			MouseMove, -3, 0, 0, R
	}
	else if PlaceLS = U
		Break
}

Left:
Loop
{
	GetKeyState, PlaceL, Left, P
	GetKeyState, Place6, Up, P
	GetKeyState, Place7, Down, P
	if PlaceL = D
	{
		if Place6 = D
			MouseMove, -20, -20, 0, R
		else if Place7 = D
			MouseMove, -20, 20, 0, R
		else
			MouseMove, -20, 0, 0, R
	}
	else if PlaceL = U
		Break
}
Return

~LControl & Down::
Loop
{
	GetKeyState, PlaceDSX, Down, P
	GetKeyState, Place08X, Right, P
	GetKeyState, Place09X, Left, P
	if PlaceDSX = D
	{
		if Place08X = D
			MouseMove, 45, 45, 0, R
		else if Place09X = D
			MouseMove, -45, 45, 0, R
		else
			MouseMove, 0, 45, 0, R
	}
	else if PlaceDSX = U
		Break
}
Return

~LShift & Down::
Loop
{
	GetKeyState, PlaceDS, Down, P
	GetKeyState, Place08, Right, P
	GetKeyState, Place09, Left, P
	if PlaceDS = D
	{
		if Place08 = D
			MouseMove, 3, 3, 0, R
		else if Place09 = D
			MouseMove, -3, 3, 0, R
		else
			MouseMove, 0, 3, 0, R
	}
	else if PlaceDS = U
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
			MouseMove, 20, 20, 0, R
		else if Place9 = D
			MouseMove, -20, 20, 0, R
		else
			MouseMove, 0, 20, 0, R
	}
	else if PlaceD = U
		Break
}
Return

RShift:
MouseClick, Left,,,,, D
KeyWait, RShift
MouseClick, Left,,,,, U
Send, {Blind}{Shift Up}
Return

RControl:
MouseClick, Right,,,,, D
KeyWait, RControl
MouseClick, Right,,,,, U
Return