'Start AutoIt server script first

Set oAutoIt = GetObject("AutoIt.Application") ' the magic

WS_OVERLAPPEDWINDOW = &H00CF0000

hGui = oAutoIt.Call("GUICreate", "VBS AutoIt GUI test", -1, -1, -1, -1, WS_OVERLAPPEDWINDOW)
hButton = oAutoIt.Call("GUICtrlCreateButton", "Click", 100, 100, 100, 30)
hButton2 = oAutoIt.Call("GUICtrlCreateButton", "Click me too", 100, 300, 100, 30)

oAutoIt.Call "WinSetOnTop", "VBS AutoIt GUI test", "", 1

AW_FADE_IN  = &H00080000
oAutoIt.Call "DllCall", "user32.dll", "bool", "AnimateWindow", "hwnd", hGui, "dword", 1000, "dword", AW_FADE_IN

oAutoIt.Call "GUISetState"

Do
    Select Case oAutoIt.Call("GUIGetMsg")
        Case -3
            Exit Do
        Case hButton
            oAutoIt.Call "MsgBox", 262144+32+3, "Title", "Bzzz bzz bzzzz", 0, hGUI
        Case hButton2
			oAutoIt.Call "Beep", 500, 700
	End Select
	Wscript.Sleep(10)
Loop

oAutoIt.Call "GUIDelete"

If oAutoIt.Call("MsgBox", 4 + 48 + 262144, "?", "Kill server?") = 6 Then oAutoIt.Quit