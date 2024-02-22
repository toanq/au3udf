#RequireAdmin

#include <AutoItConstants.au3>
#include <Misc.au3>
#include <WinAPI.au3>

Global $pStub_KeyProc = DllCallbackRegister("_KeyProc", "int", "int;ptr;ptr")

MsgBox(0, "", 'Now we block all keyboard keys except a "F3"')

Global $hHook = _WinAPI_SetWindowsHookEx($WH_KEYBOARD_LL, DllCallbackGetPtr($pStub_KeyProc), _WinAPI_GetModuleHandle(0), 0)

While 1
    Sleep(100)
WEnd

Func _MyFunc()
    MsgBox(0, "MyFunc", "Function called by F3")
EndFunc

Func _KeyProc($nCode, $wParam, $lParam)
    If $nCode < 0 Then Return _WinAPI_CallNextHookEx($hHook, $nCode, $wParam, $lParam)

    Local $KBDLLHOOKSTRUCT = DllStructCreate("dword vkCode;dword scanCode;dword flags;dword time;ptr dwExtraInfo", $lParam)
    Local $vkCode = DllStructGetData($KBDLLHOOKSTRUCT, "vkCode")

    If $vkCode <> 0x72 Then
		ConsoleWrite($vkCode&@CRLF)
		Return 1
	EndIf

	_MyFunc()

    _WinAPI_CallNextHookEx($hHook, $nCode, $wParam, $lParam)
EndFunc

Func OnAutoitExit()
    DllCallbackFree($pStub_KeyProc)
    _WinAPI_UnhookWindowsHookEx($hHook)
EndFunc