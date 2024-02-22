#include <WinAPI.au3>

$hWnd = WinGetHandle("[CLASS:Chrome_WidgetWin_1]")

$hChrome = _WinAPI_GetParent($hWnd)

WinSetTrans($hWnd, "", 160)