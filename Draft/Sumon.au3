#include <Array.au3>
#include <WinAPI.au3>
#include <WinAPIGdi.au3>
Global Const $SM_CMONITORS = 80

If _WinAPI_GetSystemMetrics($SM_CMONITORS) <= 1 Then Exit

$aMonitor = _WinAPI_EnumDisplayMonitors()
$iPosX = 0

For $ii = 1 To $aMonitor[0][0]
	$aMonitorInfo = _WinAPI_GetMonitorInfo($aMonitor[$ii][0])
	If $aMonitorInfo[2] = 0 Then
		$aPos = _WinAPI_GetPosFromRect($aMonitor[$ii][1])
		$iPosX = $aPos[0]
	EndIf
Next

While 1
	Sleep(100)
	$hHandle = WinGetHandle(StringShift("Kd`ftdneKdfdmcr'SL(Bkhdms", 1))
	If Not WinExists($hHandle) Then ContinueLoop

	$aPos = WinGetPos($hHandle)
	If @error Then ContinueLoop

	If $aPos[0] = $iPosX Then ContinueLoop
	If $aPos[2] < 100 Or $aPos[3] < 100 Then ContinueLoop

	Sleep(1000)
	WinMove($hHandle, "", $iPosX, 0)
WEnd

Func CharShift($sChar, $iShift = 1)
	Local $iResult = Mod(Asc($sChar) - 31 + $iShift, 94)

	$iResult = ($iResult < 0) ? $iResult + 94 : $iResult

	Return Chr($iResult + 31)
EndFunc

Func StringShift($sInput, $iShift = 1)
	Local $ii, $sChar
	For $ii = 1 To StringLen($sInput)
		$sChar = StringMid($sInput, $ii, 1)
		$sInput = StringLeft($sInput, $ii - 1) & CharShift($sChar, $iShift) & StringRight($sInput, StringLen($sInput) - $ii)
	Next

	Return $sInput
EndFunc