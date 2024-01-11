#include <Misc.au3>

Local $dll = DllOpen("main.dll") ; open dll

$hTime = TimerInit()
$iSum = 0
For $i = 1 To 1000000
	$iSum = Add($iSum, $i)
Next
ConsoleWrite("Add: "&TimerDiff($hTime)&@CRLF)

$hTime = TimerInit()
ExternalAdd(100000000000)
ConsoleWrite("ExternalAdd: "&TimerDiff($hTime)&@CRLF)


Func Add($a, $b)
	Return $a+$b
EndFunc

Func ExternalAdd($iMax)
	Local $v = DllCall($dll, 'int:cdecl', 'externalAdd', 'int', $iMax)
	Return $v[0]
EndFunc