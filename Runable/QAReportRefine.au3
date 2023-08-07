#include <Array.au3>

$sData = ClipGet()
$aData = StringSplit($sData, @CR, 1)
$sResult = ""
$iTotal = -1
$iCompletion = 0

For $sItem In $aData
	If Not StringRegExp($sItem, "TOTAL") Then ContinueLoop

	$sCount = StringFormat("%02d", StringRegExp($sItem, ": (\d+)", 3)[0])
	$sPercent = ""
	$iPercent = 0
	If $iTotal < 0 Then
		$iTotal = $sCount
	Else
		$iPercent = $sCount/$iTotal*100
		$sPercent = " - "&StringFormat("%05.2f", $iPercent)&"%"
	EndIf

	$sText = StringLeft(StringRegExp($sItem, "(TOTAL.+?):", 3)[0]&"                      ", 20)
	If StringRegExp($sText, "(?:PASSED|UNTESTABLE|BLOCKED)") Then
		$iCompletion += $iPercent
	EndIf

	$sResult &= $sText&": "&$sCount&$sPercent&@CRLF
Next

$sResult &= "================================="&@CRLF
$sResult &= "COMPLETION          : "&StringFormat("%05.2f", $iCompletion)&"%"&@CRLF

ConsoleWrite($sResult)
ClipPut($sResult)