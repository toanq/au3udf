#include <Array.au3>
Opt("ExpandVarStrings", 1)

$sData = ClipGet()
$aData = StringSplit($sData, @CRLF, 3)
Global $aTask[0]
Global $aResult[0]
Local $iLastIndex = 0
$iCurrentTaskIndex = 0
For $i = 0 To UBound($aData)-1
	If (StringRegExp($aData[$i], "^(?:Bug|Product Backlog Item)")) Then
			$iCurrentTaskIndex = $i
			$iLastIndex = UBound($aTask)
			ReDim $aTask[$iLastIndex+1]
			$aTask[$iLastIndex] = $aData[$i]
	ElseIf (StringRegExp($aData[$i], "^Task")) Then
		MsgBox(0,'ERROR', 'Data was not clean 1@CRLF@'&$aData[$i])
		Exit
	Else
		$aTask[$iLastIndex] &= "\n"&$aData[$i]
		$aData[$iCurrentTaskIndex] &= "\n"&$aData[$i]
		$aData[$i] = "[$iCurrentTaskIndex$] "&$aData[$i]
	EndIf
Next

;_ArrayDisplay($aTask)

For $i = UBound($aTask)-1 To 1 Step -1
	If StringRegExp($aTask[$i], "^###") Then
			;ConsoleWrite("Skip@CRLF@")
			ContinueLoop
	EndIf

	If Not StringRegExp($aTask[$i], "^(?:Bug|Product Backlog Item)") Then
		MsgBox(0,'ERROR', 'Data was not clean 2@CRLF@'&$aTask[$i])
		Exit
	EndIf

	$iCurrentTaskId = GetTaskId($aTask[$i])
	;ConsoleWrite("Current task id: $iCurrentTaskId$@CRLF@")
	For $x = $i-1 To 0 Step -1
		If $iCurrentTaskId = GetTaskId($aTask[$x]) Then
			$aTask[$x] = "### "&$aTask[$x]
			;ConsoleWrite("!>match@CRLF@")
		EndIf
	Next
Next

;_ArrayDisplay($aTask)

For $i = 0 To UBound($aTask)-1
	If Not StringRegExp($aTask[$i], "^(?:Bug|Product Backlog Item)") Then ContinueLoop

	$iLastIndex = UBound($aResult)
	ReDim $aResult[$iLastIndex+1]
	$aResult[$iLastIndex] = $aTask[$i]
	ConsoleWrite($aTask[$i]&@CRLF)
Next

_ArrayDisplay($aResult)

Func GetTaskId($sTaskName)
	Local $_aResult = StringRegExp($sTaskName, "(?:Bug|Product Backlog Item) (\d+)", 3)

	If @error Then
		MsgBox(0,'',"Cannot get Bug/PBI id from:@CRLF@$sTaskName$")
		Exit
	EndIf

	Return $_aResult[0]
EndFunc