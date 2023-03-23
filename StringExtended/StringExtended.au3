#include-once
#include <String.au3>
#include <MsgBoxConstants.au3>

Opt("ExpandVarStrings", 1)

Func _StringAppend(ByRef $sSource, $sValue)
	$sSource &= $sValue
	Return $sSource
EndFunc

Func _StringAppendLine(ByRef $sSource, $sValue = "")
	$sSource &= "@CRLF@$sValue$"
	Return $sSource
EndFunc

Func _StringRegExpFirstOrError($sTest, $sPattern, $bExitOnError = True)
	Local $aResult = StringRegExp($sTest, $sPattern, 3)
	Local $sErrorMessage = ''

	If @error Then
		_StringAppendLine($sErrorMessage, "!> ERROR")
		_StringAppendLine($sErrorMessage, "!> Function: _StringRegExpFirstOrError")
		_StringAppendLine($sErrorMessage, "!> $$sTest: $sTest$")
		_StringAppendLine($sErrorMessage, "!> $$sPattern: $sPattern$")
		_StringAppendLine($sErrorMessage)

		ConsoleWrite($sErrorMessage)
		If $bExitOnError Then Exit
	EndIf

	Return SetError(@error, 0, $aResult[0])
EndFunc