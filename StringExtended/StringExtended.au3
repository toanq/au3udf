#include-once
#include <String.au3>

Func _StringAppend(ByRef $sSource, $sValue)
	$sSource &= $sValue
EndFunc

Func _StringAppendLine(ByRef $sSource, $sValue)
	$sSource &= "@CRLF@$sValue$"
EndFunc