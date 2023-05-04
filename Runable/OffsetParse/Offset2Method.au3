FileDelete("dump.txt")
Global $aOffset = [0x5A38C58 ,0x5A38D54 ,0x5A2D5AC ,0x3544F34 ,0x354FBDC ,0x5A21E00 ,0x5A26090 ,0x5A26804 ,0x5A26C38 ,0x5A80A54 ,0x5BBA9E4 ,0x382B14C ,0x58722E8,0x5952CB8 ,0x5952F78 ,0x59533A4 ,0x5953490 ,0x595357C ,0x59545EC ,0x595530C ,0x5738A48 ,0x5739254 ,0x5738EF4 ,0x5738FA8 ,0x573A46C ,0x573A534 ,0x573A590 ,0x573AF68 ,0x573B00C ,0x573B08C ,0x5FED248 ,0x5070E64 ,0x5070E74 ,0x5070EE0 ,0x5070F64 ,0x5070FD8 ,0x5071090 ,0x50713BC ,0x9FAE4]
Global $iOffset = UBound($aOffset)
Global $aFound[$iOffset]

$sResult = ""
For $ii In $aOffset
	$sResult &= "0x"&StringRegExpReplace(Hex($ii),'^0*','')&"|"
Next
$sPattern = "(?:"&StringTrimRight($sResult,1)&")"


$aData = FileReadToArray("dump.cs")
$iData = UBound($aData)-1

For $ii=0 To $iData
	If StringRegExp($aData[$ii], $sPattern) Then
		$iCurrent = StringRegExp($aData[$ii], "(0x[0-9A-Fa-f]{2,8})", 3)[0]
		FileWrite("dump.txt",$iCurrent&" => "&$aData[$ii+1]&@CRLF)
		For $jj = 0 To $iOffset-1
			If $aOffset[$jj] = $iCurrent Then $aFound[$jj] += 1
		Next
	EndIf
Next

FileWrite("dump.txt","Not found: ")
For $jj = 0 To $iOffset-1
	If $aFound[$jj] = 0 Then FileWrite("dump.txt","0x"&StringRegExpReplace(Hex($aOffset[$jj]),'^0*','')&", ")
Next
FileWrite("dump.txt",@CRLF)