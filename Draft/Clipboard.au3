;https://learn.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format

#include <ClipBoard.au3>
; Open the clipboard
If Not _ClipBoard_Open (0) Then msgbox (0,'Error',"_ClipBoard_Open failed")
$iFormat = _ClipBoard_RegisterFormat ("HTML Format")

If $iFormat = 0 Then msgbox(0,'',"_ClipBoard_RegisterFormat failed")
;msgbox(0,'','$iformat is ' & $iFormat)
$i = _ClipBoard_GetData($iFormat)
msgbox(0,'$i is',BinaryToString($i))
$m_sDescription = _
"Version:1.0" & @CrLf & _
"StartHTML:aaaaaaaaaa" & @CrLf & _
"EndHTML:bbbbbbbbbb" & @CrLf & _
"StartFragment:cccccccccc" & @CrLf & _
"EndFragment:dddddddddd" & @CrLf
$sStart = "<HTML><BODY><FONT FACE=Arial SIZE=1 COLOR=BLUE>"
$sFrag = '<B>This is bold</B> and <I>this is italic.</I> And this is a <a href="http://www.yahoo.com">hyperlink</a>'
$sEnd = "</FONT></BODY></HTML>"
;msgbox(0,'test',stringFormat( "%010s",stringLen($m_sDescription)))
$sData = $m_sDescription & $sStart & $sFrag & $sEnd
$sDatax = $sData
$sDatax = stringReplace($sData, "aaaaaaaaaa", stringFormat("%010s",stringLen($m_sDescription)))
$sDatax = stringReplace($sData, "bbbbbbbbbb", stringFormat("%010s",stringLen($sData)))
$sDatax = stringReplace($sData, "cccccccccc", stringFormat("%010s",stringLen($m_sDescription & $sStart)))
$sDatax = stringReplace($sData, "dddddddddd", stringFormat("%010s",stringLen( $m_sDescription & $sStart & $sFrag)))
$sData = stringReplace($sData, "aaaaaaaaaa", stringFormat("%d",stringLen($m_sDescription)))
$sData = stringReplace($sData, "bbbbbbbbbb", stringFormat("%d",stringLen($sData)))
$sData = stringReplace($sData, "cccccccccc", stringFormat("%d",stringLen($m_sDescription & $sStart)))
$sData = stringReplace($sData, "dddddddddd", stringFormat("%d",stringLen( $m_sDescription & $sStart & $sFrag)))
;msgbox(0,'sdata',$sData)
if (my_ClipBoard_SetData($sData,$iformat,1) == 0 ) then msgbox(0,'Error','Error putting data on clipboard')
_ClipBoard_Close()

Func My_ClipBoard_SetData($vData, $iFormat = 1,$force = 0)
Local $tData, $hLock, $hMemory, $iSize
If ($force ==1) Then
$iSize = StringLen($vData) + 1
$hMemory = _MemGlobalAlloc($iSize, $GHND)
If $hMemory = 0 Then Return SetError(-1, 0, 0)
$hLock = _MemGlobalLock($hMemory)
If $hLock = 0 Then Return SetError(-2, 0, 0)
$tData = DllStructCreate("char Text[" & $iSize & "]", $hLock)
DllStructSetData($tData, "Text", $vData)
_MemGlobalUnlock($hMemory)
Else
; Assume all other formats are a pointer or a handle (until users tell me otherwise) :)
$hMemory = $vData
EndIf
If Not _ClipBoard_Open(0) Then Return SetError(-5, 0, 0)
If Not _ClipBoard_Empty() Then Return SetError(-6, 0, 0)
If Not _ClipBoard_SetDataEx($hMemory, $iFormat) Then
_ClipBoard_Close()
Return SetError(-7, 0, 0)
EndIf
_ClipBoard_Close()
Return $hMemory
EndFunc ;==>my_ClipBoard_SetData