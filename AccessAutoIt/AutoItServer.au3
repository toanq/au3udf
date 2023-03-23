#include "AutoitObject.au3"

Opt("MustDeclareVars", 1)

; Error monitoring
Global $oError = ObjEvent("AutoIt.Error", "_ErrFunc")
Func _ErrFunc()
;~  ConsoleWrite("! COM Error !  Number: 0x" & Hex($oError.number, 8) & "   ScriptLine: " & $oError.scriptline & " - " & $oError.windescription & @CRLF)
    Return
EndFunc   ;==>_ErrFunc

; Initialize AutoItObject
_AutoItObject_StartUp()

; Create object
Global $oObject = _SomeObject()

; Register it globaly with some identifier
Global $sMyCLSID = "AutoIt.Application"
Global $hObj = _AutoItObject_RegisterObject($oObject, $sMyCLSID)

While 1
    Sleep(100)
WEnd

_QuitObject(0)

; Define object
Func _SomeObject()
    Local $oClassObject = _AutoItObject_Class()
    $oClassObject.AddMethod("Call", "_Obj_CallAny")
    $oClassObject.AddMethod("Quit", "_QuitObject")
    Return $oClassObject.Object
EndFunc   ;==>_SomeObject

Func _Obj_CallAny($oSelf, $sFunc, _
        $vParam1 = 0, $vParam2 = 0, $vParam3 = 0, $vParam4 = 0, $vParam5 = 0, $vParam6 = 0, _
        $vParam7 = 0, $vParam8 = 0, $vParam9 = 0, $vParam10 = 0, $vParam11 = 0, $vParam12 = 0, _
        $vParam13 = 0, $vParam14 = 0, $vParam15 = 0, $vParam16 = 0, $vParam17 = 0, $vParam18 = 0, _
        $vParam19 = 0, $vParam20 = 0, $vParam21 = 0, $vParam22 = 0, $vParam23 = 0, $vParam24 = 0, _
        $vParam25 = 0, $vParam26 = 0, $vParam27 = 0, $vParam28 = 0, $vParam29 = 0, $vParam30 = 0)
    Local $sExec = $sFunc & "("
    If @NumParams = 3 And IsArray($vParam1) And UBound($vParam1, 0) = 1 Then
        For $n = 0 To UBound($vParam1) - 1
            $sExec &= "$vParam1[" & $n & "],"
        Next
    Else
        If @NumParams = 2 Then
            $sExec &= ")"
        Else
            For $n = 1 To @NumParams - 2
                $sExec &= "$vParam" & $n & ","
            Next
        EndIf
    EndIf
    Local $vRet = Execute(StringTrimRight($sExec, 1) & ")")
    If IsPtr($vRet) Then Return Number($vRet)
    For $i = 0 To UBound($vRet) - 1
        If IsPtr($vRet[$i]) Then $vRet[$i] = Number($vRet[$i])
    Next
    Return $vRet
EndFunc   ;==>_Obj_CallAny

Func _QuitObject($oSelf)
    MsgBox(262144 + 64, "Server says...", "bye bye", 3)
    _AutoItObject_UnregisterObject($hObj)
    $oObject = 0
    Exit
EndFunc   ;==>_QuitObject
