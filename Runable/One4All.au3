#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

Opt("ExpandVarStrings", 1)

Global Const $ROW_HEIGHT = 32
Global Const $sNPP = "@ProgramFilesDir@\Notepad++\notepad++.exe"

Global $aFunction[0]
Global $aControl[0]
Global $iFeature = 0
Global $aHiddenFn[0]
Global $aHiddenEv[0]
Global $iHiddenFn = 0

AddFeature("NotepadPlusPlus")
AddFeature("Close", -3)

$hMainForm = GUICreate("", 259, $ROW_HEIGHT*$iFeature+16)

InitControls()

GUISetState(@SW_SHOW)

While 1
	HandleEvent()
WEnd

Func GetWorkingDir($sExePath)
	Local $aResult = StringRegExp($sExePath, "^.*\\", 3)
	Return $aResult[0]
EndFunc

Func InitControls()
	For $ii = 0 To $iFeature - 1
		If $aControl[$ii] = 0 Then
			$aControl[$ii] = GUICtrlCreateButton($aFunction[$ii], 24, $ROW_HEIGHT*$ii+8, 211, 25)
		Else
			$aControl[$ii] &=","&GUICtrlCreateButton($aFunction[$ii], 24, $ROW_HEIGHT*$ii+8, 211, 25)
		EndIf
	Next
EndFunc

Func AddFeature($fnFunction, $iCustomControl = 0xDEADBEEF)
	Local $iLastIndex = UBound($aFunction)
	$iFeature = $iLastIndex + 1

	ReDim $aControl[$iFeature]
	ReDim $aFunction[$iFeature]

	$aFunction[$iLastIndex] = $fnFunction
	ConsoleWrite("Add feature $fnFunction$@CRLF@")

	If $iCustomControl <> 0xDEADBEEF Then
		$iHiddenFn += 1
		ReDim $aHiddenFn[$iHiddenFn]
		ReDim $aHiddenEv[$iHiddenFn]
		$aHiddenFn[$iHiddenFn-1] = $fnFunction
		$aHiddenEv[$iHiddenFn-1] = $iCustomControl
	EndIf
EndFunc

Func HandleEvent()
	Local $iMsg = GUIGetMsg()
	Local $iControl = UBound($aControl) - 1

	For $ii = 0 To $iControl
		If $aControl[$ii] = $iMsg Then
			Call($aFunction[$ii])
			ExitLoop
		EndIf
	Next

	For $ii = 0 To $iHiddenFn - 1
		If $aHiddenEv[$ii] = $iMsg Then
			Call($aHiddenFn[$ii])
			ExitLoop
		EndIf
	Next
EndFunc

Func Close()
	Exit
EndFunc

Func NotepadPlusPlus()
	Run($sNPP)
EndFunc