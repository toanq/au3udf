#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

Opt("ExpandVarStrings", 1)

Global Const $ROW_HEIGHT = 32
Global Const $sNPP = "@ProgramFilesDir@\Notepad++\notepad++.exe"

Global $aFunction[0]
Global $aControl[0]
Global $iFeature = 0

AddFeature(NotepadPlusPlus)
AddFeature(Close)

$hMainForm = GUICreate("", 259, $ROW_HEIGHT*$iFeature+16)

InitControls()

ReDim $aControl[$iFeature+1]
$aControl[$iFeature] = $GUI_EVENT_CLOSE
$iFeature += 1

GUISetState(@SW_SHOW)

While 1
	HandleEvent()
WEnd

Func GetWorkingDir($sExePath)
	Local $aResult = StringRegExp($sExePath, "^.*\\", 3)
	Return $aResult[0]
EndFunc

Func InitControls()
	ReDim $aControl[$iFeature]

	For $ii = 0 To $iFeature - 1
		$aControl[$ii] = GUICtrlCreateButton(FuncName($aFunction[$ii]), 24, $ROW_HEIGHT*$ii+8, 211, 25)
	Next
EndFunc

Func AddFeature($fnFunction)
	Local $iLastIndex = UBound($aFunction)
	$iFeature = $iLastIndex + 1

	ReDim $aFunction[$iFeature]
	$aFunction[$iLastIndex] = $fnFunction

	ConsoleWrite("Add feature "&FuncName($fnFunction)&@CRLF)
EndFunc

Func HandleEvent()
	Local $iMsg = GUIGetMsg()
	Local $iControl = UBound($aControl) - 1

	For $ii = 0 To $iControl
		If $iMsg = $aControl[$ii] Then
			$aFunction[$ii]()
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