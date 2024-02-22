#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <GuiEdit.au3>
#include <File.au3>
#RequireAdmin

Opt("ExpandVarStrings", 1)
Opt("GUIOnEventMode", 1)

$IDENTIFIER = "9f79269d-7d1c-4125-bdc3-98bb1f43a3c3"
$SOFTNAME = "Carcise"
ConfigWrite("HiThere", "REG_SZ", "22222")

$hGUI = GUICreate("", 400, 300, -1, -1 , -1, $WS_EX_ACCEPTFILES)
$hFileInfo = GUICtrlCreateEdit("Drop files here", 10, 10, 380, 100, BitOR($GUI_SS_DEFAULT_EDIT, $WS_BORDER, $ES_READONLY))
_GUICtrlEdit_SetPadding ( $hFileInfo, 2, 2)
$sCurrentSheet = ""

GUISetOnEvent($GUI_EVENT_CLOSE, OnClose)
GUISetOnEvent($GUI_EVENT_DROPPED, OnChooseFile)
GUICtrlSetOnEvent($hFileInfo, OnChooseFile)
GUICtrlSetState($hFileInfo, $GUI_DROPACCEPTED)
GUISetState(@SW_SHOW)

While 1
	Sleep(100)
WEnd

Func OnClose()
	Exit
EndFunc

Func OnChooseFile()
	Local $sFile

	Switch @GUI_CtrlId
		Case $hFileInfo
			$sFile = FileOpenDialog("Select file", "", "All (*.*)", 11, "", $hGUI)
		Case $GUI_EVENT_DROPPED
			$sFile = @GUI_DragFile
	EndSwitch

	If StringLen($sFile) = 0 Then Return

	ChooseFile($sFile)
EndFunc

Func ChooseFile($sPath)
	$sCurrentSheet = $sPath
	Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
	Local $aPath = _PathSplit($sPath, $sDrive, $sDir, $sFileName, $sExtension)

	Local $sFullFileName = "$sFileName$$sExtension$"
	Local $sFileInfo = _
		"File: $sFullFileName$"&@CRLF& _
		"Size: "&FileGetSize($sPath)&" bytes"&@CRLF
	GUICtrlSetData($hFileInfo, $sFileInfo)
EndFunc

Func ConfigWrite($sKey, $sType, $sValue)
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\"&$SOFTNAME&"\"&$IDENTIFIER, $sKey, $sType, $sValue)
EndFunc