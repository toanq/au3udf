#include <GUIConstantsEx.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>
#include <GuiTreeView.au3>
#include <Array.au3>

Local Const $_ARRAYCONSTANT_GUI_DOCKBOTTOM = 64
Local Const $_ARRAYCONSTANT_GUI_DOCKBORDERS = 102
Local Const $_ARRAYCONSTANT_GUI_DOCKHEIGHT = 512
Local Const $_ARRAYCONSTANT_GUI_DOCKLEFT = 2
Local Const $_ARRAYCONSTANT_GUI_DOCKRIGHT = 4
Local Const $_ARRAYCONSTANT_GUI_DOCKHCENTER = 8
Local Const $_ARRAYCONSTANT_GUI_EVENT_CLOSE = -3
Local Const $_ARRAYCONSTANT_GUI_EVENT_ARRAY = 1
Local Const $_ARRAYCONSTANT_GUI_FOCUS = 256
Local Const $_ARRAYCONSTANT_SS_CENTER = 0x1
Local Const $_ARRAYCONSTANT_SS_CENTERIMAGE = 0x0200
Local Const $_ARRAYCONSTANT_LVM_GETITEMRECT = (0x1000 + 14)
Local Const $_ARRAYCONSTANT_LVM_GETITEMSTATE = (0x1000 + 44)
Local Const $_ARRAYCONSTANT_LVM_GETSELECTEDCOUNT = (0x1000 + 50)
Local Const $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE = (0x1000 + 54)
Local Const $_ARRAYCONSTANT_LVS_EX_GRIDLINES = 0x1
Local Const $_ARRAYCONSTANT_LVIS_SELECTED = 0x0002
Local Const $_ARRAYCONSTANT_LVS_SHOWSELALWAYS = 0x8
Local Const $_ARRAYCONSTANT_LVS_OWNERDATA = 0x1000
Local Const $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT = 0x20
Local Const $_ARRAYCONSTANT_LVS_EX_DOUBLEBUFFER = 0x00010000 ; Paints via double-buffering, which reduces flicker
Local Const $_ARRAYCONSTANT_WS_EX_CLIENTEDGE = 0x0200
Local Const $_ARRAYCONSTANT_WS_MAXIMIZEBOX = 0x00010000
Local Const $_ARRAYCONSTANT_WS_MINIMIZEBOX = 0x00020000
Local Const $_ARRAYCONSTANT_WS_SIZEBOX = 0x00040000


Func TreeDisplay($oData, $sTitle = "")
	Local $iViewWidth = 400, $iViewHeigh = 300

	Local $hMain = GUICreate("", $iViewWidth, $iViewHeigh, Default, Default, BitOR($_ARRAYCONSTANT_WS_SIZEBOX, $_ARRAYCONSTANT_WS_MINIMIZEBOX, $_ARRAYCONSTANT_WS_MAXIMIZEBOX))
	Local $hTreeView = GUICtrlCreateTreeView(0, 0, $iViewWidth, $iViewHeigh, -1,  BitOR($GUI_SS_DEFAULT_TREEVIEW, $WS_HSCROLL))
	Local $hRootItem = GUICtrlCreateTreeViewItem("[[root]]", $hTreeView)
	GUICtrlSendMsg($hTreeView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_GRIDLINES, $_ARRAYCONSTANT_LVS_EX_GRIDLINES)
	GUICtrlSendMsg($hTreeView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT, $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT)
	GUICtrlSendMsg($hTreeView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_DOUBLEBUFFER, $_ARRAYCONSTANT_LVS_EX_DOUBLEBUFFER)
	GUICtrlSendMsg($hTreeView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_WS_EX_CLIENTEDGE, $_ARRAYCONSTANT_WS_EX_CLIENTEDGE)

	GUICtrlSetResizing($hMain, $_ARRAYCONSTANT_GUI_DOCKBORDERS)
	ParseChild($oData, $hRootItem)

	GUICtrlSetState($hRootItem, $GUI_EXPAND)
	GUISetState(@SW_SHOW)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUIDelete($hMain)
				ExitLoop
		EndSwitch
	WEnd
EndFunc

Func ParseChild($oInput, $hParentNode)
  ; Convert array to map
  If IsArray($oInput) Then
    Local $mTemp[], $itr

    For $itr=0 To UBound($oInput)-1
      $mTemp["["&$itr&"]"] = $oInput[$itr]
    Next

    $oInput = $mTemp
  EndIf

	$aKeys = MapKeys($oInput)
  If UBound($aKeys) = 0 Then
    GUICtrlSetData($hParentNode, GUICtrlRead($hParentNode, 1)&": Empty")
  EndIf

	For $key In $aKeys
    Local $hChild
    Local $oRecord = $oInput[$key]
		Local $sNodeLable = $key&" ("&VarGetType($oRecord)&")"

		If (IsMap($oRecord) Or IsArray($oRecord)) Then
      $hChild = GUICtrlCreateTreeViewItem($sNodeLable, $hParentNode)
      ParseChild($oRecord, $hChild)
    Else
			$sNodeLable &= ": "&$oRecord
      $hChild = GUICtrlCreateTreeViewItem($sNodeLable, $hParentNode)
		EndIf

		GUICtrlSetState($hChild, $GUI_EXPAND)
	Next
EndFunc
