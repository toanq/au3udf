#include <Date.au3>
#include <Excel.au3>

Opt("ExpandVarStrings", 1)
Local $sWorkbook = _Path_FullPath("ow-9999-week-99.xlsx")
Local $iCurrentWeekNumber = _WeekNumberISO()
Local $sTarget = _Path_FullPath($sCurrentWeekSheetName)
Local $sCurrentWeekSheetName = "ow-@YEAR@-week-$iCurrentWeekNumber$.xlsx"

If FileExists($sTarget) Then
	MsgBox(0,'',"Report for current week are existing!!!", 1.5)
	Exit
EndIf

Local $aWeek = _Date_CurrentWeekDays()
Local $sCurrentSheetName = ""

Local $dateStart = StringRegExpReplace($aWeek[1], "(\d{4}).(\d{2}).(\d{2})(.*)", "${3}/${2}/${1}")
Local $dateEnd = StringRegExpReplace($aWeek[5], "(\d{4}).(\d{2}).(\d{2})(.*)", "${3}/${2}/${1}")
Local $sCurrentWeek = "Week #$iCurrentWeekNumber$ ($dateStart$ - $dateEnd$)"

Local $oExcel = _Excel_Open()
Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)
Local $aSheet = _Excel_SheetList($oWorkbook)

For $ix = 1 To 5
	$sCurrentSheetName = _Date_DateToSheetName($aWeek[$ix])
	$aSheet[$ix-1][1].Name = $sCurrentSheetName
Next

Local $oOverallSheet = $aSheet[5][1]
_Excel_RangeWrite($oWorkbook, $oOverallSheet, "[Delivery][Group4][OpenWindows] Weekly Report - $sCurrentWeek$", "A1")
_Excel_RangeWrite($oWorkbook, $oOverallSheet, "[OW] Toan Nguyen Weekly Report - $sCurrentWeek$", "A2")

_Excel_BookSaveAs($oWorkbook, $sTarget)

_Excel_Close($oExcel)

Func _Path_FullPath($sFileName, $sPath = @ScriptDir)
	Local $_sFileName = $sFileName
	Local $_sPath = $sPath

	$_sFileName = StringRegExpReplace($sFileName, "^(?:\/|\\)+", "")
	$_sPath = StringRegExpReplace($sPath, "(?:\/|\\)+$", "")

	Return "$_sPath$\$_sFileName$"
EndFunc

Func _Date_CurrentWeekDays()
	; Weeks start at Sunday
	Local $dateSunday = _DateAdd('d', 1-@WDAY, _NowCalc())
	Local $aResult[7]

	For $ii = 0 To 6
		$aResult[$ii] = _DateAdd('d', $ii, $dateSunday)
	Next

	Return $aResult
EndFunc

Func _Date_DateToSheetName($dateValue)
	Local $aDate, $aTime
	_DateTimeSplit($dateValue, $aDate, $aTime)

	Return StringFormat("%02d%02d", $aDate[2], $aDate[3])
EndFunc