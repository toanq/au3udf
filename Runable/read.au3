#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Program Files (x86)\AutoIt3\Icons\MyAutoIt3_Blue.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include "D:\Workplace\au3udf\_HttpRequest\_HttpRequest.au3"
#include "D:\Workplace\au3udf\ConsoleWriteVi\ConsoleWriteVi.au3"
#include <Date.au3>
#include <Misc.au3>

Opt("ExpandVarStrings", 1)
HotKeySet("!{ESC}", OnExit)
TraySetState(2)

Global Const $HOME_URI = "https://www.nettruyenus.com/"
Global Const $COOKIE_FILE = @TempDir & "/" & @ScriptName & ".bin"
Global Const $DEBUG = False
Global $COOKIE = FileRead($COOKIE_FILE)

Local $iRound = 0
Local $sComic = GetComic("https://www.nettruyenus.com/truyen-tranh/doraemon-57230")
Local $aChapterId = GetComicInfos($sComic)
Local $sCheckAuth = CheckAuth()
Local $sUserToken = GetUserToken($sCheckAuth)
Local $iError = 0

While True
	If StringInStr($sCheckAuth, '"userGuid":null') Then
		Login()
		$sCheckAuth = CheckAuth()
		$sUserToken = GetUserToken($sCheckAuth)
	EndIf

	Local $iProgress = GetProgress()
	TraySetToolTip("Current progress: "&$iProgress&"/100")
	For $iml = 1 To $aChapterId[0][0]
		$sToken = ChapterLoadedToken(ChapterLoaded($aChapterId[$iml][0], $aChapterId[$iml][1], $sUserToken))
		Read($aChapterId[$iml][0], (@error) ? "" : $sToken)
		LogWrite(ProgressFormat($iml, $aChapterId[0][0], $iRound))

		If $iError > 10 Then
			$iError = 0
			ExitLoop
		EndIf
	Next
	$iRound += 1
WEnd

#Region Utils
Func StringItp($sValue)
	Return StringInterpolation($sValue)
EndFunc   ;==>StringItp

Func StringInterpolation($sValue)
	Local $sResult = $sValue
	$sResult = StringRegExpReplace($sResult, "({{([^{]+?)}}{{(%0\d+?[A-Za-z])}})", '"&StringFormat("$3",$2)&"')
	$sResult = StringRegExpReplace($sResult, "({{(.+?)}})", '"&$2&"')
	$sResult = StringReplace($sResult, "\{\{", "{{")
	$sResult = StringReplace($sResult, "\}\}", "}}")
	$sResult = StringReplace($sResult, "\r", @CR)
	$sResult = StringReplace($sResult, "\n", @LF)
	$sResult = '"' & $sResult & '"'
	$sResult = Execute($sResult)

	Return $sResult
EndFunc   ;==>StringInterpolation

Func ErrorWrite($sValue)
	$iError += 1
	LogWrite($sValue, "!>")
EndFunc   ;==>LogWrite

Func DebugWrite($sValue)
	If $DEBUG Then LogWrite($sValue, "->")
EndFunc   ;==>LogWrite

Func InfoWrite($sValue)
	LogWrite($sValue, ">>")
EndFunc

Func LogWrite($sValue, $sPrefix = "+>")
	Local $sLine = $sPrefix &_NowTime(5)& " " & $sValue & @LF
	ConsoleWrite(BinaryToString(StringToBinary($sLine, 4)))
	; FileWrite(@ScriptName & ".log", $sLine)
EndFunc   ;==>LogWrite

Func RemoveLastChar(ByRef $sInput, $iCount = 1)
	$sInput = StringMid($sInput, 1, StringLen($sInput) - $iCount)

	Return $sInput
EndFunc   ;==>RemoveLastChar

Func ProgressFormat($iCurrent, $iTotal, $iRound)
	Local $zeroPadding = Floor(Log($iTotal) / Log(10)) + 1
	Return StringFormat("Reading chap %0$zeroPadding$d/%d, Round: %d, Total: %d", _
		$iCurrent, $iTotal, $iRound, $iCurrent + $iRound*$iTotal)
EndFunc

Func OnExit()
	ProcessClose("AutoIt3.exe")
	Exit
EndFunc   ;==>OnExit
#EndRegion Utils
#Region Response processing
Func GetProgress()
	InfoWrite("GetProgress")
	Local $sData = Dashboard()
	Local $aProgress = StringRegExp($sData, 'progress-bar.+?(\d+?)%<', 3)
	Local $iProgress = (@error) ? Default : $aProgress[0]

	Return $iProgress
EndFunc   ;==>GetProgress

Func GetComicInfos($sComicPage)
	InfoWrite("GetComicInfos")
	Local $igci, $iResult
	Local $aMatch = StringRegExp($sComicPage, 'gOpts.key=[''"](.+?)[''"]', 3)
	Local $sComicToken = (@error) ? "" : $aMatch[0]

	Local $aChapter = StringRegExp($sComicPage, '<a href=".+?" data-id="(\d+?)">.+?<\/a>', 3)
	Local $aName = StringRegExp($sComicPage, '<a href=".+?" data-id="\d+?">(.+?)<\/a>', 3)
	Local $aResult[1][3] = [[0, "Comic token", "Name"]]

	For $igci = UBound($aChapter) - 1 To 0 Step -1
		$iResult = UBound($aResult)
		ReDim $aResult[$iResult+1][3]
		$aResult[$iResult][0] = $aChapter[$igci]
		$aResult[$iResult][1] = $sComicToken
		$aResult[$iResult][2] = StringStripWS($aName[$igci], 3)

		$aResult[0][0] = $iResult
	Next

	Return $aResult
EndFunc   ;==>GetComicInfos

Func GetUserToken($sCheckAuth)
	Local $aMatch = StringRegExp($sCheckAuth, '"token":"(.+?)".+?"readToken":"(.+?)"', 3)

	Return (@error) ? "" : $aMatch[0]
EndFunc

Func ChapterLoadedToken($sChapterLoaded)
	Local $aMatch = StringRegExp($sChapterLoaded, '"token":"(.+?)"', 3)

	Return (@error) ? "" : $aMatch[0]
EndFunc
#EndRegion Response processing
#Region API call
Func GetComic($sComic)
	InfoWrite("GetComic, sComic=$sComic$")
	Local $sResponse = _HttpRequest(2, _
			$sComic, _
			"", _
			$COOKIE, _
			$HOME_URI)

	Return $sResponse
EndFunc   ;==>GetComic

Func Dashboard()
	Local $sResponse = _HttpRequest(2, _
			StringItp("{{$HOME_URI}}/Secure/Dashboard.aspx"), _
			"", _
			$COOKIE, _
			$HOME_URI)

	Return $sResponse
EndFunc   ;==>Dashboard

Func Read($iChapterId, $sChapterToken)
	DebugWrite("Read, iChapterId=$iChapterId$, sChapterToken=$sChapterToken$")
	Local $sResponse = _HttpRequest(2, _
			"https://f.nettruyenus.com/Comic/Services/ComicService.asmx/Read", _
			"chapterId="&$iChapterId&"&token="&$sChapterToken, _
			$COOKIE, _
			$HOME_URI)

	If StringInStr($sResponse, '"success":false') Then ErrorWrite("Read Error, Response: $sResponse$")
	DebugWrite("Read, $sResponse$")

	Return $sResponse
EndFunc   ;==>Read

Func ChapterLoaded($iChapterId, $sComicToken, $sUserToken)
	DebugWrite("ChapterLoaded, iChapterId=$iChapterId$, sComicToken=$sComicToken$, sUserToken=$sUserToken$")
	Local $sResponse = _HttpRequest(2, _
			"https://f.nettruyenus.com/Comic/Services/ComicService.asmx/ChapterLoaded", _
			"chapterId="&$iChapterId&"&comicToken="&$sComicToken&"&userToken="&_URIEncode($sUserToken), _
			$COOKIE, _
			$HOME_URI)

	If StringInStr($sResponse, '"success":false') Then ErrorWrite("ChapterLoaded Error, Response: $sResponse$")
	DebugWrite("ChapterLoaded, $sResponse$")

	Return $sResponse
EndFunc   ;==>ChapterLoaded

Func CheckAuth()
	InfoWrite("CheckAuth")
	Local $sResponse = _HttpRequest(2, _
			"https://f.nettruyenus.com/Comic/Services/ComicService.asmx/CheckAuth", _
			"", _
			$COOKIE, _
			$HOME_URI)

	Return $sResponse
EndFunc   ;==>CheckAuth

Func Login()
	InfoWrite("Login")
	Local $sResponse = _HttpRequest(2, _
			"$HOME_URI$/Secure/Login.aspx?returnurl=%2F", _
			"", _
			$COOKIE, _
			$HOME_URI)

	Local $sData = ""
	$aForm = StringRegExp($sResponse, '<input.+?name=".+?".+?(?:value=".+?")?.+?\/>', 3)
	For $sInput In $aForm
		Local $mRow[]
		Local $aMatch

		$aMatch = StringRegExp($sInput, 'name="(.+?)"', 3)
		$mRow["Key"] = (@error) ? "" : _URIEncode($aMatch[0])

		$aMatch = StringRegExp($sInput, 'value="(|.+?)"', 3)
		$mRow["Value"] = (@error) ? "" : _URIEncode($aMatch[0])
		If StringInStr($mRow["Key"], "UserName", 1) Then $mRow["Value"] = "quoctoan.c2b@gmail.com"
		If StringInStr($mRow["Key"], "Password", 1) Then $mRow["Value"] = "231199"
		If StringInStr($mRow["Key"], "RememberMe", 1) Then $mRow["Value"] = "on"

		$sData &= $mRow["Key"] & "=" & $mRow["Value"] & "&"
	Next
	$sData = RemoveLastChar($sData)

	$sResponse = _HttpRequest(2, _
			"$HOME_URI$/Secure/Login.aspx", _
			$sData, _
			$COOKIE, _
			$HOME_URI)

	Local $_sCookie = _GetCookie()
	Local $hFile = FileOpen($COOKIE_FILE, 2)
	FileWrite($hFile, $_sCookie)
	FileClose($hFile)
EndFunc   ;==>Login
#EndRegion API call
