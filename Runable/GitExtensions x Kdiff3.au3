;https://download.kde.org/stable/kdiff3/?C=M;O=D
#include <_HttpRequest\_HttpRequest.au3>
#include <StringExtended\StringExtended.au3>

InstallGitExtensions()

Func InstallGitExtensions()
	If FileExists("@ProgramFilesDir@ (x86)\GitExtensions\GitExtensions.exe") Or _
		FileExists("@ProgramFilesDir@\GitExtensions\GitExtensions.exe") Then
		Return -1
	EndIf

	_HttpRequest('*0', 'https://github.com/gitextensions/gitextensions/releases/latest')
	Local $sLocation = _HttpRequest_QueryHeaders(33)
	Local $sLatestVer = _StringRegExpFirstOrError($sLocation, ".*\/(.+?)$")

	Local $sAssetList = _HttpRequest('2', 'https://github.com/gitextensions/gitextensions/releases/expanded_assets/$sLatestVer$')
	Local $aAssetList = StringRegExp($sAssetList, "href=['""](.+?)['""]", 3)

	Local $bPassed = False

	For $sURL In $aAssetList
		If StringRegExp($sURL, "\.msi$") Then
			$bPassed = True
			ExitLoop
		EndIf
	Next

	If Not $bPassed Then
		ConsoleWrite("!> ERROR: Cannot retrieve asset URL!!!")
		Exit
	EndIf

	$hexAsset = _HttpRequest(3, "https://github.com$sURL$")
	$sDestFolder = "@TempDir@\GExKD"
	$sInstaller = "$sDestFolder$\GitExtensions.msi"

	If FileExists($sInstaller) Then
		FileDelete($sInstaller)
	EndIf

	DirCreate($sDestFolder)
	FileWrite($sInstaller, $hexAsset)
	$hexAsset = 0

	Run("MsiExec.exe /i ""$sInstaller$""")
EndFunc