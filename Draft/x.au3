#include <Misc.au3>

;~ While Not _IsPressed("10")
;~ 	$aMouse = MouseGetPos()
;~ 	ToolTip($aMouse[0]&"/"&$aMouse[1],0,0)
;~ 	Sleep(100)
;~ WEnd

;~ Sleep(1000)
;~ While Not _IsPressed("10")
;~ 	MouseMove(Random(0, @DesktopWidth,1), Random(0, @DesktopHeight,1), 5)
;~ 	MouseClick("left", 404, 299, 1, 0)
;~ 	MouseClick("left", 935, 602, 1, 0)
;~ 	MouseClick("left", 1016, 768, 1, 0)
;~ WEnd
m

Sleep(3000)
Send("{W DOWN}")
While Not _IsPressed("10")
	Sleep(100)
WEnd
Send("{W UP}")
