#include "TreeDisplay.au3"

Func CreateMap()
	Local $mResult[]
	Return $mResult
EndFunc   ;==>CreateMap

Local $mData[]
$mData.x = CreateMap()
$mData.x.a = CreateMap()
$mData.x.a.a = CreateMap()
$mData.x.a.a.a = CreateMap()
$mData.x.a.a.a.a = CreateMap()
$mData.x.a.a.a.a.a = CreateMap()
$mData.x.a.a.a.a.a.a = CreateMap()
$mData.x.a.a.a.a.a.a.a = CreateMap()
$mData.x.a.a.a.a.a.a.a.a = CreateMap()
$mData.x.a.a.a.a.a.a.a.a.a = CreateMap()
$mData.x.a.a.a.a.a.a.a.a.a.a = CreateMap()
$mData.x.a.a.a.a.a.a.a.a.a.a.a = CreateMap()
$mData.x.a.a.a.a.a.a.a.a.a.a.a.a = CreateMap()
$mData.x.a.a.a.a.a.a.a.a.a.a.a.a.a = CreateMap()
$mData.x.a.a.a.a.a.a.a.a.a.a.a.a.a.a = CreateMap()
$mData.num = 123
$mData.str = "abc"
Local $aData = [1, CreateMap(), 3, 4]
$aData[1].zxc = 113
$aData[1].map = CreateMap()
$aData[1].map.xxx = 1231231
$mData.arr = $aData
Local $aData2[0]
$mData.empty = $aData2


TreeDisplay($mData)