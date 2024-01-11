#include <WinAPICom.au3>
#include <Misc.au3>

$sFileName = _WinAPI_CreateGUID()
$sData = ""

For $jj=1 To 500000
	For $ii=1 To 1000
		$sData &= Random(0,9,1)
	Next
	$sData &= @CRLF
Next

$hFile = FileOpen($sFileName&".pdf", 3)
_StringToPDF($sData, $hFile)
FileClose($hFile)


; Function Name:    _StringToPDF()
; Description:      Create PDF File
; Parameter(s):     $Text               - Text in the PDF (I love Autoit)
;                   $File               - Path and Filename of the PDF (c:\Test.pdf)
;                   $Size               - Papersize A4 or A3
;                   $Rand_x             - Spacing between the text and the vertical edges, Default 20
;                   $Rand_y             - Spacing between the text and the horizontal edges, Default 24
;                   $Schriftart         - Fonts (Times-Roman, Helvetica and Courier), Default Courier
;                   $Fett               - Value if the font is bold, Default 0 (normal)
;                   $Kursiv             - Value if the font is italic, Default 0 (normal)
;                   $Schrift            - Size of the font, Default 12
;                   $Autor              - Name of the autor, Default "unknown"
;                   $Titel              - Title of the PDF, Default "MyPDF"
;
; Author(s):        Christian Korittke <Christian_Korittke@web.de>
;                   Tamer Hosgör <Tamer@TamTech.info>
; https://www.autoitscript.com/forum/topic/32261-_stringtopdf/



Func _StringToPDF ( $Text, $File, $Size="A4", $Rand_x=20, $Rand_y=24, $Schriftart="Courier", $Fett=0, $Kursiv=0, $Schrift=12, $Autor="unknown", $Titel="MyPDF")

    If $Size        = "A4" Then
        $Size_x     = 210
        $Size_y     = 297
    ElseIf $Size    = "A3" Then
        $Size_x     = 420
        $Size_y     = 297
    EndIf

               $Zeilen      =   1

    If $Fett = 1 Or $Kursiv = 1 Then
        If $Schriftart = "Times-Roman" Then
            If $Fett = 1 Then
                $Schriftart = "Times-Bold"
            ElseIf $Kursiv = 1 Then
                $Schriftart = "Times-Italic"
            EndIf
            If $Fett = 1 And $Kursiv = 1 Then $Schriftart =  "Times-BoldItalic"
        ElseIf $Schriftart = "Helvetica" Then
            If $Fett = 1 Then
                $Schriftart = "Helvetica-Bold"
            ElseIf $Kursiv = 1 Then
                $Schriftart = "Helvetica-Oblique"
            EndIf
            If $Fett = 1 And $Kursiv = 1 Then $Schriftart =  "Helvetica-BoldOblique"
        Else
            If $Fett = 1 Then
                $Schriftart = "Courier-Bold"
            ElseIf $Kursiv = 1 Then
                $Schriftart = "Courier-Oblique"
            EndIf
            If $Fett = 1 And $Kursiv = 1 Then $Schriftart =  "Courier-BoldOblique"
        EndIf
    EndIf


    If $Schrift     = 8 Then
        $Abstand    = 9
    ElseIf $Schrift = 9 Then
        $Abstand    = 11
    ElseIf $Schrift = 10 Then
        $Abstand    = 12
    ElseIf $Schrift = 11 Then
        $Abstand    = 13
    ElseIf $Schrift = 12 Then
        $Abstand    = 15
    ElseIf $Schrift = 14 Then
        $Abstand    = 17
    ElseIf $Schrift = 16 Then
        $Abstand    = 19
    ElseIf $Schrift = 18 Then
        $Abstand    = 21
    ElseIf $Schrift = 20 Then
        $Abstand    = 24
    ElseIf $Schrift = 22 Then
        $Abstand    = 26
    ElseIf $Schrift = 24 Then
        $Abstand    = 28
    ElseIf $Schrift = 26 Then
        $Abstand    = 30
    ElseIf $Schrift = 28 Then
        $Abstand    = 32
    ElseIf $Schrift = 36 Then
        $Abstand    = 41
    ElseIf $Schrift = 48 Then
        $Abstand    = 55
    EndIf










    If Not StringInStr($Text,@CRLF) = 0 Then
        $Text = StringSplit($Text,@CRLF)
        $Zeilen = $Text[0] / 2 + 1
    EndIf

    ; Umrechnung
    $Wert           = 2.834175
    $Size_y         = Round($Size_y * $Wert)
    $Size_x         = Round($Size_x * $Wert)
    $Rand_x         = Round($Rand_x * $Wert)
    $Rand_y         = Round($Rand_y * $Wert)

    FileWriteLine($File,"%PDF-1.2")
    FileWriteLine($File,"%âãÏÓ")

    FileWriteLine($File,"1 0 obj")
    FileWriteLine($File,"<<")
    FileWriteLine($File,"/Author ("&$Autor&")")
    FileWriteLine($File,"/CreationDate (D:"&@YEAR&@MON&@MDAY&@HOUR&@MIN&@SEC&")")
    FileWriteLine($File,"/Creator (Ahnungslos)")
    FileWriteLine($File,"/Producer (Ahnungslos)")
    FileWriteLine($File,"/Title ("&$Titel&")")
    FileWriteLine($File,">>")
    FileWriteLine($File,"endobj")

    FileWriteLine($File,"4 0 obj")
    FileWriteLine($File,"<<")
    FileWriteLine($File,"/Type /Font")
    FileWriteLine($File,"/Subtype /Type1")
    FileWriteLine($File,"/Name /F1")
    FileWriteLine($File,"/Encoding 5 0 R")
    FileWriteLine($File,"/BaseFont /"&$Schriftart)
    FileWriteLine($File,">>")
    FileWriteLine($File,"endobj")

    FileWriteLine($File,"5 0 obj")
    FileWriteLine($File,"<<")
    FileWriteLine($File,"/Type /Encoding")
    FileWriteLine($File,"/BaseEncoding /WinAnsiEncoding")
    FileWriteLine($File,">>")
    FileWriteLine($File,"endobj")

    FileWriteLine($File,"6 0 obj")
    FileWriteLine($File,"<<")
    FileWriteLine($File,"  /Font << /F1 4 0 R >>")
    FileWriteLine($File,"  /ProcSet [ /PDF /Text ]")
    FileWriteLine($File,">>")
    FileWriteLine($File,"endobj")

    FileWriteLine($File,"7 0 obj")
    FileWriteLine($File,"<<")
    FileWriteLine($File,"/Type /Page")
    FileWriteLine($File,"/Parent 3 0 R")
    FileWriteLine($File,"/Resources 6 0 R")
    FileWriteLine($File,"/Contents 8 0 R")
    FileWriteLine($File,"/Rotate 0")
    FileWriteLine($File,">>")
    FileWriteLine($File,"endobj")

    FileWriteLine($File,"8 0 obj")
    FileWriteLine($File,"<<")
    FileWriteLine($File,"/Length 9 0 R")
    FileWriteLine($File,">>")
    FileWriteLine($File,"stream")
    FileWriteLine($File,"BT")

    If $Zeilen = 1 Then
        FileWriteLine($File,"/F1 "&$Schrift&" Tf")
        FileWriteLine($File,"1 0 0 1 "&$Rand_y&" "&$Size_y - $Rand_x - $Abstand&" Tm")
        FileWriteLine($File,"("&$Text&") Tj")
    Else
        For $Counter = 1 To $Zeilen
            FileWriteLine($File,"/F1 "&$Schrift&" Tf")
            FileWriteLine($File,"1 0 0 1 "&$Rand_y&" "&$Size_y - $Rand_x - $Abstand * $Counter&" Tm")
            FileWriteLine($File,"("&$Text[$Counter * 2 - 1]&") Tj")
        Next
    EndIf

    FileWriteLine($File,"ET")
    FileWriteLine($File,"endstream")
    FileWriteLine($File,"endobj")

    FileWriteLine($File,"9 0 obj")
    FileWriteLine($File,"78")
    FileWriteLine($File,"endobj")

    FileWriteLine($File,"2 0 obj")
    FileWriteLine($File,"<<")
    FileWriteLine($File,"/Type /Catalog")
    FileWriteLine($File,"/Pages 3 0 R")
    FileWriteLine($File,">>")
    FileWriteLine($File,"endobj")

    FileWriteLine($File,"3 0 obj")
    FileWriteLine($File,"<<")
    FileWriteLine($File,"/Type /Pages")
    FileWriteLine($File,"/Count 1")
    FileWriteLine($File,"/MediaBox [ 0 0 "&$Size_x&" "&$Size_y&" ]")
    FileWriteLine($File,"/Kids [ 7 0 R ]")
    FileWriteLine($File,">>")
    FileWriteLine($File,"endobj")

    FileWriteLine($File,"0 10")
    FileWriteLine($File,"0000000000 65535 f ")
    FileWriteLine($File,"0000000013 00000 n ")
    FileWriteLine($File,"0000000591 00000 n ")
    FileWriteLine($File,"0000000634 00000 n ")
    FileWriteLine($File,"0000000156 00000 n ")
    FileWriteLine($File,"0000000245 00000 n ")
    FileWriteLine($File,"0000000307 00000 n ")
    FileWriteLine($File,"0000000372 00000 n ")
    FileWriteLine($File,"0000000453 00000 n ")
    FileWriteLine($File,"0000000576 00000 n ")
    FileWriteLine($File,"trailer")
    FileWriteLine($File,"<<")
    FileWriteLine($File,"/Size 10")
    FileWriteLine($File,"/Root 2 0 R")
    FileWriteLine($File,"/Info 1 0 R")
    FileWriteLine($File,">>")
    FileWriteLine($File,"startxref")
    FileWriteLine($File,"712")
    FileWriteLine($File,"%%EOF")

    FileClose($File)

EndFunc