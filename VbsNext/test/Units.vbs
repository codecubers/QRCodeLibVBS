Option Explicit
Include "QRCode.vbs"

Const FORE_COLOR = "#000000"
Const BACK_COLOR = "#FFFFFF"
Const SCALE = 5



Public Sub Example1()
    Dim sbls: Set sbls = CreateSymbols(ECR_M, 40, False)
    Call sbls.AppendText("Hello World")

    ' 24bpp bitmap
'    Call sbls.Item(0).Save24bppDIB("qrcode.bmp", SCALE, FORE_COLOR, BACK_COLOR)
    ' 1bpp bitmap
    Call sbls.Item(0).Save1bppDIB("test\\units-qrcode.bmp", SCALE, FORE_COLOR, BACK_COLOR)
    ' SVG
'    Call sbls.Item(0).SaveSvg("qrcode.svg", SCALE, FORE_COLOR)
End Sub


Public Sub Example2()
    Dim sbls: Set sbls = CreateSymbols(ECR_M, 1, True)
    Call sbls.AppendText("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ")

    Dim fName
    Dim i
    For i = 0 To sbls.Count - 1
        fName = "test\\units-qrcode" & CStr(i) & ".bmp"
        Call sbls.Item(i).Save24bppDIB(fName, SCALE, FORE_COLOR, BACK_COLOR)
    Next
End Sub



Call Example1
Call Example2