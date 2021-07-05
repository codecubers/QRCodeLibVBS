Option Explicit

Include("src\Symbols")
Include("src\NumericEncoder")
Include("src\AlphanumericEncoder")
Include("src\ByteEncoder")
Include("src\KanjiEncoder")

Include("src\AlignmentPattern")
Include("src\CharCountIndicator")
Include("src\Codeword")
Include("src\DataCodeword")
Include("src\FinderPattern")
Include("src\FormatInfo")
Include("src\GaloisField256")
Include("src\GeneratorPolynomials")
Include("src\Masking")
Include("src\MaskingPenaltyScore")
Include("src\Module")
Include("src\QuietZone")
Include("src\RemainderBit")
Include("src\RSBlock")
Include("src\Separator")
Include("src\TimingPattern")
Include("src\VersionInfo")
Include("src\ColorCode")
Include("src\Graphics")

Public Const MIN_VERSION = 1
Public Const MAX_VERSION = 40

Public Const ECR_L = 0
Public Const ECR_M = 1
Public Const ECR_Q = 2
Public Const ECR_H = 3

Private Const MODE_UNKNOWN       = 0
Private Const MODE_NUMERIC       = 1
Private Const MODE_ALPHA_NUMERIC = 2
Private Const MODE_BYTE          = 3
Private Const MODE_KANJI         = 4

Private Const MODEINDICATOR_LENGTH = 4
Private Const MODEINDICATOR_TERMINATOR_VALUE           = &H0
Private Const MODEINDICATOR_NUMERIC_VALUE              = &H1
Private Const MODEINDICATOR_ALPAHNUMERIC_VALUE         = &H2
Private Const MODEINDICATOR_STRUCTURED_APPEND_VALUE    = &H3
Private Const MODEINDICATOR_BYTE_VALUE                 = &H4
Private Const MODEINDICATOR_KANJI_VALUE                = &H8

Private Const SYMBOLSEQUENCEINDICATOR_POSITION_LENGTH     = 4
Private Const SYMBOLSEQUENCEINDICATOR_TOTAL_NUMBER_LENGTH = 4

Private Const STRUCTUREDAPPEND_PARITY_DATA_LENGTH = 8
Private Const STRUCTUREDAPPEND_HEADER_LENGTH      = 20

Private Const QUIET_ZONE_MIN_WIDTH = 4

Private Const adTypeBinary = 1
Private Const adTypeText = 2
Private Const adSaveCreateOverWrite = 2

Private Const DIRECTION_UP = 0
Private Const DIRECTION_DOWN = 1
Private Const DIRECTION_LEFT = 2
Private Const DIRECTION_RIGHT = 3

Private Const MIN_MODULE_SIZE = 2

Private Const BLANK         = 0
Private Const WORD          = 1
Private Const ALIGNMENT_PTN = 2
Private Const FINDER_PTN    = 3
Private Const FORMAT_INFO   = 4
Private Const SEPARATOR_PTN = 5
Private Const TIMING_PTN    = 6
Private Const VERSION_INFO  = 7

Private AlignmentPattern:     Set AlignmentPattern = New AlignmentPattern_
Private CharCountIndicator:   Set CharCountIndicator = New CharCountIndicator_
Private Codeword:             Set Codeword = New Codeword_
Private DataCodeword:         Set DataCodeword = New DataCodeword_
Private FinderPattern:        Set FinderPattern = New FinderPattern_
Private FormatInfo:           Set FormatInfo = New FormatInfo_
Private GaloisField256:       Set GaloisField256 = New GaloisField256_
Private GeneratorPolynomials: Set GeneratorPolynomials = New GeneratorPolynomials_
Private Masking:              Set Masking = New Masking_
Private MaskingPenaltyScore:  Set MaskingPenaltyScore = New MaskingPenaltyScore_
Private Module:               Set Module = New Module_
Private QuietZone:            Set QuietZone = New QuietZone_
Private RemainderBit:         Set RemainderBit = New RemainderBit_
Private RSBlock:              Set RSBlock = New RSBlock_
Private Separator:            Set Separator = New Separator_
Private TimingPattern:        Set TimingPattern = New TimingPattern_
Private VersionInfo:          Set VersionInfo = New VersionInfo_
Private ColorCode:            Set ColorCode = New ColorCode_
Private Graphics:             Set Graphics = New Graphics_


Call Main(WScript.Arguments)


Public Sub Main(ByVal args)
    If args.Count = 0 Then Exit Sub

    Dim params
    Set params = GetParams(args)
    If params Is Nothing Then
        Call WScript.Quit(-1)
    End If

    Dim sbls
    Set sbls = CreateSymbols(params("ecr"), MAX_VERSION, False)
    Call sbls.AppendText(params("data"))

    Select Case params("filetype")
        Case "bmp"
            Select Case params("colordepth")
                Case 1
                    Call sbls.Item(0).Save1bppDIB( _
                        params("out"), params("scale"), params("forecolor"), params("backcolor"))
                Case 24
                    Call sbls.Item(0).Save24bppDIB( _
                        params("out"), params("scale"), params("forecolor"), params("backcolor"))
                Case Else
                    Call Err.Raise(51)
            End Select
        Case "svg"
            Call sbls.Item(0).SaveSvg(params("out"), params("scale"), params("forecolor"))
        Case Else
            Call Err.Raise(51)
    End Select

    Call WScript.Quit(0)
End Sub

Private Function GetParams(ByVal args)
    Dim ks
    ks = Array("data", "out", "forecolor", "backcolor", "colordepth", "ecr", "scale", "filetype")

    Dim params
    Set params = CreateObject("Scripting.Dictionary")
    Dim k, v

    For Each k In ks
        Call params.Add(k, Empty)
    Next

    Dim fso, ts
    Set fso = CreateObject("Scripting.FileSystemObject")

    If args.UnNamed.Count > 0 Then
        If Not fso.FileExists(args.UnNamed(0)) Then
            Call WScript.Echo("file not found")
            Exit Function
        End If

        Set ts = fso.OpenTextFile(args.UnNamed(0))
        params("data") = ts.ReadAll()
        ts.Close
    End If

    params("scale") = 5
    params("forecolor") = ColorCode.BLACK
    params("backcolor") = ColorCode.WHITE
    params("colordepth") = 24
    params("ecr") = "M"
    params("filetype") = "bmp"

    For Each k In ks
        If args.Named.Exists(k) Then
            v = args.Named.Item(k)
            If Len(v) = 0 Then
                Call WScript.Echo("argument error '" & k  & "'")
                Exit Function
            End If
            If IsNumeric(v) Then
                v = CLng(v)
            End If
            params(k) = v
        End IF
    Next

    If Len(params("out")) = 0 Then
        Call WScript.Echo("argument error 'out'")
        Exit Function
    End If

    If params("colordepth")  <> 1 And params("colordepth") <> 24 Then
        Call WScript.Echo("argument error 'colordepth'")
        Exit Function
    End If

    If Not ColorCode.IsWebColor(params("forecolor")) Then
        Call WScript.Echo("argument error 'forecolor'")
        Exit Function
    End If

    If Not ColorCode.IsWebColor(params("backcolor")) Then
        Call WScript.Echo("argument error 'backcolor'")
        Exit Function
    End If

    If Not IsNumeric(params("scale")) Then
        Call WScript.Echo("argument error 'scale'")
        Exit Function
    End If

    If params("scale") < MIN_MODULE_SIZE Then
        Call WScript.Echo("argument error 'scale'")
        Exit Function
    End If

    Select Case UCase(params("ecr"))
        Case "L"
            v = ECR_L
        Case "M"
            v = ECR_M
        Case "Q"
            v = ECR_Q
        Case "H"
            v = ECR_H
        Case Else
            Call WScript.Echo("argument error 'ecr'")
            Exit Function
    End Select
    params("ecr") = v

    params("filetype") = LCase(fso.GetExtensionName(params("out")))
    Select Case params("filetype")
        Case "bmp", "svg"
            ' NOP
        Case Else
            Call WScript.Echo("argument error 'out'")
            Exit Function
    End Select

    Set GetParams = params
End Function

Public Function CreateSymbols(ByVal ecLevel, ByVal maxVer, ByVal allowStructuredAppend)
    Select Case ecLevel
        Case ECR_L ,ECR_M, ECR_Q, ECR_H
            ' NOP
        Case Else
            Call Err.Raise(5)
    End Select

    If Not (1 <= maxVer And maxVer <= 40) Then Call Err.Raise(5)

    Dim ret
    Set ret = New Symbols
    Call ret.Init(ecLevel, maxVer, allowStructuredAppend)

    Set CreateSymbols = ret
End Function

Private Function CreateEncoder(ByVal encMode)
    Dim ret

    Select Case encMode
        Case MODE_NUMERIC
            Set ret = New NumericEncoder
        Case MODE_ALPHA_NUMERIC
            Set ret = New AlphanumericEncoder
        Case MODE_BYTE
            Set ret = New ByteEncoder
        Case MODE_KANJI
            Set ret = New KanjiEncoder
        Case Else
            Call Err.Raise(5)
    End Select

    Set CreateEncoder = ret
End Function

Private Function IsDark(ByVal arg)
    IsDark = arg > BLANK
End Function


