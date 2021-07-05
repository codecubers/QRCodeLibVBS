Include("NumericEncoder")
Include("BitSequence")
'TODO:  extends NumericEncoder (internally extends BitSequence)
Class AlphanumericEncoder

    Private m_data
    Private m_charCounter
    Private m_bitCounter

    Private m_encNumeric

    Private Sub Class_Initialize()
        m_data = Empty
        m_charCounter = 0
        m_bitCounter = 0

        Set m_encNumeric = New NumericEncoder
    End Sub

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = MODE_ALPHA_NUMERIC
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_ALPAHNUMERIC_VALUE
    End Property

    Public Function Append(ByVal c)
        Dim wd
        wd = ConvertCharCode(c)

        If m_charCounter Mod 2 = 0 Then
            If m_charCounter = 0 Then
                ReDim m_data(0)
            Else
                ReDim Preserve m_data(UBound(m_data) + 1)
            End If

            m_data(UBound(m_data)) = wd
        Else
            m_data(UBound(m_data)) = m_data(UBound(m_data)) * 45
            m_data(UBound(m_data)) = m_data(UBound(m_data)) + wd
        End If

        Dim ret
        ret = GetCodewordBitLength(c)
        m_bitCounter = m_bitCounter + ret
        m_charCounter = m_charCounter + 1

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        If m_charCounter Mod 2 = 0 Then
            GetCodewordBitLength = 6
        Else
            GetCodewordBitLength = 5
        End If
    End Function

    Public Function GetBytes()
        Dim bs
        Set bs = New BitSequence

        Dim bitLength
        bitLength = 11

        Dim i
        For i = 0 To UBound(m_data) - 1
            Call bs.Append(m_data(i), bitLength)
        Next

        If m_charCounter Mod 2 = 0 Then
            bitLength = 11
        Else
            bitLength = 6
        End If

        Call bs.Append(m_data(UBound(m_data)), bitLength)

        GetBytes = bs.GetBytes()
    End Function

    Public Function ConvertCharCode(ByVal c)
        Dim code
        code = Asc(c)

        ' (Space)
        If code = 32 Then
            ConvertCharCode = 36
        ' $ %
        ElseIf code = 36 Or code = 37 Then
            ConvertCharCode = code + 1
        ' * +
        ElseIf code = 42 Or code = 43 Then
            ConvertCharCode = code - 3
        ' - .
        ElseIf code = 45 Or code = 46 Then
            ConvertCharCode = code - 4
        ' /
        ElseIf code = 47 Then
            ConvertCharCode = 43
        ' 0 - 9
        ElseIf 48 <= code And code <= 57 Then
            ConvertCharCode = code - 48
        ' :
        ElseIf code = 58 Then
            ConvertCharCode = 44
        ' A - Z
        ElseIf 65 <= code And code <= 90 Then
            ConvertCharCode = code - 55
        Else
            ConvertCharCode = -1
        End If
    End Function

    Public Function InSubset(ByVal c)
        InSubset = ConvertCharCode(c) > -1
    End Function

    Public Function InExclusiveSubset(ByVal c)
        If m_encNumeric.InSubset(c) Then
            InExclusiveSubset = False
            Exit Function
        End If

        InExclusiveSubset = InSubset(c)
    End Function

End Class