Include("AlphanumericEncoder")
Include("BitSequence")

Class KanjiEncoder

    Private m_data
    Private m_charCounter
    Private m_bitCounter

    Private m_encAlpha

    Private Sub Class_Initialize()
        m_data = Empty
        m_charCounter = 0
        m_bitCounter = 0

        Set m_encAlpha = New AlphanumericEncoder
    End Sub

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = MODE_KANJI
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_KANJI_VALUE
    End Property

    Public Function Append(ByVal c)
        Dim wd
        wd = Asc(c) And &HFFFF&

        If &H8140& <= wd And wd <= &H9FFC& Then
            wd = wd - &H8140&
        ElseIf &HE040& <= wd And wd <= &HEBBF& Then
            wd = wd - &HC140&
        Else
            Call Err.Raise(5)
        End If

        wd = ((wd \ 2 ^ 8) * &HC0&) + (wd And &HFF&)
        If m_charCounter = 0 Then
            ReDim m_data(0)
        Else
            ReDim Preserve m_data(UBound(m_data) + 1)
        End If

        m_data(UBound(m_data)) = wd

        Dim ret
        ret = GetCodewordBitLength(c)
        m_bitCounter = m_bitCounter + ret
        m_charCounter = m_charCounter + 1

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        GetCodewordBitLength = 13
    End Function

    Public Function GetBytes()
        Dim bs
        Set bs = New BitSequence

        Dim v
        For Each v In m_data
            Call bs.Append(v, 13)
        Next

        GetBytes = bs.GetBytes()
    End Function

    Public Function InSubset(ByVal c)
        Dim code
        code = Asc(c) And &HFFFF&

        Dim lsb
        lsb = code And &HFF&

        If &H8140& <= code And code <= &H9FFC& Or _
           &HE040& <= code And code <= &HEBBF& Then
            InSubset = &H40& <= lsb And lsb <= &HFC& And _
                       &H7F& <> lsb
            Exit Function
        End If

        InSubset = False
    End Function

    Public Function InExclusiveSubset(ByVal c)
        If m_encAlpha.InSubset(c) Then
            InExclusiveSubset = False
        End If

        InExclusiveSubset = InSubset(c)
    End Function

End Class

