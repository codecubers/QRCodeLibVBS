Include("AlphanumericEncoder")
Include("KanjiEncoder")

Class ByteEncoder

    Private m_data
    Private m_charCounter
    Private m_bitCounter

    Private m_encAlpha
    Private m_encKanji

    Private Sub Class_Initialize()
        m_data = Empty
        m_charCounter = 0
        m_bitCounter = 0

        Set m_encAlpha = New AlphanumericEncoder
        Set m_encKanji = New KanjiEncoder
    End Sub

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = MODE_BYTE
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_BYTE_VALUE
    End Property

    Public Function Append(ByVal c)
        If m_charCounter = 0 Then
            ReDim m_data(0)
        Else
            ReDim Preserve m_data(UBound(m_data) + 1)
        End If

        Dim wd
        wd = Asc(c) And &HFFFF&
        m_data(UBound(m_data)) = wd

        Dim ret
        ret = GetCodewordBitLength(c)
        m_bitCounter = m_bitCounter + ret
        m_charCounter = m_charCounter + (ret \ 8)

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        Dim code
        code = Asc(c) And &HFFFF&

        If code > &HFF Then
            GetCodewordBitLength = 16
        Else
            GetCodewordBitLength = 8
        End If
    End Function

    Public Function GetBytes()
        GetBytes = m_data
    End Function

    Public Function InSubset(ByVal c)
        InSubset = True
    End Function

    Public Function InExclusiveSubset(ByVal c)
        If m_encAlpha.InSubset(c) Then
            InExclusiveSubset = False
            Exit Function
        End If

        If m_encKanji.InSubset(c) Then
            InExclusiveSubset = False
            Exit Function
        End If

        InExclusiveSubset = InSubset(c)
    End Function

End Class