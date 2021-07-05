Include("BitSequence")
'TODO: extends BitSequence
Class NumericEncoder

    Private m_data
    Private m_charCounter
    Private m_bitCounter

    Private Sub Class_Initialize()
        m_data = Empty
        m_charCounter = 0
        m_bitCounter = 0
    End Sub

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = MODE_NUMERIC
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_NUMERIC_VALUE
    End Property

    Public Function Append(ByVal c)
        If m_charCounter Mod 3 = 0 Then
            If m_charCounter = 0 Then
                ReDim m_data(0)
            Else
                ReDim Preserve m_data(UBound(m_data) + 1)
            End If

            m_data(UBound(m_data)) = CLng(c)
        Else
            m_data(UBound(m_data)) = m_data(UBound(m_data)) * 10 + CLng(c)
        End If

        Dim ret
        ret = GetCodewordBitLength(c)
        m_bitCounter = m_bitCounter + ret
        m_charCounter = m_charCounter + 1

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        If m_charCounter Mod 3 = 0 Then
            GetCodewordBitLength = 4
        Else
            GetCodewordBitLength = 3
        End If
    End Function

    Public Function GetBytes()
        Dim bs
        Set bs = New BitSequence

        Dim i
        For i = 0 To UBound(m_data) - 1
            Call bs.Append(m_data(i), 10)
        Next

        Select Case m_charCounter Mod 3
            Case 1
                Call bs.Append(m_data(UBound(m_data)), 4)
            Case 2
                Call bs.Append(m_data(UBound(m_data)), 7)
            Case Else
                Call bs.Append(m_data(UBound(m_data)), 10)
        End Select

        GetBytes = bs.GetBytes()
    End Function

    Public Function InSubset(ByVal c)
        InSubset = "0" <= c And c <= "9"
    End Function

    Public Function InExclusiveSubset(ByVal c)
        InExclusiveSubset = InSubset(c)
    End Function

End Class




