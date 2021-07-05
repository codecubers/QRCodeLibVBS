Include("List")
Include("Symbol")

Class Symbols

    Private m_items

    Private m_minVersion
    Private m_maxVersion
    Private m_errorCorrectionLevel
    Private m_structuredAppendAllowed
    Private m_byteModeCharsetName

    Private m_parity

    Private m_currSymbol

    Private m_encNum
    Private m_encAlpha
    Private m_encByte
    Private m_encKanji

    Public Sub Init(ByVal ecLevel, ByVal maxVer, ByVal allowStructuredAppend)
        If Not (MIN_VERSION <= maxVer And maxVer <= MAX_VERSION) Then
            Call Err.Raise(5)
        End If

        Set m_items = New List

        Set m_encNum = CreateEncoder(MODE_NUMERIC)
        Set m_encAlpha = CreateEncoder(MODE_ALPHA_NUMERIC)
        Set m_encByte = CreateEncoder(MODE_BYTE)
        Set m_encKanji = CreateEncoder(MODE_KANJI)

        m_minVersion = 1
        m_maxVersion = maxVer
        m_errorCorrectionLevel = ecLevel
        m_structuredAppendAllowed = allowStructuredAppend

        m_parity = 0

        Set m_currSymbol = New Symbol
        Call m_currSymbol.Init(Me)
        Call m_items.Add(m_currSymbol)
    End Sub

    Public Property Get Item(ByVal idx)
        Set Item = m_items.Item(idx)
    End Property

    Public Property Get Count()
        Count = m_items.Count
    End Property

    Public Property Get StructuredAppendAllowed()
        StructuredAppendAllowed = m_structuredAppendAllowed
    End Property

    Public Property Get Parity()
        Parity = m_parity
    End Property

    Public Property Get MinVersion()
        MinVersion = m_minVersion
    End Property
    Public Property Let MinVersion(ByVal Value)
        m_minVersion = Value
    End Property

    Public Property Get MaxVersion()
        MaxVersion = m_maxVersion
    End Property

    Public Property Get ErrorCorrectionLevel()
        ErrorCorrectionLevel = m_errorCorrectionLevel
    End Property

    Private Function Add()
        Set m_currSymbol = New Symbol
        Call m_currSymbol.Init(Me)
        Call m_items.Add(m_currSymbol)

        Set Add = m_currSymbol
    End Function

    Public Sub AppendText(ByVal s)
        Dim oldMode
        Dim newMode
        Dim i

        If Len(s) = 0 Then Call Err.Raise(5)

        For i = 1 To Len(s)
            oldMode = m_currSymbol.CurrentEncodingMode

            Select Case oldMode
                Case MODE_UNKNOWN
                    newMode = SelectInitialMode(s, i)
                Case MODE_NUMERIC
                    newMode = SelectModeWhileInNumeric(s, i)
                Case MODE_ALPHA_NUMERIC
                    newMode = SelectModeWhileInAlphanumeric(s, i)
                Case MODE_BYTE
                    newMode = SelectModeWhileInByte(s, i)
                Case MODE_KANJI
                    newMode = SelectInitialMode(s, i)
                Case Else
                    Call Err.Raise(51)
            End Select

            If newMode <> oldMode Then
                If Not m_currSymbol.TrySetEncodingMode(newMode, Mid(s, i, 1)) Then
                    If Not m_structuredAppendAllowed Or m_items.Count = 16 Then
                        Call Err.Raise(6)
                    End If

                    Call Add
                    newMode = SelectInitialMode(s, i)
                    Call m_currSymbol.TrySetEncodingMode(newMode, Mid(s, i, 1))
                End If
            End If

            If Not m_currSymbol.TryAppend(Mid(s, i, 1)) Then
                If Not m_structuredAppendAllowed Or m_items.Count = 16 Then
                    Call Err.Raise(6)
                End If

                Call Add
                newMode = SelectInitialMode(s, i)
                Call m_currSymbol.TrySetEncodingMode(newMode, Mid(s, i, 1))
                Call m_currSymbol.TryAppend(Mid(s, i, 1))
            End If
        Next
    End Sub

    Public Sub UpdateParity(ByVal c)
        Dim code
        code = Asc(c) And &HFFFF&

        Dim msb
        Dim lsb

        msb = (code And &HFF00&) \ 2 ^ 8
        lsb = code And &HFF&

        If msb > 0 Then
            m_parity = m_parity Xor msb
        End If

        m_parity = m_parity Xor lsb
    End Sub

    Private Function SelectInitialMode(ByRef s, ByVal startIndex)
        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectInitialMode = MODE_KANJI
            Exit Function
        End If

        If m_encByte.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectInitialMode = MODE_BYTE
            Exit Function
        End If

        If m_encAlpha.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectInitialMode = SelectModeWhenInitialDataAlphaNumeric(s, startIndex)
            Exit Function
        End If

        If m_encNum.InSubset(Mid(s, startIndex, 1)) Then
            SelectInitialMode = SelectModeWhenInitialDataNumeric(s, startIndex)
            Exit Function
        End If

        Call Err.Raise(51)
    End Function

    Private Function SelectModeWhenInitialDataAlphaNumeric(ByRef s, ByVal startIndex)
        Dim cnt
        cnt = 0

        Dim i

        For i = startIndex To Len(s)
            If m_encAlpha.InExclusiveSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            Else
                Exit For
            End If
        Next

        Dim flg
        flg = False

        Dim ver
        ver = m_currSymbol.Version

        If 1 <= ver And ver <= 9 Then
            flg = cnt < 6
        ElseIf 10 <= ver And ver <= 26 Then
            flg = cnt < 7
        ElseIf 27 <= ver And ver <= 40 Then
            flg = cnt < 8
        Else
            Call Err.Raise(51)
        End If

        If flg Then
            If (startIndex + cnt) <= Len(s) Then
                If m_encByte.InSubset(Mid(s, startIndex + cnt, 1)) Then
                    SelectModeWhenInitialDataAlphaNumeric = MODE_BYTE
                    Exit Function
                End If
            End If
        End If

        SelectModeWhenInitialDataAlphaNumeric = MODE_ALPHA_NUMERIC
    End Function

    Private Function SelectModeWhenInitialDataNumeric(ByRef s, ByVal startIndex)
        Dim cnt
        cnt = 0

        Dim i

        For i = startIndex To Len(s)
            If m_encNum.InSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            Else
                Exit For
            End If
        Next

        Dim flg

        Dim ver
        ver = m_currSymbol.Version

        If 1 <= ver And ver <= 9 Then
            flg = cnt < 4
        ElseIf 10 <= ver And ver <= 26 Then
            flg = cnt < 4
        ElseIf 27 <= ver And ver <= 40 Then
            flg = cnt < 5
        Else
            Call Err.Raise(51)
        End If

        If flg Then
            If (startIndex + cnt) <= Len(s) Then
                SelectModeWhenInitialDataNumeric = MODE_BYTE
                Exit Function
            End If
        End If

        If 1 <= ver And ver <= 9 Then
            flg = cnt < 7
        ElseIf 10 <= ver And ver <= 26 Then
            flg = cnt < 8
        ElseIf 27 <= ver And ver <= 40 Then
            flg = cnt < 9
        Else
            Call Err.Raise(51)
        End If

        If flg Then
            If (startIndex + cnt) <= Len(s) Then
                SelectModeWhenInitialDataNumeric = MODE_ALPHA_NUMERIC
                Exit Function
            End If
        End If

        SelectModeWhenInitialDataNumeric = MODE_NUMERIC
    End Function

    Private Function SelectModeWhileInNumeric(ByRef s, ByVal startIndex)
        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInNumeric = MODE_KANJI
            Exit Function
        End If

        If m_encByte.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInNumeric = MODE_BYTE
            Exit Function
        End If

        If m_encAlpha.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInNumeric = MODE_ALPHA_NUMERIC
            Exit Function
        End If

        SelectModeWhileInNumeric = MODE_NUMERIC
    End Function

    Private Function SelectModeWhileInAlphanumeric(ByRef s, ByVal startIndex)
        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInAlphanumeric = MODE_KANJI
            Exit Function
        End If

        If m_encByte.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInAlphanumeric = MODE_BYTE
            Exit Function
        End If

        If MustChangeAlphanumericToNumeric(s, startIndex) Then
            SelectModeWhileInAlphanumeric = MODE_NUMERIC
            Exit Function
        End If

        SelectModeWhileInAlphanumeric = MODE_ALPHA_NUMERIC
    End Function

    Private Function MustChangeAlphanumericToNumeric(ByRef s, ByVal startIndex)
        Dim cnt
        cnt = 0

        Dim ret
        ret = False

        Dim i

        For i = startIndex To Len(s)
            If Not m_encAlpha.InSubset(Mid(s, i, 1)) Then
                Exit For
            End If

            If m_encNum.InSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            Else
                ret = True
                Exit For
            End If
        Next

        Dim ver
        ver = m_currSymbol.Version

        If ret Then
            If 1 <= ver And ver <= 9 Then
                ret = cnt >= 13
            ElseIf 10 <= ver And ver <= 26 Then
                ret = cnt >= 15
            ElseIf 27 <= ver And ver <= 40 Then
                ret = cnt >= 17
            Else
                Call Err.Raise(51)
            End If
        End If

        MustChangeAlphanumericToNumeric = ret
    End Function

    Private Function SelectModeWhileInByte(ByRef s, ByVal startIndex)
        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInByte = MODE_KANJI
            Exit Function
        End If

        If MustChangeByteToNumeric(s, startIndex) Then
            SelectModeWhileInByte = MODE_NUMERIC
            Exit Function
        End If

        If MustChangeByteToAlphanumeric(s, startIndex) Then
            SelectModeWhileInByte = MODE_ALPHA_NUMERIC
            Exit Function
        End If

        SelectModeWhileInByte = MODE_BYTE
    End Function

    Private Function MustChangeByteToNumeric(ByRef s, ByVal startIndex)
        Dim cnt
        cnt = 0

        Dim ret
        ret = False

        Dim i

        For i = startIndex To Len(s)
            If Not m_encByte.InSubset(Mid(s, i, 1)) Then
                Exit For
            End If

            If m_encNum.InSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            ElseIf m_encByte.InExclusiveSubset(Mid(s, i, 1)) Then
                ret = True
                Exit For
            Else
                Exit For
            End If
        Next

        Dim ver
        ver = m_currSymbol.Version

        If ret Then
            If 1 <= ver And ver <= 9 Then
                ret = cnt >= 6
            ElseIf 10 <= ver And ver <= 26 Then
                ret = cnt >= 8
            ElseIf 27 <= ver And ver <= 40 Then
                ret = cnt >= 9
            Else
                Call Err.Raise(51)
            End If
        End If

        MustChangeByteToNumeric = ret
    End Function

    Private Function MustChangeByteToAlphanumeric(ByRef s, ByVal startIndex)
        Dim ret

        Dim cnt
        cnt = 0

        Dim i

        For i = startIndex To Len(s)
            If Not m_encByte.InSubset(Mid(s, i, 1)) Then
                Exit For
            End If

            If m_encAlpha.InExclusiveSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            ElseIf m_encByte.InExclusiveSubset(Mid(s, i, 1)) Then
                ret = True
                Exit For
            Else
                Exit For
            End If
        Next

        Dim ver
        ver = m_currSymbol.Version

        If ret Then
            If 1 <= ver And ver <= 9 Then
                ret = cnt >= 11
            ElseIf 10 <= ver And ver <= 26 Then
                ret = cnt >= 15
            ElseIf 27 <= ver And ver <= 40 Then
                ret = cnt >= 16
            Else
                Call Err.Raise(51)
            End If
        End If

        MustChangeByteToAlphanumeric = ret
    End Function

End Class

