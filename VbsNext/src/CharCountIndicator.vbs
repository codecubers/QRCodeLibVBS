Class CharCountIndicator_

    Public Function GetLength(ByVal ver, ByVal encMode)
        If 1 <= ver And ver <= 9 Then
            Select Case encMode
                Case MODE_NUMERIC
                    GetLength = 10
                Case MODE_ALPHA_NUMERIC
                    GetLength = 9
                Case MODE_BYTE
                    GetLength = 8
                Case MODE_KANJI
                    GetLength = 8
                Case Else
                    Call Err.Raise(5)
            End Select
        ElseIf 10 <= ver And ver <= 26 Then
            Select Case encMode
                Case MODE_NUMERIC
                    GetLength = 12
                Case MODE_ALPHA_NUMERIC
                    GetLength = 11
                Case MODE_BYTE
                    GetLength = 16
                Case MODE_KANJI
                    GetLength = 10
                Case Else
                    Call Err.Raise(5)
            End Select
        ElseIf 27 <= ver And ver <= 40 Then
            Select Case encMode
                Case MODE_NUMERIC
                    GetLength = 14
                Case MODE_ALPHA_NUMERIC
                    GetLength = 13
                Case MODE_BYTE
                    GetLength = 16
                Case MODE_KANJI
                    GetLength = 12
                Case Else
                    Call Err.Raise(5)
            End Select
        Else
            Call Err.Raise(5)
        End If
    End Function

End Class