Include("MaskingConditions")

Class Masking_

    Public Function Apply(ByVal ver, ByVal ecLevel, ByRef moduleMatrix)
        Dim minPenalty
        minPenalty = &H7FFFFFFF

        Dim temp
        Dim penalty
        Dim maskPatternReference
        Dim maskedMatrix

        Dim i

        For i = 0 To 7
            temp = moduleMatrix

            Call Mask(i, temp)
            Call FormatInfo.Place(ecLevel, i, temp)

            If ver >= 7 Then
                Call VersionInfo.Place(ver, temp)
            End If

            penalty = MaskingPenaltyScore.CalcTotal(temp)

            If penalty < minPenalty Then
                minPenalty = penalty
                maskPatternReference = i
                maskedMatrix = temp
            End If
        Next

        moduleMatrix = maskedMatrix
        Apply = maskPatternReference
    End Function

    Private Sub Mask(ByVal maskPatternReference, ByRef moduleMatrix())
        Dim condition
        Set condition = GetCondition(maskPatternReference)

        Dim r, c

        For r = 0 To UBound(moduleMatrix)
            For c = 0 To UBound(moduleMatrix(r))
                If Abs(moduleMatrix(r)(c)) = WORD Then
                    If condition.Evaluate(r, c) Then
                        moduleMatrix(r)(c) = moduleMatrix(r)(c) * -1
                    End If
                End If
            Next
        Next
    End Sub

    Private Function GetCondition(ByVal maskPatternReference)
        Dim ret

        Select Case maskPatternReference
            Case 0
                Set ret = New MaskingCondition0
            Case 1
                Set ret = New MaskingCondition1
            Case 2
                Set ret = New MaskingCondition2
            Case 3
                Set ret = New MaskingCondition3
            Case 4
                Set ret = New MaskingCondition4
            Case 5
                Set ret = New MaskingCondition5
            Case 6
                Set ret = New MaskingCondition6
            Case 7
                Set ret = New MaskingCondition7
            Case Else
                Call Err.Raise(5)
        End Select

        Set GetCondition = ret
    End Function

End Class


