Class RemainderBit_

    Public Sub Place(ByRef moduleMatrix())
        Dim r, c

        For r = 0 To UBound(moduleMatrix)
            For c = 0 To UBound(moduleMatrix(r))
                If moduleMatrix(r)(c) = BLANK Then
                    moduleMatrix(r)(c) = -WORD
                End If
            Next
        Next
    End Sub

End Class