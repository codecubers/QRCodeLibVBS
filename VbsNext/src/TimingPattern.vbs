Class TimingPattern_

    Public Sub Place(ByRef moduleMatrix())
        Dim i
        Dim v

        For i = 8 To UBound(moduleMatrix) - 8
            If i Mod 2 = 0 Then
                v = TIMING_PTN
            Else
                v = -TIMING_PTN
            End If

            moduleMatrix(6)(i) = v
            moduleMatrix(i)(6) = v
        Next
    End Sub

End Class