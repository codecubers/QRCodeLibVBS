Class Separator_

    Public Sub Place(ByRef moduleMatrix())
        DIM VAL
        VAL = SEPARATOR_PTN

        Dim offset
        offset = UBound(moduleMatrix) - 7

        Dim i
        For i = 0 To 7
             moduleMatrix(i)(7) = -VAL
             moduleMatrix(7)(i) = -VAL

             moduleMatrix(offset + i)(7) = -VAL
             moduleMatrix(offset + 0)(i) = -VAL

             moduleMatrix(i)(offset + 0) = -VAL
             moduleMatrix(7)(offset + i) = -VAL
         Next
    End Sub

End Class


