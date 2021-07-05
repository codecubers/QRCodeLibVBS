Class QuietZone_

    Private m_width

    Private Sub Class_Initialize()
        m_width = QUIET_ZONE_MIN_WIDTH
    End Sub

    Public Property Get Width()
        Width = m_width
    End Property
    Public Property Let Width(ByVal Value)
        If Value < QUIET_ZONE_MIN_WIDTH Then
            Call Err.Raise(5)
        End If

        m_width = Value
    End Property

    Public Function Place(ByRef moduleMatrix())
        Dim ret()
        ReDim ret(UBound(moduleMatrix) + Width * 2)

        Dim i
        Dim cols()

        For i = 0 To UBound(ret)
            ReDim cols(UBound(ret))
            ret(i) = cols
        Next

        Dim r, c

        For r = 0 To UBound(moduleMatrix)
            For c = 0 To UBound(moduleMatrix(r))
                ret(r + Width)(c + Width) = moduleMatrix(r)(c)
            Next
        Next

        Place = ret
    End Function

End Class

