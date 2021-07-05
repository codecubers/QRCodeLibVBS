Class FinderPattern_

    Private m_finderPattern

    Private Sub Class_Initialize()
        Dim VAL
        VAL = FINDER_PTN

        m_finderPattern = Array( _
            Array(VAL,  VAL,  VAL,  VAL,  VAL,  VAL,  VAL), _
            Array(VAL, -VAL, -VAL, -VAL, -VAL, -VAL,  VAL), _
            Array(VAL, -VAL,  VAL,  VAL,  VAL, -VAL,  VAL), _
            Array(VAL, -VAL,  VAL,  VAL,  VAL, -VAL,  VAL), _
            Array(VAL, -VAL,  VAL,  VAL,  VAL, -VAL,  VAL), _
            Array(VAL, -VAL, -VAL, -VAL, -VAL, -VAL,  VAL), _
            Array(VAL,  VAL,  VAL,  VAL,  VAL,  VAL,  VAL) _
        )
    End Sub

    Public Sub Place(ByRef moduleMatrix())
        Dim offset
        offset = (UBound(moduleMatrix) + 1) - (UBound(m_finderPattern) + 1)

        Dim i, j
        Dim v

        For i = 0 To UBound(m_finderPattern)
            For j = 0 To UBound(m_finderPattern(i))
                v = m_finderPattern(i)(j)

                moduleMatrix(i)(j) = v
                moduleMatrix(i)(j + offset) = v
                moduleMatrix(i + offset)(j) = v
            Next
        Next
    End Sub

End Class

