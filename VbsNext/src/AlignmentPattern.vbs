Class AlignmentPattern_

    Private m_lst(40)

    Private Sub Class_Initialize()
        m_lst(2)  = Array(6, 18)
        m_lst(3)  = Array(6, 22)
        m_lst(4)  = Array(6, 26)
        m_lst(5)  = Array(6, 30)
        m_lst(6)  = Array(6, 34)
        m_lst(7)  = Array(6, 22, 38)
        m_lst(8)  = Array(6, 24, 42)
        m_lst(9)  = Array(6, 26, 46)
        m_lst(10) = Array(6, 28, 50)
        m_lst(11) = Array(6, 30, 54)
        m_lst(12) = Array(6, 32, 58)
        m_lst(13) = Array(6, 34, 62)
        m_lst(14) = Array(6, 26, 46, 66)
        m_lst(15) = Array(6, 26, 48, 70)
        m_lst(16) = Array(6, 26, 50, 74)
        m_lst(17) = Array(6, 30, 54, 78)
        m_lst(18) = Array(6, 30, 56, 82)
        m_lst(19) = Array(6, 30, 58, 86)
        m_lst(20) = Array(6, 34, 62, 90)
        m_lst(21) = Array(6, 28, 50, 72, 94)
        m_lst(22) = Array(6, 26, 50, 74, 98)
        m_lst(23) = Array(6, 30, 54, 78, 102)
        m_lst(24) = Array(6, 28, 54, 80, 106)
        m_lst(25) = Array(6, 32, 58, 84, 110)
        m_lst(26) = Array(6, 30, 58, 86, 114)
        m_lst(27) = Array(6, 34, 62, 90, 118)
        m_lst(28) = Array(6, 26, 50, 74, 98, 122)
        m_lst(29) = Array(6, 30, 54, 78, 102, 126)
        m_lst(30) = Array(6, 26, 52, 78, 104, 130)
        m_lst(31) = Array(6, 30, 56, 82, 108, 134)
        m_lst(32) = Array(6, 34, 60, 86, 112, 138)
        m_lst(33) = Array(6, 30, 58, 86, 114, 142)
        m_lst(34) = Array(6, 34, 62, 90, 118, 146)
        m_lst(35) = Array(6, 30, 54, 78, 102, 126, 150)
        m_lst(36) = Array(6, 24, 50, 76, 102, 128, 154)
        m_lst(37) = Array(6, 28, 54, 80, 106, 132, 158)
        m_lst(38) = Array(6, 32, 58, 84, 110, 136, 162)
        m_lst(39) = Array(6, 26, 54, 82, 110, 138, 166)
        m_lst(40) = Array(6, 30, 58, 86, 114, 142, 170)
    End Sub

    Public Sub Place(ByVal ver, ByRef moduleMatrix())
        Dim VAL
        VAL = ALIGNMENT_PTN

        Dim centerArray
        centerArray = m_lst(ver)

        Dim maxIndex
        maxIndex = UBound(centerArray)

        Dim i, j
        Dim r, c

        For i = 0 To maxIndex
            r = centerArray(i)

            For j = 0 To maxIndex
                c = centerArray(j)

                If (i = 0 And j = 0 Or _
                    i = 0 And j = maxIndex Or _
                    i = maxIndex And j = 0) = False Then

                    moduleMatrix(r - 2)(c - 2) = VAL
                    moduleMatrix(r - 2)(c - 1) = VAL
                    moduleMatrix(r - 2)(c + 0) = VAL
                    moduleMatrix(r - 2)(c + 1) = VAL
                    moduleMatrix(r - 2)(c + 2) = VAL

                    moduleMatrix(r - 1)(c - 2) = VAL
                    moduleMatrix(r - 1)(c - 1) = -VAL
                    moduleMatrix(r - 1)(c + 0) = -VAL
                    moduleMatrix(r - 1)(c + 1) = -VAL
                    moduleMatrix(r - 1)(c + 2) = VAL

                    moduleMatrix(r + 0)(c - 2) = VAL
                    moduleMatrix(r + 0)(c - 1) = -VAL
                    moduleMatrix(r + 0)(c + 0) = VAL
                    moduleMatrix(r + 0)(c + 1) = -VAL
                    moduleMatrix(r + 0)(c + 2) = VAL

                    moduleMatrix(r + 1)(c - 2) = VAL
                    moduleMatrix(r + 1)(c - 1) = -VAL
                    moduleMatrix(r + 1)(c + 0) = -VAL
                    moduleMatrix(r + 1)(c + 1) = -VAL
                    moduleMatrix(r + 1)(c + 2) = VAL

                    moduleMatrix(r + 2)(c - 2) = VAL
                    moduleMatrix(r + 2)(c - 1) = VAL
                    moduleMatrix(r + 2)(c + 0) = VAL
                    moduleMatrix(r + 2)(c + 1) = VAL
                    moduleMatrix(r + 2)(c + 2) = VAL
                End If
            Next
        Next
    End Sub

End Class
