
Class Codeword_

    Private m_totalNumbers

    Private Sub Class_Initialize()
        m_totalNumbers = Array( _
              -1, _
              26,   44,   70,  100,  134,  172,  196,  242,  292,  346, _
             404,  466,  532,  581,  655,  733,  815,  901,  991, 1085, _
            1156, 1258, 1364, 1474, 1588, 1706, 1828, 1921, 2051, 2185, _
            2323, 2465, 2611, 2761, 2876, 3034, 3196, 3362, 3532, 3706 _
        )
    End Sub

    Public Function GetTotalNumber(ByVal ver)
        GetTotalNumber = m_totalNumbers(ver)
    End Function

End Class



