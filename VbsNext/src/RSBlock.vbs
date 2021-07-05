Class RSBlock_

    Private m_totalNumbers

    Private Sub Class_Initialize()
        m_totalNumbers = Array( _
            Array( 0, _
                   1,  1,  1,  1,  1,  2,  2,  2,  2,  4, _
                   4,  4,  4,  4,  6,  6,  6,  6,  7,  8, _
                   8,  9,  9, 10, 12, 12, 12, 13, 14, 15, _
                  16, 17, 18, 19, 19, 20, 21, 22, 24, 25), _
            Array( 0, _
                   1,  1,  1,  2,  2,  4,  4,  4,  5,  5, _
                   5,  8,  9,  9, 10, 10, 11, 13, 14, 16, _
                  17, 17, 18, 20, 21, 23, 25, 26, 28, 29, _
                  31, 33, 35, 37, 38, 40, 43, 45, 47, 49), _
            Array( 0, _
                   1,  1,  2,  2,  4,  4,  6,  6,  8,  8, _
                   8, 10, 12, 16, 12, 17, 16, 18, 21, 20, _
                  23, 23, 25, 27, 29, 34, 34, 35, 38, 40, _
                  43, 45, 48, 51, 53, 56, 59, 62, 65, 68), _
            Array( 0, _
                   1,  1,  2,  4,  4,  4,  5,  6,  8,  8, _
                  11, 11, 16, 16, 18, 16, 19, 21, 25, 25, _
                  25, 34, 30, 32, 35, 37, 40, 42, 45, 48, _
                  51, 54, 57, 60, 63, 66, 70, 74, 77, 81) _
        )
    End Sub

    Public Function GetTotalNumber(ByVal ecLevel, ByVal ver, ByVal preceding)
        Dim dataWordCapacity
        Dim blockCount

        dataWordCapacity = DataCodeword.GetTotalNumber(ecLevel, ver)
        blockCount = m_totalNumbers(ecLevel)

        If preceding Then
            GetTotalNumber = blockCount(ver) - (dataWordCapacity Mod blockCount(ver))
        Else
            GetTotalNumber = dataWordCapacity Mod blockCount(ver)
        End If
    End Function

    Public Function GetNumberDataCodewords(ByVal ecLevel, ByVal ver, ByVal preceding)
        Dim ret

        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber(ecLevel, ver)

        Dim numBlocks
        numBlocks = m_totalNumbers(ecLevel)(ver)

        Dim numPreBlockCodewords
        numPreBlockCodewords = numDataCodewords \ numBlocks

        Dim numPreBlocks
        Dim numFolBlocks

        If preceding Then
            ret = numPreBlockCodewords
        Else
            numPreBlocks = GetTotalNumber(ecLevel, ver, True)
            numFolBlocks = GetTotalNumber(ecLevel, ver, False)

            If numFolBlocks > 0 Then
                ret = (numDataCodewords - numPreBlockCodewords * numPreBlocks) \ numFolBlocks
            Else
                ret = 0
            End If
        End If

        GetNumberDataCodewords = ret
    End Function

    Public Function GetNumberECCodewords(ByVal ecLevel, ByVal ver)
        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber(ecLevel, ver)

        Dim numBlocks
        numBlocks = m_totalNumbers(ecLevel)(ver)

        GetNumberECCodewords = _
            (Codeword.GetTotalNumber(ver) \ numBlocks) - _
                (numDataCodewords \ numBlocks)
    End Function

End Class


