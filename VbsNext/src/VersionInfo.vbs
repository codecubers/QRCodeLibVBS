Class VersionInfo_

    Private m_versionInfoValues

    Private Sub Class_Initialize()
        m_versionInfoValues = Array( _
            -1, -1, -1, -1, -1, -1, -1, _
            &H7C94&, &H85BC&, &H9A99&, &HA4D3&, &HBBF6&, &HC762&, &HD847&, &HE60D&, _
            &HF928&, &H10B78, &H1145D, &H12A17, &H13532, &H149A6, &H15683, &H168C9, _
            &H177EC, &H18EC4, &H191E1, &H1AFAB, &H1B08E, &H1CC1A, &H1D33F, &H1ED75, _
            &H1F250, &H209D5, &H216F0, &H228BA, &H2379F, &H24B0B, &H2542E, &H26A64, _
            &H27541, &H28C69 _
        )
    End Sub

    Public Sub Place(ByVal ver, ByRef moduleMatrix())
        Dim numModulesPerSide
        numModulesPerSide = UBound(moduleMatrix) + 1

        Dim versionInfoValue
        versionInfoValue = m_versionInfoValues(ver)

        Dim p1
        p1 = 0

        Dim p2
        p2 = numModulesPerSide - 11

        Dim i
        Dim v

        For i = 0 To 17
            If (versionInfoValue And 2 ^ i) > 0 Then
                v = VERSION_INFO
            Else
                v = -VERSION_INFO
            End IF

            moduleMatrix(p1)(p2) = v
            moduleMatrix(p2)(p1) = v

            p2 = p2 + 1

            If i Mod 3 = 2 Then
                p1 = p1 + 1
                p2 = numModulesPerSide - 11
            End If
        Next
    End Sub

    Public Sub PlaceTempBlank(ByRef moduleMatrix())
        Dim numModulesPerSide
        numModulesPerSide = UBound(moduleMatrix) + 1

        Dim i, j

        For i = 0 To 5
            For j = numModulesPerSide - 11 To numModulesPerSide - 9
                moduleMatrix(i)(j) = -VERSION_INFO
                moduleMatrix(j)(i) = -VERSION_INFO
            Next
        Next
    End Sub

End Class

