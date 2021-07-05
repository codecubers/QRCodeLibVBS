Class FormatInfo_

    Private VAL
    Private m_formatInfoValues
    Private m_formatInfoMaskArray

    Private Sub Class_Initialize()
        VAL = FORMAT_INFO
        m_formatInfoValues = Array( _
               &H0&,  &H537&,  &HA6E&,  &HF59&, &H11EB&, &H14DC&, &H1B85&, &H1EB2&, &H23D6&, &H26E1&, _
            &H29B8&, &H2C8F&, &H323D&, &H370A&, &H3853&, &H3D64&, &H429B&, &H47AC&, &H48F5&, &H4DC2&, _
            &H5370&, &H5647&, &H591E&, &H5C29&, &H614D&, &H647A&, &H6B23&, &H6E14&, &H70A6&, &H7591&, _
            &H7AC8&, &H7FFF& _
        )

        m_formatInfoMaskArray = Array(0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 0, 1)
    End Sub

    Public Sub Place(ByVal ecLevel, ByVal maskPatternReference, ByRef moduleMatrix())
        Dim formatInfoValue
        formatInfoValue = GetFormatInfoValue(ecLevel, maskPatternReference)

        Dim temp
        Dim v

        Dim i

        Dim r1
        r1 = 0

        Dim c1
        c1 = UBound(moduleMatrix)

        For i = 0 To 7
            If (formatInfoValue And (2 ^ i)) > 0 Then
                temp = 1 Xor m_formatInfoMaskArray(i)
            Else
                temp = 0 Xor m_formatInfoMaskArray(i)
            End If

            If temp > 0 Then
                v = VAL
            Else
                v = -VAL
            End IF

            moduleMatrix(r1)(8) = v
            moduleMatrix(8)(c1) = v

            r1 = r1 + 1
            c1 = c1 - 1

            If r1 = 6 Then
                r1 = r1 + 1
            End If
        Next

        Dim r2
        r2 = UBound(moduleMatrix) - 6

        Dim c2
        c2 = 7

        For i = 8 To 14
            If (formatInfoValue And (2 ^ i)) > 0 Then
                temp = 1 Xor m_formatInfoMaskArray(i)
            Else
                temp = 0 Xor m_formatInfoMaskArray(i)
            End If

            If temp > 0 Then
                v = VAL
            Else
                v = -VAL
            End IF

            moduleMatrix(r2)(8) = v
            moduleMatrix(8)(c2) = v

            r2 = r2 + 1
            c2 = c2 - 1

            If c2 = 6 Then
                c2 = c2 - 1
            End If
        Next

        moduleMatrix(UBound(moduleMatrix) - 7)(8) = VAL
    End Sub

    Public Sub PlaceTempBlank(ByRef moduleMatrix())
        Dim i
        For i = 0 To 8
            If i <> 6 Then
                moduleMatrix(8)(i) = -VAL
                moduleMatrix(i)(8) = -VAL
            End If
        Next

        Dim numModulesPerSide
        numModulesPerSide = UBound(moduleMatrix) + 1

        For i = UBound(moduleMatrix) - 7 To UBound(moduleMatrix)
            moduleMatrix(8)(i) = -VAL
            moduleMatrix(i)(8) = -VAL
        Next

        moduleMatrix(UBound(moduleMatrix) - 7)(8) = -VAL
    End Sub

    Private Function GetFormatInfoValue(ByVal ecLevel, ByVal maskPatternReference)
        Dim indicator

        Select Case ecLevel
            Case ECR_L
                indicator = 1
            Case ECR_M
                indicator = 0
            Case ECR_Q
                indicator = 3
            Case ECR_H
                indicator = 2
            Case Else
                Call Err.Raise(5)
        End Select

        GetFormatInfoValue = m_formatInfoValues((indicator * 2 ^ 3) Or maskPatternReference)
    End Function

End Class


