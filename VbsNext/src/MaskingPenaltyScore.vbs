Class MaskingPenaltyScore_

    Public Function CalcTotal(ByRef moduleMatrix())
        Dim total
        Dim penalty

        penalty = CalcAdjacentModulesInSameColor(moduleMatrix)
        total = total + penalty

        penalty = CalcBlockOfModulesInSameColor(moduleMatrix)
        total = total + penalty

        penalty = CalcModuleRatio(moduleMatrix)
        total = total + penalty

        penalty = CalcProportionOfDarkModules(moduleMatrix)
        total = total + penalty

        CalcTotal = total
    End Function

    Private Function CalcAdjacentModulesInSameColor(ByRef moduleMatrix())
        Dim penalty
        penalty = 0

        penalty = penalty + CalcAdjacentModulesInRowInSameColor(moduleMatrix)
        penalty = penalty + CalcAdjacentModulesInRowInSameColor(MatrixRotate90(moduleMatrix))

        CalcAdjacentModulesInSameColor = penalty
    End Function

    Private Function CalcAdjacentModulesInRowInSameColor(ByRef moduleMatrix())
        Dim penalty
        penalty = 0

        Dim rowArray
        Dim i
        Dim cnt

        For Each rowArray In moduleMatrix
            cnt = 1

            For i = 0 To UBound(rowArray) - 1
                If IsDark(rowArray(i)) = IsDark(rowArray(i + 1)) Then
                    cnt = cnt + 1
                Else
                    If cnt >= 5 Then
                        penalty = penalty + (3 + (cnt - 5))
                    End If

                    cnt = 1
                End If
            Next

            If cnt >= 5 Then
                penalty = penalty + (3 + (cnt - 5))
            End If
        Next

        CalcAdjacentModulesInRowInSameColor = penalty
    End Function

    Private Function CalcBlockOfModulesInSameColor(ByRef moduleMatrix())
        Dim penalty
        Dim r, c
        Dim temp

        For r = 0 To UBound(moduleMatrix) - 1
            For c = 0 To UBound(moduleMatrix(r)) - 1
                temp = IsDark(moduleMatrix(r)(c))

                If (IsDark(moduleMatrix(r + 0)(c + 1)) = temp) And _
                   (IsDark(moduleMatrix(r + 1)(c + 0)) = temp) And _
                   (IsDark(moduleMatrix(r + 1)(c + 1)) = temp) Then
                    penalty = penalty + 3
                End If
            Next
        Next

        CalcBlockOfModulesInSameColor = penalty
    End Function

    Private Function CalcModuleRatio(ByRef moduleMatrix())
        Dim moduleMatrixTemp
        moduleMatrixTemp = QuietZone.Place(moduleMatrix)

        Dim penalty
        penalty = 0

        penalty = penalty + CalcModuleRatioInRow(moduleMatrixTemp)
        penalty = penalty + CalcModuleRatioInRow(MatrixRotate90(moduleMatrixTemp))

        CalcModuleRatio = penalty
    End Function

    Private Function CalcModuleRatioInRow(ByRef moduleMatrix())
        Dim penalty

        Dim ratio3Ranges
        Dim rowArray

        Dim ratio1, ratio3, ratio4

        Dim i
        Dim cnt
        Dim flg
        Dim impose

        Dim rng

        For Each rowArray In moduleMatrix
            ratio3Ranges = GetRatio3Ranges(rowArray)

            For Each rng In ratio3Ranges
                ratio3 = rng(1) + 1 - rng(0)
                ratio1 = ratio3 \ 3
                ratio4 = ratio1 * 4
                flg = True
                impose = False

                i = rng(0) - 1

                If flg Then
                    ' light ratio 1
                    cnt = 0
                    Do While i >= 0
                        If Not IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i - 1
                        Else
                            Exit Do
                        End If
                    Loop

                    flg = cnt = ratio1
                End If

                If flg Then
                    ' dark ratio 1
                    cnt = 0
                    Do While i >= 0
                        If IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i - 1
                        Else
                            Exit Do
                        End If
                    Loop

                    flg = cnt = ratio1
                End If

                If flg Then
                    ' light ratio 4
                    cnt = 0
                    Do While i >= 0
                        If Not IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i - 1
                        Else
                            Exit Do
                        End If
                    Loop

                    If cnt >= ratio4 Then
                        impose = True
                    End If
                End If

                i = rng(1) + 1

                If flg Then
                    ' light ratio 1
                    cnt = 0
                    Do While i <= UBound(rowArray)
                        If Not IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i + 1
                        Else
                            Exit Do
                        End If
                    Loop

                    flg = cnt = ratio1
                End If

                If flg Then
                    ' dark ratio 1
                    cnt = 0
                    Do While i <= UBound(rowArray)
                        If IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i + 1
                        Else
                            Exit Do
                        End If
                    Loop

                    flg = cnt = ratio1
                End If

                If flg Then
                    ' light ratio 4
                    cnt = 0
                    Do While i <= UBound(rowArray)
                        If Not IsDark(rowArray(i)) Then
                            cnt = cnt + 1
                            i = i + 1
                        Else
                            Exit Do
                        End If
                    Loop

                    If cnt >= ratio4 Then
                        impose = True
                    End If
                End If

                If flg And impose Then
                    penalty = penalty + 40
                End If
            Next
        Next

        CalcModuleRatioInRow = penalty
    End Function

    Private Function GetRatio3Ranges(ByRef arg)
        Dim ret
        ret = Array()

        Dim s, i

        For i = 1 To UBound(arg) - 1
            If IsDark(arg(i)) Then
                If Not IsDark(arg(i - 1)) Then
                    s = i
                End If

                If Not IsDark(arg(i + 1)) Then
                    If (i + 1 - s) Mod 3 = 0 Then
                        ReDim Preserve ret(UBound(ret) + 1)
                        ret(UBound(ret)) = Array(s, i)
                    End If
                End If
            End If
        Next

        GetRatio3Ranges = ret
    End Function

    Private Function CalcProportionOfDarkModules(ByRef moduleMatrix())
        Dim darkCount

        Dim rowArray
        Dim v

        For Each rowArray In moduleMatrix
            For Each v In rowArray
                If IsDark(v) Then
                    darkCount = darkCount + 1
                End If
            Next
        Next

        Dim numModules
        numModules = (UBound(moduleMatrix) + 1) ^ 2

        Dim k
        k = darkCount / numModules * 100
        k = Abs(k - 50)
        k = Int(k / 5)
        Dim penalty
        penalty = CInt(k) * 10

        CalcProportionOfDarkModules = penalty
    End Function

    Private Function MatrixRotate90(ByRef arg())
        Dim ret()
        ReDim ret(UBound(arg(0)))

        Dim i, j
        Dim cols()

        For i = 0 To UBound(ret)
            ReDim cols(UBound(arg))
            ret(i) = cols
        Next

        Dim k
        k = UBound(ret)

        For i = 0 To UBound(ret)
            For j = 0 To UBound(ret(i))
                ret(i)(j) = arg(j)(k - i)
            Next
        Next

        MatrixRotate90 = ret
    End Function

End Class
