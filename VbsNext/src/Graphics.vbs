Include("List")
Include("Point")
Include("BinaryWriter")
Include("BITMAPFILEHEADER")
Include("BITMAPINFOHEADER")
Include("RGBQUAD")

Class Graphics_

    Public Function Build1bppDIB( _
      ByRef bitmapData, ByVal pictWidth, ByVal pictHeight, ByVal foreRgb, ByVal backRgb)
        Dim bfh
        Set bfh = New BITMAPFILEHEADER
        With bfh
            .bfType = &H4D42
            .bfSize = 62 + bitmapData.Size
            .bfReserved1 = 0
            .bfReserved2 = 0
            .bfOffBits = 62
        End With

        Dim bih
        Set bih = New BITMAPINFOHEADER
        With bih
            .biSize = 40
            .biWidth = pictWidth
            .biHeight = pictHeight
            .biPlanes = 1
            .biBitCount = 1
            .biCompression = 0
            .biSizeImage = 0
            .biXPelsPerMeter = 0
            .biYPelsPerMeter = 0
            .biClrUsed = 0
            .biClrImportant = 0
        End With

        Dim palette(1)
        Set palette(0) = New RGBQUAD
        Set palette(1) = New RGBQUAD

        With palette(0)
            .rgbBlue = (foreRgb And &HFF0000) \ 2 ^ 16
            .rgbGreen = (foreRgb And &HFF00&) \ 2 ^ 8
            .rgbRed = foreRgb And &HFF&
            .rgbReserved = 0
        End With

        With palette(1)
            .rgbBlue = (backRgb And &HFF0000) \ 2 ^ 16
            .rgbGreen = (backRgb And &HFF00&) \ 2 ^ 8
            .rgbRed = backRgb And &HFF&
            .rgbReserved = 0
        End With

        Dim ret
        Set ret = New BinaryWriter

        With bfh
            Call ret.Append(.bfType)
            Call ret.Append(.bfSize)
            Call ret.Append(.bfReserved1)
            Call ret.Append(.bfReserved2)
            Call ret.Append(.bfOffBits)
        End With

        With bih
            Call ret.Append(.biSize)
            Call ret.Append(.biWidth)
            Call ret.Append(.biHeight)
            Call ret.Append(.biPlanes)
            Call ret.Append(.biBitCount)
            Call ret.Append(.biCompression)
            Call ret.Append(.biSizeImage)
            Call ret.Append(.biXPelsPerMeter)
            Call ret.Append(.biYPelsPerMeter)
            Call ret.Append(.biClrUsed)
            Call ret.Append(.biClrImportant)
        End With

        With palette(0)
            Call ret.Append(.rgbBlue)
            Call ret.Append(.rgbGreen)
            Call ret.Append(.rgbRed)
            Call ret.Append(.rgbReserved)
        End With

        With palette(1)
            Call ret.Append(.rgbBlue)
            Call ret.Append(.rgbGreen)
            Call ret.Append(.rgbRed)
            Call ret.Append(.rgbReserved)
        End With

        Call bitmapData.CopyTo(ret)

        Set Build1bppDIB = ret
    End Function

    Public Function Build24bppDIB( _
      ByRef bitmapData, ByVal pictWidth, ByVal pictHeight)
        Dim bfh
        Set bfh = New BITMAPFILEHEADER

        With bfh
            .bfType = &H4D42
            .bfSize = 54 + bitmapData.Size
            .bfReserved1 = 0
            .bfReserved2 = 0
            .bfOffBits = 54
        End With

        Dim bih
        Set bih = New BITMAPINFOHEADER

        With bih
            .biSize = 40
            .biWidth = pictWidth
            .biHeight = pictHeight
            .biPlanes = 1
            .biBitCount = 24
            .biCompression = 0
            .biSizeImage = 0
            .biXPelsPerMeter = 0
            .biYPelsPerMeter = 0
            .biClrUsed = 0
            .biClrImportant = 0
        End With

        Dim ret
        Set ret = New BinaryWriter

        With bfh
            Call ret.Append(.bfType)
            Call ret.Append(.bfSize)
            Call ret.Append(.bfReserved1)
            Call ret.Append(.bfReserved2)
            Call ret.Append(.bfOffBits)
        End With

        With bih
            Call ret.Append(.biSize)
            Call ret.Append(.biWidth)
            Call ret.Append(.biHeight)
            Call ret.Append(.biPlanes)
            Call ret.Append(.biBitCount)
            Call ret.Append(.biCompression)
            Call ret.Append(.biSizeImage)
            Call ret.Append(.biXPelsPerMeter)
            Call ret.Append(.biYPelsPerMeter)
            Call ret.Append(.biClrUsed)
            Call ret.Append(.biClrImportant)
        End With

        Call bitmapData.CopyTo(ret)

        Set Build24bppDIB = ret
    End Function

    Public Function FindContours(ByRef img)
        Dim MAX_VALUE
        MAX_VALUE = &H7FFFFFFF

        Dim gpPaths
        Set gpPaths = New List

        Dim st, dr
        Dim x, y
        Dim p
        Dim gpPath

        For y = 0 To UBound(img) - 1
            For x = 0 To UBound(img(y)) - 1
                If Not (img(y)(x) = MAX_VALUE) And _
                    (img(y)(x) > 0 And img(y)(x + 1) <= 0) Then

                    img(y)(x) = MAX_VALUE
                    Set st = New Point
                    Call st.Init(x, y)
                    Set gpPath = New List
                    Call gpPath.Add(st)

                    dr = DIRECTION_UP
                    Set p = st.Clone()
                    p.y = p.y - 1

                    Do
                        Select Case dr
                            Case DIRECTION_UP
                                If img(p.y)(p.x) > 0 Then
                                    img(p.y)(p.x) = MAX_VALUE

                                    If img(p.y)(p.x + 1) <= 0 Then
                                        Set p = p.Clone()
                                        p.y = p.y - 1
                                    Else
                                        Call gpPath.Add(p)
                                        dr = DIRECTION_RIGHT
                                        Set p = p.Clone()
                                        p.x = p.x + 1
                                    End If
                                Else
                                    Set p = p.Clone()
                                    p.y = p.y + 1
                                    Call gpPath.Add(p)

                                    dr = DIRECTION_LEFT
                                    Set p = p.Clone()
                                    p.x = p.x - 1
                                End If

                            Case DIRECTION_DOWN
                                If img(p.y)(p.x) > 0 Then
                                    img(p.y)(p.x) = MAX_VALUE

                                    If img(p.y)(p.x - 1) <= 0 Then
                                        Set p = p.Clone()
                                        p.y = p.y + 1
                                    Else
                                        Call gpPath.Add(p)

                                        dr = DIRECTION_LEFT
                                        Set p = p.Clone()
                                        p.x = p.x - 1
                                    End If
                                Else
                                    Set p = p.Clone()
                                    p.y = p.y - 1
                                    Call gpPath.Add(p)

                                    dr = DIRECTION_RIGHT
                                    Set p = p.Clone()
                                    p.x = p.x + 1
                                End If

                            Case DIRECTION_LEFT
                                If img(p.y)(p.x) > 0 Then
                                    img(p.y)(p.x) = MAX_VALUE

                                    If img(p.y - 1)(p.x) <= 0 Then
                                        Set p = p.Clone()
                                        p.x = p.x - 1
                                    Else
                                        Call gpPath.Add(p)

                                        dr = DIRECTION_UP
                                        Set p = p.Clone()
                                        p.y = p.y - 1
                                    End If
                                Else
                                    Set p = p.Clone()
                                    p.x = p.x + 1
                                    Call gpPath.Add(p)

                                    dr = DIRECTION_DOWN
                                    Set p = p.Clone()
                                    p.y = p.y + 1
                                End If

                            Case DIRECTION_RIGHT
                                If img(p.y)(p.x) > 0 Then
                                    img(p.y)(p.x) = MAX_VALUE

                                    If img(p.y + 1)(p.x) <= 0 Then
                                        Set p = p.Clone()
                                        p.x = p.x + 1
                                    Else
                                        Call gpPath.Add(p)

                                        dr = DIRECTION_DOWN
                                        Set p = p.Clone()
                                        p.y = p.y + 1
                                    End If
                                Else
                                    Set p = p.Clone()
                                    p.x = p.x - 1
                                    Call gpPath.Add(p)

                                    dr = DIRECTION_UP
                                    Set p = p.Clone()
                                    p.y = p.y - 1
                                End If
                            Case Else
                                Call Err.Raise(51)
                        End Select
                    Loop While Not p.Equals(st)

                    Call gpPaths.Add(gpPath.Items())
                End If
            Next
        Next

        FindContours = gpPaths.Items()
    End Function

End Class