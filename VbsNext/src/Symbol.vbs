Include("List")
Include("BitSequence")
Include("BinaryWriter")

Class Symbol

    Private m_parent

    Private m_position

    Private m_currEncoder
    Private m_currEncodingMode
    Private m_currVersion

    Private m_dataBitCapacity
    Private m_dataBitCounter

    Private m_segments
    Private m_segmentCounter

    Private Sub Class_Initialize()
        Set m_segments = New List
        Set m_segmentCounter = CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub Init(ByVal parentObj)
        Set m_parent = parentObj

        m_position = parentObj.Count

        Set m_currEncoder = Nothing
        m_currEncodingMode = MODE_UNKNOWN
        m_currVersion = parentObj.MinVersion

        m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
            parentObj.ErrorCorrectionLevel, parentObj.MinVersion)

        m_dataBitCounter = 0

        Call m_segmentCounter.Add(MODE_NUMERIC, 0)
        Call m_segmentCounter.Add(MODE_ALPHA_NUMERIC, 0)
        Call m_segmentCounter.Add(MODE_BYTE, 0)
        Call m_segmentCounter.Add(MODE_KANJI, 0)

        If parentObj.StructuredAppendAllowed Then
            m_dataBitCapacity = m_dataBitCapacity - STRUCTUREDAPPEND_HEADER_LENGTH
        End If
    End Sub

    Public Property Get Parent()
        Set Parent = m_parent
    End Property

    Public Property Get Version()
        Version = m_currVersion
    End Property

    Public Property Get CurrentEncodingMode()
        CurrentEncodingMode = m_currEncodingMode
    End Property

    Public Function TryAppend(ByVal c)
        Dim bitLength
        bitLength = m_currEncoder.GetCodewordBitLength(c)

        Do While (m_dataBitCapacity < m_dataBitCounter + bitLength)
            If m_currVersion >= m_parent.MaxVersion Then
                TryAppend = False
                Exit Function
            End If

            Call SelectVersion
        Loop

        Call m_currEncoder.Append(c)
        m_dataBitCounter = m_dataBitCounter + bitLength
        Call m_parent.UpdateParity(c)

        TryAppend = True
    End Function

    Public Function TrySetEncodingMode(ByVal encMode, ByVal c)
        Dim encoder
        Set encoder = CreateEncoder(encMode)

        Dim bitLength
        bitLength = encoder.GetCodewordBitLength(c)

        Do While (m_dataBitCapacity < _
                    m_dataBitCounter + _
                    MODEINDICATOR_LENGTH + _
                    CharCountIndicator.GetLength(m_currVersion, encMode) + _
                    bitLength)

            If m_currVersion >= m_parent.MaxVersion Then
                TrySetEncodingMode = False
                Exit Function
            End If

            Call SelectVersion
        Loop

        m_dataBitCounter = m_dataBitCounter + _
                           MODEINDICATOR_LENGTH + _
                           CharCountIndicator.GetLength(m_currVersion, encMode)

        Set m_currEncoder = encoder
        Call m_segments.Add(encoder)
        m_segmentCounter(encMode) = m_segmentCounter(encMode) + 1
        m_currEncodingMode = encMode

        TrySetEncodingMode = True
    End Function

    Private Sub SelectVersion()
        Dim encMode
        Dim num

        For Each encMode In m_segmentCounter.Keys()
            num = m_segmentCounter(encMode)

            m_dataBitCounter = m_dataBitCounter + _
                               num * CharCountIndicator.GetLength( _
                                    m_currVersion + 1, encMode) - _
                               num * CharCountIndicator.GetLength( _
                                    m_currVersion + 0, encMode)
        Next

        m_currVersion = m_currVersion + 1
        m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion)
        m_parent.MinVersion = m_currVersion

        If m_parent.StructuredAppendAllowed Then
            m_dataBitCapacity = m_dataBitCapacity - STRUCTUREDAPPEND_HEADER_LENGTH
        End If
    End Sub

    Private Function BuildDataBlock()
        Dim dataBytes
        dataBytes = GetMessageBytes()

        Dim numPreBlocks
        numPreBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, True)

        Dim numFolBlocks
        numFolBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, False)

        Dim ret()
        ReDim ret(numPreBlocks + numFolBlocks - 1)

        Dim dataIdx
        dataIdx = 0

        Dim numPreBlockDataCodewords
        numPreBlockDataCodewords = RSBlock.GetNumberDataCodewords( _
            m_parent.ErrorCorrectionLevel, m_currVersion, True)

        Dim data()
        Dim i, j

        For i = 0 To numPreBlocks - 1
            ReDim data(numPreBlockDataCodewords - 1)

            For j = 0 To UBound(data)
                data(j) = dataBytes(dataIdx)
                dataIdx = dataIdx + 1
            Next

            ret(i) = data
        Next

        Dim numFolBlockDataCodewords
        numFolBlockDataCodewords = RSBlock.GetNumberDataCodewords( _
            m_parent.ErrorCorrectionLevel, m_currVersion, False)

        For i = numPreBlocks To numPreBlocks + numFolBlocks - 1
            ReDim data(numFolBlockDataCodewords - 1)

            For j = 0 To UBound(data)
                data(j) = dataBytes(dataIdx)
                dataIdx = dataIdx + 1
            Next

            ret(i) = data
        Next

        BuildDataBlock = ret
    End Function

    Private Function BuildErrorCorrectionBlock(ByRef dataBlock())
        Dim i, j

        Dim numECCodewords
        numECCodewords = RSBlock.GetNumberECCodewords( _
            m_parent.ErrorCorrectionLevel, m_currVersion)

        Dim numPreBlocks
        numPreBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, True)

        Dim numFolBlocks
        numFolBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, False)

        Dim ret()
        ReDim ret(numPreBlocks + numFolBlocks - 1)

        Dim eccDataTmp()
        ReDim eccDataTmp(numECCodewords - 1)

        For i = 0 To UBound(ret)
            ret(i) = eccDataTmp
        Next

        Dim gp
        gp = GeneratorPolynomials.Item(numECCodewords)

        Dim eccIdx
        Dim blockIdx
        Dim data()
        Dim exp

        For blockIdx = 0 To UBound(dataBlock)
            ReDim data(UBound(dataBlock(blockIdx)) + UBound(ret(blockIdx)) + 1)
            eccIdx = UBound(data)

            For i = 0 To UBound(dataBlock(blockIdx))
                data(eccIdx) = dataBlock(blockIdx)(i)
                eccIdx = eccIdx - 1
            Next

            For i = UBound(data) To numECCodewords Step -1
                If data(i) > 0 Then
                    exp = GaloisField256.ToExp(data(i))
                    eccIdx = i

                    For j = UBound(gp) To 0 Step -1
                        data(eccIdx) = data(eccIdx) Xor _
                                       GaloisField256.ToInt((gp(j) + exp) Mod 255)
                        eccIdx = eccIdx - 1
                    Next
                End If
            Next

            eccIdx = numECCodewords - 1

            For i = 0 To UBound(ret(blockIdx))
                ret(blockIdx)(i) = data(eccIdx)
                eccIdx = eccIdx - 1
            Next
        Next

        BuildErrorCorrectionBlock = ret
    End Function

    Private Function GetEncodingRegionBytes()
        Dim dataBlock
        dataBlock = BuildDataBlock()

        Dim ecBlock
        ecBlock = BuildErrorCorrectionBlock(dataBlock)

        Dim numCodewords
        numCodewords = Codeword.GetTotalNumber(m_currVersion)

        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion)

        Dim ret()
        ReDim ret(numCodewords - 1)

        Dim r, c

        Dim idx
        idx = 0

        Dim n
        n = 0

        Do While idx < numDataCodewords
            r = n Mod (UBound(dataBlock) + 1)
            c = n \ (UBound(dataBlock) + 1)

            If c <= UBound(dataBlock(r)) Then
                ret(idx) = dataBlock(r)(c)
                idx = idx + 1
            End If

            n = n + 1
        Loop

        n = 0

        Do While idx < numCodewords
            r = n Mod (UBound(ecBlock) + 1)
            c = n \ (UBound(ecBlock) + 1)

            If c <= UBound(ecBlock(r)) Then
                ret(idx) = ecBlock(r)(c)
                idx = idx + 1
            End If

            n = n + 1
        Loop

        GetEncodingRegionBytes = ret
    End Function

    Private Function GetMessageBytes()
        Dim bs
        Set bs = New BitSequence

        If m_parent.Count > 1 Then
            Call WriteStructuredAppendHeader(bs)
        End If

        Call WriteSegments(bs)
        Call WriteTerminator(bs)
        Call WritePaddingBits(bs)
        Call WritePadCodewords(bs)

        GetMessageBytes = bs.GetBytes()
    End Function

    Private Sub WriteStructuredAppendHeader(ByVal bs)
        Call bs.Append(MODEINDICATOR_STRUCTURED_APPEND_VALUE, _
                       MODEINDICATOR_LENGTH)
        Call bs.Append(m_position, _
                       SYMBOLSEQUENCEINDICATOR_POSITION_LENGTH)
        Call bs.Append(m_parent.Count - 1, _
                       SYMBOLSEQUENCEINDICATOR_TOTAL_NUMBER_LENGTH)
        Call bs.Append(m_parent.Parity, _
                       STRUCTUREDAPPEND_PARITY_DATA_LENGTH)
    End Sub

    Private Sub WriteSegments(ByVal bs)
        Dim i
        Dim data
        Dim codewordBitLength

        Dim segment

        For Each segment In m_segments.Items()
            Call bs.Append(segment.ModeIndicator, MODEINDICATOR_LENGTH)
            Call bs.Append(segment.CharCount, _
                           CharCountIndicator.GetLength( _
                                m_currVersion, segment.EncodingMode))

            data = segment.GetBytes()

            For i = 0 To UBound(data) - 1
                Call bs.Append(data(i), 8)
            Next

            codewordBitLength = segment.BitCount Mod 8

            If codewordBitLength = 0 Then
                codewordBitLength = 8
            End If

            Call bs.Append(data(UBound(data)) \ _
                           2 ^ (8 - codewordBitLength), codewordBitLength)
        Next
    End Sub

    Private Sub WriteTerminator(ByVal bs)
        Dim terminatorLength
        terminatorLength = m_dataBitCapacity - m_dataBitCounter

        If terminatorLength > MODEINDICATOR_LENGTH Then
            terminatorLength = MODEINDICATOR_LENGTH
        End If

        Call bs.Append(MODEINDICATOR_TERMINATOR_VALUE, terminatorLength)
    End Sub

    Private Sub WritePaddingBits(ByVal bs)
        If bs.Length Mod 8 > 0 Then
            Call bs.Append(&H0, 8 - (bs.Length Mod 8))
        End If
    End Sub

    Private Sub WritePadCodewords(ByVal bs)
        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion)

        Dim flag
        flag = True

        Dim v

        Do While bs.Length < 8 * numDataCodewords
            If flag Then
                v = 236
            Else
                v = 17
            End If
            Call bs.Append(v, 8)
            flag = Not flag
        Loop
    End Sub

    Private Function GetModuleMatrix()
        Dim numModulesPerSide
        numModulesPerSide = Module.GetNumModulesPerSide(m_currVersion)

        Dim moduleMatrix
        ReDim moduleMatrix(numModulesPerSide - 1)

        Dim i
        Dim cols()

        For i = 0 To UBound(moduleMatrix)
            ReDim cols(numModulesPerSide - 1)
            moduleMatrix(i) = cols
        Next

        Call FinderPattern.Place(moduleMatrix)
        Call Separator.Place(moduleMatrix)
        Call TimingPattern.Place(moduleMatrix)

        If m_currVersion >= 2 Then
            Call AlignmentPattern.Place(m_currVersion, moduleMatrix)
        End If

        Call FormatInfo.PlaceTempBlank(moduleMatrix)

        If m_currVersion >= 7 Then
            Call VersionInfo.PlaceTempBlank(moduleMatrix)
        End If

        Call PlaceSymbolChar(moduleMatrix)
        Call RemainderBit.Place(moduleMatrix)

        Call Masking.Apply(m_currVersion, m_parent.ErrorCorrectionLevel, moduleMatrix)

        GetModuleMatrix = moduleMatrix
    End Function

    Private Sub PlaceSymbolChar(ByRef moduleMatrix())
        Dim data
        data = GetEncodingRegionBytes()

        Dim r
        r = UBound(moduleMatrix)

        Dim c
        c = UBound(moduleMatrix(0))

        Dim toLeft
        toLeft = True

        Dim rowDirection
        rowDirection = -1

        Dim bitPos
        Dim v

        For Each v In data
            bitPos = 7

            Do While bitPos >= 0
                If moduleMatrix(r)(c) = BLANK Then
                    If (v And 2 ^ bitPos) > 0 Then
                        moduleMatrix(r)(c) = WORD
                    Else
                        moduleMatrix(r)(c) = -WORD
                    End If

                    bitPos = bitPos - 1
                End If

                If toLeft Then
                    c = c - 1
                Else
                    If (r + rowDirection) < 0 Then
                        r = 0
                        rowDirection = 1
                        c = c - 1

                        If c = 6 Then
                            c = 5
                        End If

                    ElseIf ((r + rowDirection) > UBound(moduleMatrix)) Then
                        r = UBound(moduleMatrix)
                        rowDirection = -1
                        c = c - 1

                        If c = 6 Then
                            c = 5
                        End If

                    Else
                        r = r + rowDirection
                        c = c + 1
                    End If
                End If

                toLeft = Not toLeft
            Loop
        Next
    End Sub

    Private Function GetBitmap1bpp(ByVal moduleSize, ByVal foreColor, ByVal backColor)
        Dim foreRgb
        foreRgb = ColorCode.ToRGB(foreColor)
        Dim backRgb
        backRgb = ColorCode.ToRGB(backColor)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        Dim moduleMatrix
        moduleMatrix = QuietZone.Place(GetModuleMatrix())

        Dim moduleCount
        moduleCount = UBound(moduleMatrix) + 1

        Dim pictWidth
        pictWidth = moduleCount * moduleSize

        Dim pictHeight
        pictHeight = moduleCount * moduleSize

        Dim rowBytesLen
        rowBytesLen = (pictWidth + 7) \ 8

        Dim pack8bit
        If pictWidth Mod 8 > 0 Then
            pack8bit = 8 - (pictWidth Mod 8)
        End If

        Dim pack32bit
        If rowBytesLen Mod 4 > 0 Then
            pack32bit = 8 * (4 - (rowBytesLen Mod 4))
        End If

        Dim rowSize
        rowSize = (pictWidth + pack8bit + pack32bit) \ 8

        Dim bitmapData
        Set bitmapData = New BinaryWriter

        Dim offset
        offset = 0

        Dim bs
        Set bs = New BitSequence

        Dim r
        Dim i
        Dim v
        Dim pixelColor
        Dim bitmapRow

        For r = UBound(moduleMatrix) To 0 Step -1
            Call bs.Clear

            For Each v In moduleMatrix(r)
                If IsDark(v) Then
                    pixelColor = 0
                Else
                    pixelColor = 1
                End If

                For i = 1 To moduleSize
                    Call bs.Append(pixelColor, 1)
                Next
            Next

            Call bs.Append(0, pack8bit)
            Call bs.Append(0, pack32bit)

            bitmapRow = bs.GetBytes()

            For i = 1 To moduleSize
                Call bitmapData.Append(bitmapRow)
            Next
        Next

        Dim ret
        Set ret = Graphics.Build1bppDIB(bitmapData, pictWidth, pictHeight, foreRgb, backRgb)

        Set GetBitmap1bpp = ret
    End Function

    Private Function GetBitmap24bpp(ByVal moduleSize, ByVal foreColor, ByVal backColor)
        Dim foreRgb
        foreRgb = ColorCode.ToRGB(foreColor)
        Dim backRgb
        backRgb = ColorCode.ToRGB(backColor)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        Dim moduleMatrix
        moduleMatrix = QuietZone.Place(GetModuleMatrix())

        Dim pictWidth
        pictWidth = (UBound(moduleMatrix) + 1) * moduleSize

        Dim pictHeight
        pictHeight = pictWidth

        Dim rowBytesLen
        rowBytesLen = 3 * pictWidth

        Dim pack4byte
        If rowBytesLen Mod 4 > 0 Then
            pack4byte = 4 - (rowBytesLen Mod 4)
        End If

        Dim rowSize
        rowSize = rowBytesLen + pack4byte

        Dim bitmapData
        Set bitmapData = New BinaryWriter

        Dim offset
        offset = 0

        Dim r
        Dim i
        Dim v

        Dim colorRGB
        Dim bitmapRow()
        Dim idx

        For r = UBound(moduleMatrix) To 0 Step -1
            ReDim bitmapRow(rowSize - 1)
            idx = 0

            For Each v In moduleMatrix(r)
                If IsDark(v) Then
                    colorRGB = foreRgb
                Else
                    colorRGB = backRgb
                End If

                For i = 1 To moduleSize
                    bitmapRow(idx + 0) = CByte((colorRGB And &HFF0000) \ 2 ^ 16)
                    bitmapRow(idx + 1) = CByte((colorRGB And &HFF00&) \ 2 ^ 8)
                    bitmapRow(idx + 2) = CByte(colorRGB And &HFF&)
                    idx = idx + 3
                Next
            Next

            For i = 1 To pack4byte
                bitmapRow(idx) = CByte(0)
                idx = idx + 1
            Next

            For i = 1 To moduleSize
                Call bitmapData.Append(bitmapRow)
            Next
        Next

        Dim ret
        Set ret = Graphics.Build24bppDIB(bitmapData, pictWidth, pictHeight)

        Set GetBitmap24bpp = ret
    End Function

    Public Sub Save1bppDIB(ByVal filePath, ByVal moduleSize, ByVal foreRgb, ByVal backRgb)
        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        If Len(filePath) = 0 Then Call Err.Raise(5)
        If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
        If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)
        If ColorCode.IsWebColor(backRgb) = False Then Call Err.Raise(5)

        Dim dib
        Set dib = GetBitmap1bpp(moduleSize, foreRgb, backRgb)

        Call dib.SaveToFile(filePath, adSaveCreateOverWrite)
    End Sub

    Public Sub Save24bppDIB(ByVal filePath, ByVal moduleSize, ByVal foreRgb, ByVal backRgb)
        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        If Len(filePath) = 0 Then Call Err.Raise(5)
        If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
        If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)
        If ColorCode.IsWebColor(backRgb) = False Then Call Err.Raise(5)

        Dim dib
        Set dib = GetBitmap24bpp(moduleSize, foreRgb, backRgb)

        Call dib.SaveToFile(filePath, adSaveCreateOverWrite)
    End Sub

    Public Sub SaveSvg(ByVal filePath, ByVal moduleSize, ByVal foreRgb)
        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        If Len(filePath) = 0 Then Call Err.Raise(5)
        If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
        If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)

        Dim svg
        svg = GetSvg(moduleSize, foreRgb)
        Dim svgFile
        svgFile = _
            "<?xml version='1.0' encoding='UTF-8' standalone='no'?>" & vbNewLine & _
            "<!DOCTYPE svg PUBLIC '-//W3C//DTD SVG 20010904//EN'" & vbNewLine & _
            "    'http://www.w3.org/TR/2001/REC-SVG-20010904/DTD/svg10.dtd'>" & vbNewLine & _
            svg

        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim ts
        Set ts = fso.CreateTextFile(filePath, True)
        Call ts.WriteLine(svgFile)
        ts.Close
    End Sub

    Public Function GetSvg(ByVal moduleSize, ByVal foreRgb)
        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        If moduleSize < MIN_MODULE_SIZE Then Call Err.Raise(5)
        If ColorCode.IsWebColor(foreRgb) = False Then Call Err.Raise(5)

        Dim moduleMatrix
        moduleMatrix = QuietZone.Place(GetModuleMatrix())

        Dim imageWidth
        imageWidth = (UBound(moduleMatrix) + 1) * moduleSize

        Dim imageHeight
        imageHeight = imageWidth

        Dim img()
        ReDim img(imageHeight - 1)

        Dim imgRow()
        Dim r, c
        Dim i, j
        Dim v

        r = 0
        Dim rowArray
        For Each rowArray In moduleMatrix
            ReDim imgRow(imageWidth - 1)
            c = 0
            For Each v In rowArray
                For j = 1 To moduleSize
                    If IsDark(v) Then
                        imgRow(c) = 1
                    Else
                        imgRow(c) = 0
                    End If

                    c = c + 1
                Next
            Next

            For i = 1 To moduleSize
                img(r) = imgRow
                r = r + 1
            Next
        Next

        Dim gpPaths
        gpPaths = Graphics.FindContours(img)

        Dim buf
        Set buf = New List

        Dim indent
        indent = String(11, " ")

        Dim gpPath
        Dim k

        For Each gpPath In gpPaths
            Call buf.Add(indent & "M ")

            For k = 0 To UBound(gpPath)
                Call buf.Add(CStr(gpPath(k).x) & "," & CStr(gpPath(k).y) & " ")
            Next
            Call buf.Add("Z" & vbNewLine)
        Next

        Dim data
        data = Trim(Join(buf.Items(), ""))
        data = Left(data, Len(data) - Len(vbNewLine))
        Dim svg
        svg = _
            "<svg xmlns='http://www.w3.org/2000/svg'" & vbNewLine & _
            "    width='" & CStr(imageWidth) & "' height='" & CStr(imageHeight) & "' viewBox='0 0 " & CStr(imageWidth) & " " & CStr(imageHeight) & "'>" & vbNewLine & _
            "    <path fill='" & foreRgb & "' stroke='" & foreRgb & "' stroke-width='1'" & vbNewLine & _
            "        d='" & data & "'" & vbNewLine & _
            "    />" & vbNewLine & _
            "</svg>"

        GetSvg = svg
    End Function

End Class