Class BinaryWriter

    Private m_byteTable(255)
    Private m_stream

    Private Sub Class_Initialize()
        Call MakeByteTable

        Set m_stream = CreateObject("ADODB.Stream")
        m_stream.Type = adTypeBinary
        Call m_stream.Open
    End Sub

    Public Property Get Stream()
        Set Stream = m_stream
    End Property

    Public Property Get Size()
        Size = m_stream.Size
    End Property

    Private Sub MakeByteTable()
        Dim sr
        Set sr = CreateObject("ADODB.Stream")
        sr.Type = adTypeText
        sr.Charset = "unicode"
        Call sr.Open

        Dim i
        For i = 0 To 255
            sr.WriteText ChrW(i)
        Next

        sr.Position = 0
        sr.Type = adTypeBinary
        sr.Position = 2

        For i = 0 To 255
            m_byteTable(i) = sr.Read(1)
            sr.Position = sr.Position + 1
        Next

        Call sr.Close
    End Sub

    Public Sub Append(ByVal arg)
        If (VarType(arg) And vbArray) = 0 Then
            arg = Array(arg)
        End If

        Dim temp
        Dim v

        For Each v In arg
            Select Case VarType(v)
                Case vbByte
                    Call m_stream.Write(m_byteTable(v))
                Case vbInteger
                    temp = v And &HFF&
                    Call m_stream.Write(m_byteTable(temp))

                    temp = (v And &HFF00&) \ 2 ^ 8
                    Call m_stream.Write(m_byteTable(temp))
                Case vbLong
                    temp = v And &HFF&
                    Call m_stream.Write(m_byteTable(temp))

                    temp = (v And &HFF00&) \ 2 ^ 8
                    Call m_stream.Write(m_byteTable(temp))

                    temp = (v And &HFF0000) \ 2 ^ 16
                    Call m_stream.Write(m_byteTable(temp))

                    temp = (v And &HFF000000) \ 2 ^ 24
                    Call m_stream.Write(m_byteTable(temp))
                Case Else
                    Call Err.Raise(5)
            End Select
        Next
    End Sub

    Public Sub CopyTo(ByVal destBinaryWriter)
        m_stream.Position = 0
        Call m_stream.CopyTo(destBinaryWriter.Stream)
    End Sub

    Public Sub SaveToFile(ByVal FileName, ByVal SaveOptions)
        Call m_stream.SaveToFile(FileName, SaveOptions)
    End Sub

End Class