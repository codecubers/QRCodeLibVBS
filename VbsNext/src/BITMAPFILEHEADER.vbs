Class BITMAPFILEHEADER

    Private m_bfType
    Public Property Let bfType(ByVal Value)
        m_bfType = CInt(Value)
    End Property
    Public Property Get bfType()
        bfType = m_bfType
    End Property

    Private m_bfSize
    Public Property Let bfSize(ByVal Value)
        m_bfSize = CLng(Value)
    End Property
    Public Property Get bfSize()
        bfSize = m_bfSize
    End Property

    Private m_bfReserved1
    Public Property Let bfReserved1(ByVal Value)
        m_bfReserved1 = CInt(Value)
    End Property
    Public Property Get bfReserved1()
        bfReserved1 = m_bfReserved1
    End Property

    Private m_bfReserved2
    Public Property Let bfReserved2(ByVal Value)
        m_bfReserved2 = CInt(Value)
    End Property
    Public Property Get bfReserved2()
        bfReserved2 = m_bfReserved2
    End Property

    Private m_bfOffBits
    Public Property Let bfOffBits(ByVal Value)
        m_bfOffBits = CLng(Value)
    End Property
    Public Property Get bfOffBits()
        bfOffBits = m_bfOffBits
    End Property

End Class
