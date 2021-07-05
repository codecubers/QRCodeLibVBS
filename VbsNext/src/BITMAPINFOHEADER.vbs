Class BITMAPINFOHEADER

    Private m_biSize
    Public Property Let biSize(ByVal Value)
        m_biSize = CLng(Value)
    End Property
    Public Property Get biSize()
        biSize = m_biSize
    End Property

    Private m_biWidth
    Public Property Let biWidth(ByVal Value)
        m_biWidth = CLng(Value)
    End Property
    Public Property Get biWidth()
        biWidth = m_biWidth
    End Property

    Private m_biHeight
    Public Property Let biHeight(ByVal Value)
        m_biHeight = CLng(Value)
    End Property
    Public Property Get biHeight()
        biHeight = m_biHeight
    End Property

    Private m_biPlanes
    Public Property Let biPlanes(ByVal Value)
        m_biPlanes = CInt(Value)
    End Property
    Public Property Get biPlanes()
        biPlanes = m_biPlanes
    End Property

    Private m_biBitCount
    Public Property Let biBitCount(ByVal Value)
        m_biBitCount = CInt(Value)
    End Property
    Public Property Get biBitCount()
        biBitCount = m_biBitCount
    End Property

    Private m_biCompression
    Public Property Let biCompression(ByVal Value)
        m_biCompression = CLng(Value)
    End Property
    Public Property Get biCompression()
        biCompression = m_biCompression
    End Property

    Private m_biSizeImage
    Public Property Let biSizeImage(ByVal Value)
        m_biSizeImage = CLng(Value)
    End Property
    Public Property Get biSizeImage()
        biSizeImage = m_biSizeImage
    End Property

    Private m_biXPelsPerMeter
    Public Property Let biXPelsPerMeter(ByVal Value)
        m_biXPelsPerMeter = CLng(Value)
    End Property
    Public Property Get biXPelsPerMeter()
        biXPelsPerMeter = m_biXPelsPerMeter
    End Property

    Private m_biYPelsPerMeter
    Public Property Let biYPelsPerMeter(ByVal Value)
        m_biYPelsPerMeter = CLng(Value)
    End Property
    Public Property Get biYPelsPerMeter()
        biYPelsPerMeter = m_biYPelsPerMeter
    End Property

    Private m_biClrUsed
    Public Property Let biClrUsed(ByVal Value)
        m_biClrUsed = CLng(Value)
    End Property
    Public Property Get biClrUsed()
        biClrUsed = m_biClrUsed
    End Property

    Private m_biClrImportant
    Public Property Let biClrImportant(ByVal Value)
        m_biClrImportant = CLng(Value)
    End Property
    Public Property Get biClrImportant()
        biClrImportant = m_biClrImportant
    End Property

End Class