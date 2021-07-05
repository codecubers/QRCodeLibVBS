
Class RGBQUAD

    Private m_rgbBlue
    Public Property Let rgbBlue(ByVal Value)
        m_rgbBlue = CByte(Value)
    End Property
    Public Property Get rgbBlue()
        rgbBlue = m_rgbBlue
    End Property

    Private m_rgbGreen
    Public Property Let rgbGreen(ByVal Value)
        m_rgbGreen = CByte(Value)
    End Property
    Public Property Get rgbGreen()
        rgbGreen = m_rgbGreen
    End Property

    Private m_rgbRed
    Public Property Let rgbRed(ByVal Value)
        m_rgbRed = CByte(Value)
    End Property
    Public Property Get rgbRed()
        rgbRed = m_rgbRed
    End Property

    Private m_rgbReserved
    Public Property Let rgbReserved(ByVal Value)
        m_rgbReserved = CByte(Value)
    End Property
    Public Property Get rgbReserved()
        rgbReserved = m_rgbReserved
    End Property

End Class