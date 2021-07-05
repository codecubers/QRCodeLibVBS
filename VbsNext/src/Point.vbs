Include("Point")

Class Point

    Private m_x
    Private m_y

    Public Sub Init(ByVal x, ByVal y)
        m_x = x
        m_y = y
    End Sub

    Public Property Get x()
        x = m_x
    End Property
    Public Property Let x(ByVal Value)
        m_x = Value
    End Property

    Public Property Get y()
        y = m_y
    End Property
    Public Property Let y(ByVal Value)
        m_y = Value
    End Property

    Public Function Clone()
        Dim ret
        Set ret = New Point
        Call ret.Init(x, y)

        Set Clone = ret
    End Function

    Public Function Equals(ByVal obj)
        Equals = (x = obj.x) And (y = obj.y)
    End Function

End Class

