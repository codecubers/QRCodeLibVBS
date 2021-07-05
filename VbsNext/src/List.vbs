Class List

    Private m_items

    Private Sub Class_Initialize()
        m_items = Array()
    End Sub

    Public Sub Add(arg)
        ReDim Preserve m_items(UBound(m_items) + 1)

        If VarType(arg) = vbObject Then
            Set m_items(UBound(m_items)) = arg
        Else
            m_items(UBound(m_items)) = arg
        End If
    End Sub

    Public Property Get Count()
        Count = UBound(m_items) + 1
    End Property

    Public Property Get Item(ByVal idx)
        If VarType(m_items(idx)) = vbObject Then
            Set Item = m_items(idx)
        Else
            Item = m_items(idx)
        End If
    End Property

    Public Property Get Items()
        Items = m_items
    End Property

End Class

