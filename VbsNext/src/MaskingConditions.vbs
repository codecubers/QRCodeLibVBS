
Class MaskingCondition0

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = (r + c) Mod 2 = 0
    End Function

End Class


Class MaskingCondition1

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = r Mod 2 = 0
    End Function

End Class


Class MaskingCondition2

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = c Mod 3 = 0
    End Function

End Class


Class MaskingCondition3

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = (r + c) Mod 3 = 0
    End Function

End Class


Class MaskingCondition4

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r \ 2) + (c \ 3)) Mod 2 = 0
    End Function

End Class


Class MaskingCondition5

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r * c) Mod 2 + (r * c) Mod 3) = 0
    End Function

End Class


Class MaskingCondition6

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r * c) Mod 2 + (r * c) Mod 3) Mod 2 = 0
    End Function

End Class


Class MaskingCondition7

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r + c) Mod 2 + (r * c) Mod 3) Mod 2 = 0
    End Function

End Class
