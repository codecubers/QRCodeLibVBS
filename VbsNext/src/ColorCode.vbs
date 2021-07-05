Class ColorCode_

    Public Property Get BLACK()
        BLACK = "#000000"
    End Property

    Public Property Get WHITE()
        WHITE = "#FFFFFF"
    End Property

    Public Function IsWebColor(arg)
        Dim re
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^#[0-9A-Fa-f]{6}$"
        Dim ret
        ret = re.Test(arg)
        IsWebColor = ret
    End Function

    Public Function ToRGB(ByVal arg)
        If Not IsWebColor(arg) Then Call Err.Raise(5)

        Dim ret
        ret = RGB(CInt("&h" & Mid(arg, 2, 2)), _
                  CInt("&h" & Mid(arg, 4, 2)), _
                  CInt("&h" & Mid(arg, 6, 2)))

        ToRGB = ret
    End Function

End Class
