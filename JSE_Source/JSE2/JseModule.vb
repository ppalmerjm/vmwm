Imports GV8APILib

Module Jse
    Public server As String = "10.100.100.56"
    Public port As String = "11997"

    Public username As String = ""
    Public OldPassword As String = ""
    Public NewPassword As String = ""

    Public mApi As GV8APILib.GlobalVision8API
    Public CodeFromApi As String
    Public MessageFromApi As String

    Public connected As String = "False"
    Public PasswordChanged As String = "False"

End Module
