Imports System.IO

Imports System.Xml

Public Class ResetPassword
    Dim oldPass As String, newPass As String, confirmPass As String


    Private Sub btnPassword_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPassword.Click
        Dim sResult As String
        oldPass = TextBox1.Text.Trim()
        newPass = TextBox2.Text.Trim()
        confirmPass = TextBox3.Text.Trim()

        'all fields are required
        If oldPass.Length = 0 Then
            MsgBox("Old password is required")
            TextBox3.Focus()
            Exit Sub
        End If

        If newPass.Trim().Length = 0 Then
            MsgBox("New Password is required")
            TextBox2.Focus()
            Exit Sub
        End If

        If confirmPass.Length = 0 Then
            MsgBox("Confirm Password is required")
            TextBox3.Focus()
            Exit Sub
        End If

        ' old password and new password must not be same
        If oldPass.Equals(newPass) Then
            MsgBox("Old password and new password cannot be same")
            TextBox1.Focus()
            Exit Sub
        End If
        ' new password and confirm password must be same
        ' old password and new password must not be same
        If Not newPass.Equals(confirmPass) Then
            MsgBox("New password and confirm password must be same")
            TextBox2.Focus()
            Exit Sub
        End If

        Jse.NewPassword = NewPassword


        Dim sApiOtions As String
        sApiOtions = "<CONNECTIONOPTIONS><PASSWORD Old=""%old"" New=""%new""></PASSWORD></CONNECTIONOPTIONS>"
        'call api to change password()
        sApiOtions = sApiOtions.Replace("%old", Jse.OldPassword)
        sApiOtions = sApiOtions.Replace("%new", Jse.NewPassword)


        Try
            'mApi = New GV8APILib.GlobalVision8API()
            'call setAp
            sResult = mApi.SetAPIOptions(sApiOtions)

            'make call via api
            'sResult = mApi.LoginGUI()

            Dim doc As XmlReader = XmlReader.Create(New StringReader(sResult))
            Dim code As String, Msg As String, result As String


            While doc.Read()
                If doc.Name = "REPLY" Then
                    code = doc("Code")
                    Msg = doc("Message")
                    result = doc("result")

                    If code <> "0" Then
                        MsgBox("Connection Failed : Code - " & code & Environment.NewLine & _
                               " Message - " & Msg & Environment.NewLine & _
                               result)
                        Jse.CodeFromApi = doc("Code")
                        Jse.MessageFromApi = doc("Message")

                    Else
                        ' SetApiOptionsSucceeded    
                        Try
                            Dim sresult1 As String
                            sResult1 = mApi.Login(Jse.server, Jse.port, Jse.username, Jse.OldPassword)

                            Dim doc1 As XmlReader = XmlReader.Create(New StringReader(sResult1))

                            While doc1.Read()
                                If doc1.Name = "REPLY" Then
                                    code = doc1("Code")
                                    Msg = doc1("Message")
                                    result = doc1("Result")

                                    If code <> "0" Then
                                        MsgBox("Connection Failed : Code - " & code & Environment.NewLine & _
                                                " Message - " & Msg & _
                                                Environment.NewLine & result)
                                        Jse.CodeFromApi = doc("Code")

                                    Else
                                        Jse.connected = "True"
                                    End If

                                End If
                            End While
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try

                    End If

                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' close window
        Me.Close()

    End Sub



End Class