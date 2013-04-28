Imports System.IO

Imports System.Xml
Imports NLog

Public Class Login
    Private logger As NLog.Logger = NLog.LogManager.GetCurrentClassLogger()



    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Dim sResult As String = ""
        Dim code As String = ""
        Dim Msg As String = ""
        Dim result As String = ""
        Dim username As String, password As String

        'validate that server name,port,username and password are present
        If TextBox3.Text.Trim().Length = 0 Then
            MsgBox("User name is required")
            TextBox3.Focus()
            Exit Sub
        End If

        If TextBox4.Text.Trim().Length = 0 Then
            MsgBox("Password is required")
            TextBox4.Focus()
            Exit Sub
        End If

        Jse.username = TextBox3.Text.Trim()
        Jse.OldPassword = TextBox4.Text.Trim()

        Try
            mApi = New GV8APILib.GlobalVision8API()
            sResult = mApi.Login(Jse.server, Jse.port, Jse.username, Jse.OldPassword)

            'make call via api
            'sResult = mApi.LoginGUI()

            logger.Info(sResult)

            Dim doc As XmlReader = XmlReader.Create(New StringReader(sResult))

            While doc.Read()
                If doc.Name = "REPLY" Then
                    code = doc("Code")
                    Msg = doc("Message")
                    result = doc("Result")

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

        If code = "303" Then
            ' call reset password form
            Dim ResetPass As New ResetPassword()
            ResetPass.ShowDialog()
            ' get values and make api call to reset password
        End If


        Me.Close()

    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = Jse.server
        TextBox2.Text = Jse.port

    End Sub
End Class