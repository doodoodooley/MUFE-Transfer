Imports aspNetEmail

Public Class CEmail

    Private msg As EmailMessage

    Public Sub New()
  
        Dim strSMTPServer As String = GetConnectionString()
 
        msg = New EmailMessage()

        msg.Server = strSMTPServer
        msg.Port = 25

        ' load the default configurations
        msg.LoadFromConfig()

    End Sub

    Public Sub Subject(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then
            msg.Subject = strValue
        End If

    End Sub

    Public Sub AddAttachment(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then
            msg.AddAttachment(strValue)
        End If

    End Sub

    Public Sub Body(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then
            msg.Body &= strValue
        End If

    End Sub

    Public Sub BCC(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then
            msg.Bcc = strValue
        End If

    End Sub

    Public Sub CC(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then
            msg.Cc = strValue
        End If

    End Sub

    Public Sub ToAddress(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then
            msg.To = strValue
        End If

    End Sub

    Public Sub FromName(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then
            msg.FromName = strValue
        End If

    End Sub

    Public Sub SendMail()

        msg.Body = msg.Body & vbCrLf & "Sent: " & Now.ToShortDateString & " " & Now.ToLongTimeString

        Try
            msg.Send()

        Catch ex As Exception

            ' put message in windows event log
            CLog.LogApplication("Error: SendMail " & ex.Message, EventLogEntryType.Error)

            Throw ex

        End Try

    End Sub

    Private Function GetConnectionString() As String

        Dim strServer As String = "smtpserver"

        GetConnectionString = System.Configuration.ConfigurationManager.AppSettings.Item(strServer)

    End Function

End Class
