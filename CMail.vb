Public Class CMail

    Private strEmailAddressTO As String = "mdooley@pcis.lvmh-pc.com"
    Private strEmailAddressCC As String = "mdooley@pcis.lvmh-pc.com"
    Private strEmailAddressFROM As String = "mdooley@pcis.lvmh-pc.com"
    Private strAttachment As String = ""
    Private strSubject As String = "ESSENCE China sFTP Monitor: "
    Private strBody As String = "ESSENCE China sFTP eCommerce Monitoring Exception Report"

    Public Function Send() As Boolean

        Dim booReturn As Boolean = True

        Try

            booReturn = SendEmail()

        Catch ex As Exception

            booReturn = False

        End Try

        Send = booReturn

    End Function

    Public Sub EmailTO(ByVal strValue As String)

        strEmailAddressTO = strValue

    End Sub

    Public Sub EmailCC(ByVal strValue As String)

        strEmailAddressCC = strValue

    End Sub

    Public Sub EmailSubject(ByVal strValue As String)

        strSubject = strSubject & strValue

    End Sub

    Public Sub EmailBody(ByVal strValue As String)

        strBody = strValue

    End Sub

    Public Sub AddEMailAttachment(ByVal strValue As String)

        strAttachment = strValue

    End Sub

    Private Function SendEmail() As Boolean

        SendEmail = True
        Dim objMail As Object
        Dim strEmailType As String = ""

        strEmailType = System.Configuration.ConfigurationManager.AppSettings.Item("EmailEngine")
 
        Try

            Select Case strEmailType
                Case "aspNetEmail"
                    objMail = New CEmail
                Case "Outlook"
                    objMail = New COutlook
                Case Else
                    CLog.LogApplication("Error: SendEmail - Unknown Email Engine", EventLogEntryType.Error)
                    Return False
            End Select

            objMail.FromName(strEmailAddressFROM)
            objMail.ToAddress(strEmailAddressTO)
            objMail.CC(strEmailAddressCC)
            objMail.Subject(strSubject)
            objMail.Body(strBody)

            If strAttachment <> "" Then
                objMail.AddAttachment(strAttachment)
            End If

            objMail.SendMail()

        Catch ex As Exception

            ' put message in windows event log
            CLog.LogApplication("Error: SendEmail " & ex.Message, EventLogEntryType.Error)

            SendEmail = False

        End Try          

        objMail = Nothing

    End Function


End Class
