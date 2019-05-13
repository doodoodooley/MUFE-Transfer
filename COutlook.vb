 
Imports Microsoft.Office.Interop.Outlook

Public Class COutlook

    Dim app As Microsoft.Office.Interop.Outlook.Application
    Dim appNameSpace As Microsoft.Office.Interop.Outlook._NameSpace
    Dim msg As Microsoft.Office.Interop.Outlook.MailItem
    Dim booDisplay As Boolean = False

    Public Sub New()

        Try
            app = New Microsoft.Office.Interop.Outlook.Application

            appNameSpace = app.GetNamespace("MAPI")

            appNameSpace.Logon(Nothing, Nothing, False, False)

            msg = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)

        Catch ex As System.Exception

            ' put message in windows event log
            CLog.LogApplication("Error: New " & ex.Message, EventLogEntryType.Error)

            Throw ex

        End Try

    End Sub

    Public Sub Subject(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then
            msg.Subject = strValue
        End If

    End Sub

    Public Sub AddAttachment(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then

            Try
                Dim sBodyLen As String = msg.Body.Length

                Dim oAttachs As Microsoft.Office.Interop.Outlook.Attachments = msg.Attachments

                Dim oAttach As Microsoft.Office.Interop.Outlook.Attachment

                oAttach = oAttachs.Add(strValue, , sBodyLen + 1, strValue)

            Catch ex As System.Exception

                ' put message in windows event log
                CLog.LogApplication("Error: AddAttachment " & ex.Message, EventLogEntryType.Error)

                Throw ex

            End Try

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
            msg.CC = strValue
        End If

    End Sub

    Public Sub ToAddress(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then
            msg.To = strValue
        End If

    End Sub

    Public Sub FromName(ByVal strValue As String)

        If strValue.Trim.Length > 0 Then
            'msg.from = strValue
        End If

    End Sub

    Public Sub SendMail()

        msg.Body = msg.Body & vbCrLf & "Sent: " & Now.ToShortDateString & " " & Now.ToLongTimeString

        Try
            msg.Send()
 

        Catch ex As System.Exception

            ' put message in windows event log
            CLog.LogApplication("Error: SendMail " & ex.Message, EventLogEntryType.Error)

            Throw ex

        End Try

        msg = Nothing
        app = Nothing

    End Sub

End Class
