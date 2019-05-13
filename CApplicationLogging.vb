Public Class CApplicationLogging

    Dim bEventLog As Boolean
    Dim bEmailLog As Boolean

    Const strLogName As String = "ESSENCE"
    Const strApplication As String = "China-eCommerce"

    Private Sub ClearLog()

        '' only clear if its saturday 
        'If LCase(Format(Now, "dddddd")) <> "saturday" Then
        '    Exit Sub
        'End If

        'Try

        '    Dim objEventLog As New CEventLog

        '    ' temp Stop CLEARING LOG due to strange issues
        '    ' objEventLog.ClearLog(strLogName)

        '    ' log event
        '    objEventLog.WriteToEventLog("Clearing Log", strApplication, EventLogEntryType.SuccessAudit, strLogName)

        'Catch ex As Exception

        '    'don't try to log an error

        'End Try

    End Sub

    Public Sub LogApplication(ByVal strMessage As String, ByVal intLogType As Integer)
         
        strMessage = "China eCommerce OMS Monitor: " & vbCrLf & strMessage

        If bEventLog = True Then

            Try

                Dim objEventLog As New CEventLog

                objEventLog.WriteToEventLog(strMessage, strApplication, intLogType, strLogName)

            Catch ex As Exception

                'don't try to log an error

            End Try

        End If

        ' only send email if event log type is error
        If intLogType = EventLogEntryType.Error Then

            ' and email logging is on
            If bEmailLog = True Then

                Try

                    Dim strCC As String = GetAppConfigValue("FolderCC")

                    Call SendByMail(strApplication, intLogType & " ~ " & strMessage, strCC)

                Catch ex As Exception

                    'don't try to log an error

                End Try

            End If

        End If

    End Sub

    Public Sub New()

        Try

            ' clear the log every saturday
            'Call ClearLog()

            ' get event logging on flag
            bEventLog = CBool(GetAppConfigValue("EventLoggingOn"))

            ' get email logging on flag
            bEmailLog = CBool(GetAppConfigValue("EmailLoggingOn"))

            Dim strTimestamp As String = Format(Now, "MMM-dd-yyyy hh:mm tt")

            ' always log run 
            CLog.LogApplication("Running China OMS FTP Interface: " & strTimestamp, EventLogEntryType.Information)

        Catch ex As Exception

            'don't try to log an error

        End Try

    End Sub

    Private Function GetAppConfigValue(strKey As String) As String

        Dim strValue As String = ""

        Try

            strValue = ("" & System.Configuration.ConfigurationManager.AppSettings.Item(strKey)).Trim

        Catch ex As Exception

            ' put message in windows event log
            CLog.LogApplication("Error: GetAppConfigValue " & ex.Message, EventLogEntryType.Error)

        End Try

        Return strValue

    End Function

    Private Function SendByMail(strSubject As String, strBody As String, strCC As String) As Boolean

        Dim CMail As New CMail
        Dim booReturn As Boolean = True

        ' get list of CC people  
        strCC = System.Configuration.ConfigurationManager.AppSettings.Item(strCC)

        Try

            CMail.EmailSubject(strSubject)

            If strCC <> "" Then
                CMail.EmailCC(strCC)
            End If

            CMail.EmailBody(strBody)

            CMail.Send()

        Catch ex As Exception

            ' put message in windows event log
            CLog.LogApplication("Error: SendByMail " & ex.Message, EventLogEntryType.Error)

            booReturn = False

        Finally

            CMail = Nothing

        End Try

        Return booReturn

    End Function

End Class
