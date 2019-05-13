Imports System.IO

Public Class CFTP

    ' global variable
    Private bTest As Boolean

    ' global class
    Dim cChina As New FTPClass

    Dim strFTPSyncExt As String = ""
    Dim strFTPDataExt As String = ""
    Dim strAxSyncExt As String = ""
    Dim strAxDataExt As String = ""

    Public Function Main() As Boolean

        Dim booReturn As Boolean = True

        Dim strAxImportPath = GetAppConfigValue("AxImportFolder")
          
        Try
          
            ' check FTP Connection
            booReturn = FTPConnection()
             
            ' connected then carry on
            If booReturn = True Then

                ' clear work folder
                booReturn = ClearWorkFolder()

                ' get files waiting at FTP site and bring locally
                booReturn = ProcessFTPFiles(strAxImportPath)
                  
            End If

        Catch ex As Exception
             
            booReturn = False

        End Try

        booReturn = FTPDisconnect()

        Return booReturn

    End Function
     
    Private Function ClearWorkFolder() As Boolean

        Dim booReturn As Boolean = True
        Dim strWork As String = GetAppConfigValue("WorkFolder")
        Dim objWork As New DirectoryInfo(strWork)
        Dim objFile As FileInfo
        Dim objFileRoutines As New CFileRoutines

        Try

            For Each objFile In objWork.GetFiles

                ' delete file from ax work folder
                objFileRoutines.DeleteInputFile(strWork, Path.GetFileName(objFile.Name))

            Next

        Catch ex As Exception
             
            booReturn = False

        End Try

        Return booReturn

    End Function

    Private Function FTPConnection() As Boolean

        ' this routine needs to report the error

        Dim booReturn As Boolean = False

        Try

            If Not cChina.Connected Then

                ' FTP site settings
                cChina.Hostname = GetAppConfigValue("FTP-Hostname")
                cChina.Username = GetAppConfigValue("FTP-Username")
                cChina.Password = GetAppConfigValue("FTP-Password")
                cChina.Port = GetAppConfigValue("FTP-Port")

                If cChina.Connect() Then
                    booReturn = True
                Else
                     
                End If

            Else

                booReturn = True

            End If

        Catch ex As Exception
             
            booReturn = False

        End Try

        Return booReturn

    End Function

    Private Function FTPDisconnect() As Boolean

        ' this routine needs to report the error

        Dim booReturn As Boolean = False

        Try

            If cChina.Disconnect Then
                booReturn = True
            Else
               
            End If

        Catch ex As Exception
             
            booReturn = False

        End Try

        Return booReturn

    End Function

    Private Function FTPUpload(ByVal strFullFileName As String, ByVal strFileName As String) As Boolean

        Dim booReturn As Boolean = True
        
        Try

            ' report in the log
            'CLog.LogApplication("Routine: FTPUpload " & strFullFileName & " ~ " & strFileName, EventLogEntryType.Information)

            If FTPConnection() = True Then

                Call FTPDisconnect()
                Call FTPConnection()

                cChina.CurrentDirectory = GetAppConfigValue("FTP-DepositFolder")
                cChina.ChangeDirectory()

                cChina.Upload(strFullFileName, strFileName)

            End If

        Catch ex As Exception
             
            booReturn = False

        End Try

        Return booReturn

    End Function

    Private Function FTPDownload(strFTPFileName As String, ByVal strAxFileName As String) As Boolean

        Dim booReturn As Boolean = False
        Dim strWorkPath = GetAppConfigValue("WorkFolder")

        Try

            If FTPConnection() = True Then

                'Call FTPDisconnect()
                'Call FTPConnection()

                'cChina.CurrentDirectory = GetAppConfigValue("FTP-CollectFolder")
                'cChina.ChangeDirectory()

                ' ignore 
                If strFTPFileName <> ".ko" And strFTPFileName <> ".csv" And strFTPFileName <> ".ssh" Then
                    cChina.Download(strFTPFileName, strWorkPath & "\" & strAxFileName)
                End If

            End If

        Catch ex As Exception

            booReturn = False

        End Try

        Return booReturn

    End Function

    Private Function FTPDelete(strFullFileName) As Boolean

        Dim booReturn As Boolean = False

        Try

            ' report in the log
            'CLog.LogApplication("Routine: FTPDelete " & strFullFileName, EventLogEntryType.Information)

            If FTPConnection() = True Then

                'Call FTPDisconnect()
                'Call FTPConnection()

                'cChina.CurrentDirectory = GetAppConfigValue("FTP-CollectFolder")
                'cChina.ChangeDirectory()

                cChina.Delete(strFullFileName)

            End If

        Catch ex As Exception
             
            booReturn = False

        End Try

        Return booReturn

    End Function

    Private Function CheckFileCount() As Integer

        Dim intCount As Integer = cChina.DirectoryCount()

        Return intCount

    End Function

    Private Function CheckFileOutCount(strFolder As String) As Integer

        cChina.CurrentDirectory = strFolder

        Dim intCount As Integer = cChina.DirectoryCount()

        Return intCount

    End Function

    Private Function ProcessFTPFiles(strImportPath) As Boolean

        Dim booReturn As Boolean = True

        Try

            ' reset the current folder on the FTP Server
            'Call FTPDisconnect()
            'Call FTPConnection()
            'cChina.CurrentDirectory = GetAppConfigValue("FTP-CollectFolder")
            'cChina.ChangeDirectory()

            ' check if any files waiting at eCommerce site           
            Dim intCount As Integer = CheckFileCount()

            Select Case intCount
                Case Is = 0 ' no files waiting
                    ' nothing to do - report in the log
                    'CLog.LogApplication("Routine: ProcessFTPFiles " & "No Files Waiting.", EventLogEntryType.Information)
                Case Is > 0 ' files waiting
                    ' process waiting files - report in the log
                    'CLog.LogApplication("Routine: ProcessFTPFiles " & intCount & " Files waiting to be processed.", EventLogEntryType.Information)
                    booReturn = GetFTPFiles(strImportPath)
                Case Is < 0 ' error condition
                    ' report in the log
                    'CLog.LogApplication("Error: ProcessFTPFiles " & "Error Condition - File Count is less than zero.", EventLogEntryType.Error)
                    booReturn = False
            End Select

        Catch ex As Exception
             
            booReturn = False

        End Try

        Return booReturn

    End Function
 

    Private Function GetFTPFiles(strImportPath As String) As Boolean

        Dim objFileRoutines As New CFileRoutines
        Dim booReturn As Boolean = True

        Dim strWorkPath = GetAppConfigValue("WorkFolder")
        Dim strArchivePath = GetAppConfigValue("ArchiveFolder")

        Try

            ' loop thru list of files at FTP site
            Dim FTPList As New List(Of String)

            FTPList = cChina.DirectoryList()

            For Each strFile As String In FTPList

                ' copy the file to work folder
                Call FTPDownload(strFile, strFile)

                ' delete the  file
                Call FTPDelete(strFile)

            Next

            ' archive files and move to import folder
            Dim objAxRoot As New DirectoryInfo(strWorkPath)

            ' archive the files on Ax 
            For Each objFile In objAxRoot.GetFiles

                ' copy file from work folder to archive folder
                objFileRoutines.CopyFile(strWorkPath, strArchivePath, objFile.Name, True)

            Next

            ' move to import folder  
            For Each objFile In objAxRoot.GetFiles

                ' move .data file from ax work folder to Ax import folder if not blank file 

                ' removing this since non empty files seem to be getting deleted 
                'If IsEmptyFile(strWorkPath & "\" & objFile.Name) Then
                ' delete from work
                'objFileRoutines.DeleteInputFile(strWorkPath, objFile.Name)
                'Else
                If objFile.Name <> ".ssh" Then
                    objFileRoutines.MoveFile(Path.GetFileName(strWorkPath & "\" & objFile.Name), strWorkPath, strImportPath)
                End If


            Next

        Catch ex As Exception
             
            booReturn = False

        End Try

        Return booReturn

    End Function

    Private Function IsEmptyFile(strFile As String) As Boolean

        Dim booReturn As Boolean = False
        Dim sr As StreamReader = New StreamReader(strFile)
        Dim line As String

        Try

            Do
                line = sr.ReadLine()

                If line = String.Empty Then
                    booReturn = True
                End If

            Loop Until line Is Nothing Or booReturn = True

            sr.Close()

        Catch ex As Exception

            booReturn = False

        End Try

        Return booReturn

    End Function

    Private Function GetAppConfigValue(strKey As String) As String

        Dim strValue As String = ""

        Try

            strValue = ("" & System.Configuration.ConfigurationManager.AppSettings.Item(strKey)).Trim

        Catch ex As Exception
             
        End Try

        Return strValue

    End Function


End Class
