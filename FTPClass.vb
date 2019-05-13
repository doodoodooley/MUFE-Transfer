Imports rebex.net
 
Public Class FTPClass
 
    Private fHostname As String = ""
    Private fUsername As String = ""
    Private fPassword As String = ""
    Private fPort As Integer = 0
    Private fConnectionError As String = ""
    Private fDirectory As String = ""

    Private client As New Sftp

    Public Property CurrentDirectory() As String
        Get
            Return fDirectory
        End Get
        Set(ByVal Value As String)
            fDirectory = Value
        End Set
    End Property

    Public Property Hostname() As String
        Get
            Return fHostname
        End Get
        Set(ByVal Value As String)
            fHostname = Value
        End Set
    End Property

    Public Property ConnectionError() As String
        Get
            Return fConnectionError
        End Get
        Set(ByVal Value As String)
            fConnectionError = Value
        End Set
    End Property
 
    Public Property Username() As String
        Get
            Return fUsername
        End Get
        Set(ByVal Value As String)
            fUsername = Value
        End Set
    End Property
     
    Public Property Password() As String
        Get
            Return fPassword
        End Get
        Set(ByVal Value As String)
            fPassword = Value
        End Set
    End Property
     
    Public Property Port() As Integer
        Get
            Return fPort
        End Get
        Set(ByVal Value As Integer)
            fPort = Value
        End Set
    End Property

    Public Function ChangeDirectory() As Boolean

        Dim booReturn As Boolean = False

        If client.IsConnected = False Then
            ConnectionError = "FTP Client Not Connected"
            Return booReturn
        End If

        If client.IsAuthenticated = False Then
            ConnectionError = "FTP Client Not Authenticated"
            Return booReturn
        End If

        If fDirectory = "" Then
            ConnectionError = "FTP Current Directory Not Set"
            Return booReturn
        End If

        Try
            client.ChangeDirectory(fDirectory)
            booReturn = True

        Catch ex As Exception

            ConnectionError = ex.Message

        End Try

        Return booReturn

    End Function

    Public Function Connected() As Boolean

        Try

            If client.IsConnected = True Then
                Connected = True
            Else
                Connected = False
            End If

        Catch ex As Exception

            ConnectionError = ex.Message
            Throw New System.Exception("An exception has occurred checking Connection. " & ConnectionError)
            
        End Try

        Return Connected

    End Function

    Public Function Connect() As Boolean

        Try

            client.Connect(fHostname, fPort)

            client.Login(fUsername, fPassword)

            ChangeDirectory()

            Return True

        Catch ex As Exception

            ConnectionError = ex.Message

            Return False

        End Try

    End Function

    Public Function Disconnect() As Boolean

        Try

            If client.IsConnected = True Then
                client.Disconnect()
            End If

            Return True

        Catch ex As Exception

            ConnectionError = ex.Message

            Return False

        End Try

    End Function

    Public Function DirectoryCount() As Integer

        Try

            Dim list As SftpItemCollection = client.GetList()

            Return list.Count

        Catch ex As Exception

            ConnectionError = ex.Message

            Throw New System.Exception("An exception has occurred. " & ConnectionError)

        End Try

    End Function

    Public Function DirectoryList() As List(Of String)

        Dim lsResult As New List(Of String)

        Try

            Dim list As SftpItemCollection = client.GetList()

            Dim item As SftpItem

            For Each item In list

                lsResult.Add(item.Name)

            Next item

        Catch ex As Exception

            ConnectionError = ex.Message

        End Try

        Return lsResult

    End Function

    Public Function Download(ByVal RemoteFilename As String, ByVal LocalFilename As String) As Boolean

        Try

            client.GetFile(RemoteFilename, LocalFilename)

            Return True

        Catch ex As Exception

            ConnectionError = ex.Message

            Return False

        End Try

    End Function

    Public Function Upload(ByVal LocalFilename As String, ByVal RemoteFilename As String) As Boolean

        Try

            client.PutFile(LocalFilename, RemoteFilename)

            Return True

        Catch ex As Exception

            ConnectionError = ex.Message

            Throw New System.Exception("An exception has occurred. " & ConnectionError)

        End Try

    End Function

    Public Function Delete(ByVal RemoteFilename As String) As Boolean

        Try

            client.Delete(RemoteFilename, Rebex.IO.TraversalMode.MatchFilesShallow)

            Return True

        Catch ex As Exception

            ConnectionError = ex.Message

            Return False

        End Try

    End Function

End Class

