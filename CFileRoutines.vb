Imports System.IO

Public Class CFileRoutines

    Public Function CountFiles(ByVal strPath As String) As Integer
 
        If Not Directory.Exists(strPath) Then
            Throw New System.Exception("Folder " & strPath & " does not exist")
            Exit Function
        End If

        Dim arrFiles As String() = Directory.GetFiles(strPath)

        CountFiles = arrFiles.GetLength(0)

    End Function

    Public Function FileExists(ByVal strPathFrom As String, ByVal strFile As String) As Boolean

        FileExists = False

        Dim strPathFile As String = strPathFrom & "\" & strFile

        ' check if from folder exists
        If Not Directory.Exists(strPathFrom) Then
            Throw New System.Exception("Folder " & strPathFrom & " does not exist")
            Exit Function
        End If

        ' check if from file exists
        If Not File.Exists(strPathFile) Then
            Exit Function
        End If

        FileExists = True

    End Function

    Public Function FolderExists(ByVal strFolder As String) As Boolean

        FolderExists = False

        ' check if folder exists
        If Not Directory.Exists(strFolder) Then
            Exit Function
        End If

        FolderExists = True

    End Function

    Public Sub CopyFile(ByVal strPathFrom As String, ByVal strPathTo As String, ByVal strFile As String, _
                        Optional ByVal booOverwrite As Boolean = True)

        ' copy a file from 1 folder to another, force overwrite by default

        Dim strPathFileFrom As String = strPathFrom & "\" & strFile
        Dim strPathFileTo As String = strPathTo & "\" & strFile

        ' check if from folder exists
        If Not Directory.Exists(strPathFrom) Then
            Throw New System.Exception("Folder " & strPathFrom & " does not exist")
            Exit Sub
        End If

        ' check if to folder exists
        If Not Directory.Exists(strPathTo) Then
            Throw New System.Exception("Folder " & strPathTo & " does not exist")
            Exit Sub
        End If

        ' check if from file exists
        If Not File.Exists(strPathFileFrom) Then
            Throw New System.Exception("File " & strPathFileFrom & " does not exist")
            Exit Sub
        End If

        ' check to file does not exist if over write false is selected
        If File.Exists(strPathFileTo) And booOverwrite = False Then
            Throw New System.Exception("File " & strPathFileTo & " already exists")
            Exit Sub
        End If

        ' if to file exists and over write true is selected then delete to file
        If File.Exists(strPathFileTo) And booOverwrite = True Then
            Try
                Call DeleteInputFile(strPathTo, strFile)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try
        End If

        ' copy file 
        Try
            File.Copy(strPathFileFrom, strPathFileTo)
        Catch ex As Exception
            Throw ex
            Exit Sub
        End Try

    End Sub


    Public Sub DeleteInputFile(ByVal strPath As String, ByVal strFile As String)

        ' delete a particular file in a folder

        Dim strPathInput As String = strPath & "\" & strFile

        If Not File.Exists(strPathInput) Then
            Throw New System.Exception("File " & strPathInput & " does not exist")
            Exit Sub
        End If

        File.SetAttributes(strPathInput, FileAttributes.Normal)

        Try
            File.Delete(strPathInput)
        Catch ex As Exception
            Throw ex
            Exit Sub
        End Try

    End Sub


    Public Sub RenameFile(ByVal strPathFrom As String, ByVal strFile As String, _
                        ByVal strFileRename As String)

        Dim strPathFileFrom As String = strPathFrom & "\" & strFile
        Dim strPathFileTo As String = strPathFrom & "\" & strFileRename

        ' check if from folder exists
        If Not Directory.Exists(strPathFrom) Then
            Throw New System.Exception("Folder " & strPathFrom & " does not exist")
            Exit Sub
        End If

        ' check if from file exists
        If Not File.Exists(strPathFileFrom) Then
            Throw New System.Exception("File " & strPathFileFrom & " does not exist")
            Exit Sub
        End If

        ' check to file does not exist if over write false is selected
        If File.Exists(strPathFileTo) Then
            Throw New System.Exception("File " & strPathFileTo & " already exists")
            Exit Sub
        End If

        ' if to file exists and over write true is selected then delete to file
        If File.Exists(strPathFileTo) Then
            Try
                Call DeleteInputFile(strPathFrom, strFileRename)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try
        End If

        ' copy file 
        Try
            File.Copy(strPathFileFrom, strPathFileTo)
            File.Delete(strPathFileFrom)
        Catch ex As Exception
            Throw ex
            Exit Sub
        End Try

    End Sub

    Public Sub MoveFile(ByVal strFile As String, ByVal strPathFrom As String, ByVal strPathTo As String, _
                        Optional ByVal booOverwrite As Boolean = True)

        ' move a file from 1 folder to another

        Dim strPathFileFrom As String = strPathFrom & "\" & strFile
        Dim strPathFileTo As String = strPathTo & "\" & strFile

        ' check if from folder exists
        If Not Directory.Exists(strPathFrom) Then
            Throw New System.Exception("Folder " & strPathFrom & " does not exist")
            Exit Sub
        End If

        ' check if to folder exists
        If Not Directory.Exists(strPathTo) Then
            Throw New System.Exception("Folder " & strPathTo & " does not exist")
            Exit Sub
        End If

        ' check if from file exists
        If Not File.Exists(strPathFileFrom) Then
            Throw New System.Exception("File " & strPathFileFrom & " does not exist")
            Exit Sub
        End If

        ' check to file does not exist if over write false is selected
        If File.Exists(strPathFileTo) And booOverwrite = False Then
            Throw New System.Exception("File " & strPathFileTo & " already exists")
            Exit Sub
        End If

        ' if to file exists and over write true is selected then delete to file
        If File.Exists(strPathFileTo) And booOverwrite = True Then
            Try
                Call DeleteInputFile(strPathTo, strFile)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try
        End If

        Try
            File.Move(strPathFileFrom, strPathFileTo)
        Catch ex As Exception
            Throw ex
            Exit Sub
        End Try

    End Sub

  Public Sub CopyRenameFile(ByVal strFile As String, ByVal strPath As String, ByVal strNewFile As String)

    ' rename a file in same folder

    Dim strPathFileOld As String = strPath & "\" & strFile
    Dim strPathFileNew As String = strPath & "\" & strNewFile

    ' check  folder exists
    If Not Directory.Exists(strPath) Then
      Throw New System.Exception("Folder " & strPath & " does not exist")
      Exit Sub
    End If

    ' check if from file exists
    If Not File.Exists(strPathFileOld) Then
      Throw New System.Exception("File " & strPathFileOld & " does not exist")
      Exit Sub
    End If

    ' check to file does not exist  
    If File.Exists(strPathFileNew) Then
      Throw New System.Exception("File " & strPathFileNew & " already exists")
      Exit Sub
    End If

    Try
      File.Move(strPathFileOld, strPathFileNew)
    Catch ex As Exception
      Throw ex
      Exit Sub
    End Try

  End Sub
 
End Class
