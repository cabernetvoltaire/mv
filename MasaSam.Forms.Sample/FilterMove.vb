''' <summary>
''' Given a folder, move all files whose names contain the name of a subfolder to that subfolder
''' </summary>

Class FilterMove
    Public Event FilesMoved(sender As Object, e As EventArgs)
    Private mFolder As String
    Private mFolders As New List(Of IO.DirectoryInfo)
    Public Property Folder() As String
        Get
            Return mFolder
        End Get
        Set(ByVal value As String)
            mFolder = value
        End Set
    End Property

    Public Sub New()
        mRecursive = False
    End Sub

    Private mRecursive As Boolean
    Public Property Recursive() As Boolean
        Get
            Return mRecursive
        End Get
        Set(ByVal value As Boolean)
            mRecursive = value
        End Set
    End Property
    Public Sub FilterMoveFiles(foldername As String)
        GetFoldersBelow(foldername)
        Dim s As New IO.DirectoryInfo(foldername)
        For Each f In s.GetFiles
            For Each d In mFolders
                If InStr(f.Name, d.Name) <> 0 And Len(d.Name) > 1 Then
                    Try
                        My.Computer.FileSystem.MoveFile(f.FullName, d.FullName & "\" & f.Name)
                        s.Refresh()
                    Catch ex As Exception
                        ' MsgBox("While moving " & f.FullName & vbCrLf & "to " & d.FullName & vbCrLf & ex.Message)
                        Exit For
                    End Try
                End If
            Next
        Next
        RaiseEvent FilesMoved(Nothing, Nothing)
    End Sub
    Private Sub GetFoldersBelow(folderpath As String)

        Dim s As New IO.DirectoryInfo(folderpath)
        For Each m In s.GetDirectories
            If mRecursive Then GetFoldersBelow(m.FullName)
            mFolders.Add(m)
        Next

    End Sub
End Class
