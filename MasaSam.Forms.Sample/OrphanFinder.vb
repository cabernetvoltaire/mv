Public Class OrphanFinder
    Private mFolderPath As String
    Private mSHandler As ShortcutHandler
    Public Event FoundParent As EventHandler

    Public Property FolderPath() As String
        Get
            Return mFolderPath
        End Get
        Set(ByVal value As String)
            mFolderPath = value
        End Set
    End Property

    Private mOrphanList As List(Of String)
    Public Property OrphanList() As List(Of String)
        Get
            Return mOrphanList
        End Get
        Set(ByVal value As List(Of String))
            mOrphanList = value
            FindOrphans()
        End Set
    End Property

    Private mFoundParents As New Dictionary(Of String, String)
    Public Property FoundParents() As Dictionary(Of String, String)
        Get
            Return mFoundParents
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            mFoundParents = value
        End Set
    End Property

    Public Sub FindOrphans()
        For Each m In mOrphanList
            Dim f As New IO.FileInfo(m)
            Dim FilesToCheck As New List(Of String)
            Dim StartDest = New IO.DirectoryInfo(CreateObject("WScript.Shell").CreateShortcut(m).TargetPath)
            While StartDest.Exists = False And StartDest IsNot StartDest.Root
                StartDest = New IO.DirectoryInfo(StartDest.Parent.FullName)

            End While
            If StartDest Is StartDest.Root Then Exit For

            GetAllFilesBelow(StartDest.FullName, FilesToCheck)

            For Each x In FilesToCheck
                If x.Contains(f.Name.Replace(f.Extension, "")) Then
                    mFoundParents.Add(x, m)
                    RaiseEvent FoundParent(Me, Nothing)
                End If
            Next
        Next
        '    Reunite()
    End Sub

    Private Sub Reunite()
        For Each m In mFoundParents
            mSHandler.ShortcutPath = m.Key
            mSHandler.TargetPath = m.Value
            mSHandler.Create_ShortCut()

        Next
    End Sub
End Class
