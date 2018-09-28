Public Class OrphanFinder
    Private mFolderPath As String
    Private mSHandler As New ShortcutHandler
    Public Event FoundParent As EventHandler
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Property FolderPath() As String
        Get
            Return mFolderPath
        End Get
        Set(ByVal value As String)
            mFolderPath = value
        End Set
    End Property

    Private mOrphanList As List(Of String)
    ''' <summary>
    ''' List of potential orphans
    ''' </summary>
    ''' <returns></returns>
    Public Property OrphanList() As List(Of String)
        Get
            Return mOrphanList
        End Get
        Set(ByVal value As List(Of String))
            Dim temp As New List(Of String)
            'Remove live links from the list
            For Each f In value
                If Not LinkTargetExists(f) Then
                    temp.Add(f)
                End If
            Next
            mOrphanList = temp
            FindOrphans()
        End Set
    End Property
    ''' <summary>
    ''' Dictionary containing (new parent,link) pairs
    ''' </summary>
    Private mFoundParents As New Dictionary(Of String, String)
    Public Property FoundParents() As Dictionary(Of String, String)
        Get
            Return mFoundParents
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            mFoundParents = value
        End Set
    End Property
    ''' <summary>
    ''' Searches in subfolders of original destination of each link
    ''' </summary>
    Public Sub FindOrphansOld()
        For Each m In mOrphanList
            'For each directory in LinkTarget's folder
            'Find LinkTarget
            'Check all files until match found

            Dim f As New IO.FileInfo(m)
            Dim FilesToCheck As New List(Of String)
            Dim s As String = LinkTarget(m)
            Dim StartDest = New IO.DirectoryInfo(s)

            While StartDest.Exists = False And StartDest IsNot StartDest.Root
                StartDest = New IO.DirectoryInfo(StartDest.Parent.FullName)

            End While
            If StartDest Is StartDest.Root Then Exit For


            GetAllFilesBelow(StartDest.FullName, FilesToCheck) 'Terrible algorithm

            For Each x In FilesToCheck
                If x.Contains(f.Name.Replace(f.Extension, "")) Then
                    mFoundParents.Add(x, m)
                End If
            Next
        Next
        If mFoundParents.Count <> 0 Then
            RaiseEvent FoundParent(Me, Nothing)
            Reunite()
        End If

    End Sub
    Public Sub FindOrphans()
        'Original file must have been moved to a lower folder or to a different branch
        'First check if it is below. 
        'Find all lower folders, and add to a list
        'If the original folder doesn't exist ignore
        'Repeat this for all orphans.
        'Run through the list for all orphans.
        '


        For Each m In mOrphanList
            'For each directory in LinkTarget's folder
            'Check all files until match found

            '   Dim f As New IO.FileInfo(m)
            'Find LinkTarget
            Dim s As String = LinkTarget(m)
            Dim f As New IO.FileInfo(s)
            Dim StartDest = New IO.DirectoryInfo(f.Directory.FullName)
            'Find higher folder if this one doesn't exist any more
            While StartDest.Exists = False
                StartDest = StartDest.Parent
                Do Until StartDest.Parent.Exists
                    StartDest = StartDest.Parent
                Loop
            End While
            If StartDest Is StartDest.Root Then
                Throw New Exception
            End If
            Dim Found As Boolean = False
            Dim DirsToSearch As New List(Of String)
            Dim i = 0
            FindAllFoldersBelow(StartDest, DirsToSearch, True, True)
            DirsToSearch.Add(StartDest.FullName)

            For Each j In DirsToSearch
                Dim filelist = New IO.DirectoryInfo(j).GetFiles
                For Each x In filelist
                    If x.Name.Equals(f.Name.Replace(f.Extension, "")) Then
                        mFoundParents.Add(x.FullName, m)
                        Found = True
                        Exit For
                    End If

                Next
            Next
        Next
        If mFoundParents.Count <> 0 Then
            RaiseEvent FoundParent(Me, Nothing)
            Reunite()
        End If

    End Sub

    Private Function CheckListForParents(ByVal list As List(Of String), f As IO.FileInfo) As Boolean

        For Each x In list
            If x.Contains(f.Name.Replace(f.Extension, "")) Then
                Return True
                Exit Function
            Else
                Return False
            End If
        Next
        Return False
    End Function

    Public Sub Reunite()
        For Each m In mFoundParents
            Dim f As New IO.FileInfo(m.Key)

            ReAssign_ShortCutPath(m.Key, m.Value)
        Next
    End Sub
End Class
