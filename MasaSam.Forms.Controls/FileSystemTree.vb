﻿Imports System.ComponentModel
Imports System.IO

Public Class FileSystemTree

#Region "Members"

    Private m_showSystem As Boolean = False
    Private m_showHidden As Boolean = False
    Private m_showNetworkDrives As Boolean = True
    Private m_showRemovableDrives As Boolean = True
    Private m_showRamDrives As Boolean = True
    Private m_showInactiveDrives As Boolean = True
    Private m_unauthorizedAccessBehavior As UnauthorizedAccessBehavior = UnauthorizedAccessBehavior.Hide
    Private m_driveReadTimeout As Int32 = 5
    Private m_extensions As String = "*"
    Private m_showImages As Boolean = False
    Private m_trackDriveState As Boolean = False
    Private m_inactiveDrives As New List(Of String)
    Private m_initialTreeCreate As Boolean = False
    Private m_timer As Timer
    Private m_selectedNodes As New Dictionary(Of String, FileSystemNode)
    Private m_rootDrive As String
    Private m_currentNode As FileSystemNode 'pac


#End Region

#Region "Events"

    <Description("Notifies when drive has been actived.")>
    <Category("Behavior")>
    Public Event DriveActivated As EventHandler(Of DriveInfoEventArgs)

    <Description("Notifies when drive has been unactivated.")>
    <Category("Behavior")>
    Public Event DriveUnactivated As EventHandler(Of DriveInfoEventArgs)

    <Description("Notifies when drive is selected either by clicking or expanding.")>
    <Category("Behavior")>
    Public Event DriveSelected As EventHandler(Of DriveInfoEventArgs)

    <Description("Notifies when directory is selected either by clicking or expanding.")>
    <Category("Behavior")>
    Public Event DirectorySelected As EventHandler(Of DirectoryInfoEventArgs)

    <Description("Notifies when file is selected either by clicking or expanding.")>
    <Category("Behavior")>
    Public Event FileSelected As EventHandler(Of FileInfoEventArgs)
    Public Event LabelEditBegun As EventHandler(Of String)

#End Region

#Region "Properties"
    Private mAutoExpandNodes As Boolean
    Public Property AutoExpandNodes() As Boolean
        Get
            Return mAutoExpandNodes
        End Get
        Set(ByVal value As Boolean)
            mAutoExpandNodes = value
        End Set
    End Property
    Private mShowFiles As Boolean = False
    <DefaultValue(False)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets whether or not files and directories with hidden attribute should be shown.")>
    Public Property ShowHidden As Boolean
        Get
            Return Me.m_showHidden
        End Get
        Set(value As Boolean)
            Me.m_showHidden = value
        End Set
    End Property
    Private newTraverseKeyBack As Keys = Keys.Subtract
    <DefaultValue(Keys.Subtract)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets the Key which traverses the tree.")>
    Public Property TraverseKeyBack() As Keys
        Get
            Return newTraverseKeyBack
        End Get
        Set(ByVal value As Keys)
            newTraverseKeyBack = value
        End Set
    End Property

    Private newTraverseKey As Keys = Keys.Add
    <DefaultValue(Keys.Add)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets the Key which traverses the tree.")>
    Public Property TraverseKey() As Keys
        Get
            Return newTraverseKey
        End Get
        Set(ByVal value As Keys)
            newTraverseKey = value
        End Set
    End Property

    <DefaultValue(False)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets whether to show files in the tree.")>
    Public Property ShowFiles() As Boolean
        Get
            Return mShowFiles
        End Get
        Set(ByVal value As Boolean)
            mShowFiles = value
        End Set
    End Property


    <DefaultValue(False)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets whether or not files and directories with system attribute should be shown.")>
    Public Property ShowSystem As Boolean
        Get
            Return Me.m_showSystem
        End Get
        Set(value As Boolean)
            Me.m_showSystem = value
        End Set
    End Property

    <DefaultValue(True)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets whether or not network drives should be shown.")>
    Public Property ShowNetworkDrives As Boolean
        Get
            Return Me.m_showNetworkDrives
        End Get
        Set(value As Boolean)
            Me.m_showNetworkDrives = value
        End Set
    End Property

    <DefaultValue(True)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets whether or not removable drives should be shown.")>
    Public Property ShowRemovableDrives As Boolean
        Get
            Return Me.m_showRemovableDrives
        End Get
        Set(value As Boolean)
            Me.m_showRemovableDrives = value
        End Set
    End Property

    <DefaultValue(True)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets whether or not RAM drives should be shown.")>
    Public Property ShowRamDrives As Boolean
        Get
            Return Me.m_showRamDrives
        End Get
        Set(value As Boolean)
            Me.m_showRamDrives = value
        End Set
    End Property

    <DefaultValue(True)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets whether or not drives that are not ready should be shown.")>
    Public Property ShowInactiveDrives As Boolean
        Get
            Return Me.m_showInactiveDrives
        End Get
        Set(value As Boolean)
            Me.m_showInactiveDrives = value
        End Set
    End Property

    <DefaultValue(5)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets timeout in seconds for how long should wait for drive to become ready.")>
    Public Property DriveReadTimeout As Int32
        Get
            Return Me.m_driveReadTimeout
        End Get
        Set(value As Int32)
            If (value < 1 Or value > 60) Then
                Throw New ArgumentOutOfRangeException("value", value, "The drive read timeout must be between 1 and 60 seconds.")
            End If
            Me.m_driveReadTimeout = value
        End Set
    End Property

    <DefaultValue("*.*")>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Get or sets allowed file extensions. For example value '.txt;.xml' shows only files with txt or xml extension.")>
    Public Property FileExtensions As String
        Get
            Return Me.m_extensions
        End Get
        Set(value As String)
            If (String.IsNullOrWhiteSpace(value)) Then
                Me.m_extensions = "*.*"
            Else
                Me.m_extensions = value
            End If
        End Set
    End Property

    <DefaultValue(GetType(UnauthorizedAccessBehavior), "Hide")>
    <Category("Behavior")>
    <Browsable(True)>
    Public Property UnauthorizedAccessBehavior As UnauthorizedAccessBehavior
        Get
            Return Me.m_unauthorizedAccessBehavior
        End Get
        Set(value As UnauthorizedAccessBehavior)
            Me.m_unauthorizedAccessBehavior = value
        End Set
    End Property

    <Browsable(False)>
    Public ReadOnly Property HideUnauthorizedDirectories As Boolean
        Get
            Return UnauthorizedAccessBehavior = Forms.Controls.UnauthorizedAccessBehavior.Hide
        End Get
    End Property

    <DefaultValue(False)>
    <Category("Appearance")>
    <Browsable(True)>
    <Description("Gets or sets whether or not imager are shown in tree.")>
    Public Property ShowImages As Boolean
        Get
            Return Me.m_showImages
        End Get
        Set(value As Boolean)
            Me.m_showImages = value
        End Set
    End Property

    <DefaultValue(False)>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets whether or not control tracks state of drives.")>
    Public Property TrackDriveState As Boolean
        Get
            Return Me.m_trackDriveState
        End Get
        Set(value As Boolean)
            Me.m_trackDriveState = value

            If (Me.m_trackDriveState) Then
                EnableDriveStateCheck()
            Else
                DisableDriveStateCheck()
            End If

        End Set
    End Property

    <DefaultValue("")>
    <Category("Behavior")>
    <Browsable(True)>
    <Description("Gets or sets the name of the root drive. If empty or Nothing, then My Computer is root.")>
    Public Property RootDrive As String
        Get
            Return Me.m_rootDrive
        End Get
        Set(value As String)
            If (value <> Me.m_rootDrive) Then
                Me.m_rootDrive = Path.GetPathRoot(value)
                If (m_rootDrive.EndsWith(":")) Then
                    m_rootDrive += "\"
                End If
                CreateInitialTree()
            End If
        End Set
    End Property


    Private newSelectedFolder As String
    Public Property SelectedFolder() As String
        Get
            Return newSelectedFolder
        End Get
        Set(ByVal value As String)
            If value = "" Then Exit Property
            Dim d As New DirectoryInfo(value)
            newSelectedFolder = value
            ClearSelectedNodes()
            If d.Parent Is Nothing Then
                Collapse(d.Root.Name)
            Else
                Collapse(d.Parent.FullName)
            End If
            Expand(d.FullName)
            HighlightSelectedNodes()
            RaiseEvent DirectorySelected(Me, New DirectoryInfoEventArgs(d))
        End Set
    End Property

#End Region

#Region "Methods"

    Public Iterator Function GetSelectedFiles() As IEnumerable(Of FileInfo)
        For Each node As FileSystemNode In m_selectedNodes.Values
            If (node.NodeType = FileSystemNodeType.File) Then
                Dim fnode As FileNode = CType(node, FileNode)
                Yield fnode.File
            End If
        Next
    End Function

    Public Iterator Function GetSelectedDirectories() As IEnumerable(Of DirectoryInfo)
        For Each node As FileSystemNode In m_selectedNodes.Values
            If (node.NodeType = FileSystemNodeType.Directory) Then
                Dim dnode As DirectoryNode = CType(node, DirectoryNode)
                Yield dnode.Directory
            End If
        Next
    End Function


    Public Sub Expand(ByVal fullPath As String)
        Dim directoryNames As New List(Of String)(fullPath.Split(Path.DirectorySeparatorChar))
        Dim rootNode = CType(tvFiles.Nodes(0), FileSystemNode)
        If (rootNode.NodeType = FileSystemNodeType.Computer Or rootNode.NodeType = FileSystemNodeType.Drive) Then
            'Root or drive
            Dim driveName = directoryNames(0) + "\"
            If (rootNode.NodeType = FileSystemNodeType.Computer) Then
                'Root: Expand root, produce all drives.
                rootNode.Expand()
                For Each dn As DriveNode In CType(rootNode, ComputerNode).Nodes
                    ExpandDriveNode(dn, driveName, directoryNames)
                    If directoryNames.Count = 2 And fullPath(Len(fullPath) - 2) = ":" Then
                        Dim k As New TreeNode
                        k = CType(dn, TreeNode)
                        tvFiles.SelectedNode = k
                        OnDriveSelected(New DriveInfo(dn.FullName))
                    End If
                Next
            Else
                'Only where root is a drive, which I don't use. 
                Dim dn As DriveNode = CType(rootNode, DriveNode)
                If (dn.Name.ToLower() = driveName.ToLower()) Then
                    rootNode.Expand()
                    ExpandDriveNode(dn, driveName, directoryNames)

                End If
            End If
        End If
    End Sub

    Public Sub RefreshTree(ByVal path As String)
        'Dim nd As New DirectoryNode(New DirectoryInfo(path))
        'CreateDirectoryTree(nd)
        Collapse(path)
        Expand(path)
        '    RaiseEvent DirectorySelected(Me, New DirectoryInfoEventArgs(New DirectoryInfo(path)))

    End Sub
    Private Sub ExpandDriveNode(ByVal dn As DriveNode, ByVal driveName As String, ByVal directoryNames As List(Of String))
        'Only expand if it's in the path we seek
        If (dn.Drive.Name.ToLower() = driveName.ToLower()) Then
            dn.Expand()
            'If directoryNames.Count = 1 Then 'If we have reached the end of the path select that node
            '    Dim k As New TreeNode
            '    k = CType(dn, TreeNode)
            '    tvFiles.SelectedNode = k
            '    OnDriveSelected(New DriveInfo(dn.FullName))
            'Else
            Expand(dn.DirectoryNodes, directoryNames, 1, directoryNames.Count - 1)
            'End If
        End If
    End Sub

    Private Sub Expand(ByVal directories As IEnumerable(Of DirectoryNode), ByVal directoryNames As List(Of String), ByVal index As Int32, count As Int32)
        'directories is the list of just expanded directories, and directoryNames is the components of the sought path
        If (directories Is Nothing Or index > directoryNames.Count - 1) Then
            Exit Sub
        End If

        For Each d In directories 'This seeks the path to expand, of the ones just epanded
            If (directoryNames(index).ToLower() = d.Directory.Name.ToLower()) Then
                d.Expand()
                Expand(d.DirectoryNodes, directoryNames, index + 1, count) 'iterate the expansion
                'if we've finished expanding, then select the node, and highlight it
                If index = count Then 'This is triggered far too frequently, and causes problems.  
                    Dim k As New TreeNode
                    k = CType(d, TreeNode)
                    tvFiles.SelectedNode = k
                    OnDirectorySelected(New DirectoryInfo(d.FullName))
                    '                    k.BackColor = SystemColors.HighlightText
                    Exit Sub
                End If
            End If

        Next
    End Sub

    Public Sub Collapse(ByVal fullPath As String)
        Dim directoryNames As New List(Of String)(fullPath.Split(Path.DirectorySeparatorChar))
        Dim rootNode = CType(tvFiles.Nodes(0), FileSystemNode)
        If (rootNode.NodeType = FileSystemNodeType.Computer Or rootNode.NodeType = FileSystemNodeType.Drive) Then
            Dim driveName = directoryNames(0) + "\"
            If (rootNode.NodeType = FileSystemNodeType.Computer) Then
                For Each dn As DriveNode In CType(rootNode, ComputerNode).Nodes
                    CollapseDriveNode(dn, driveName, directoryNames)
                Next
                rootNode.Collapse()
            Else
                Dim dn As DriveNode = CType(rootNode, DriveNode)
                If (dn.Name.ToLower() = driveName.ToLower()) Then
                    CollapseDriveNode(dn, driveName, directoryNames)
                    rootNode.Collapse()
                End If
            End If
        End If
    End Sub

    Private Sub CollapseDriveNode(ByVal dn As DriveNode, ByVal driveName As String, ByVal directoryNames As List(Of String))
        If (dn.Drive.Name.ToLower() = driveName.ToLower()) Then
            Collapse(dn.DirectoryNodes, directoryNames, 1)
            dn.Collapse()
        End If
    End Sub

    Private Sub Collapse(ByVal directories As IEnumerable(Of DirectoryNode), ByVal directoryNames As List(Of String), ByVal index As Int32)
        If (directories Is Nothing Or index > directoryNames.Count - 1) Then
            Return
        End If
        For Each d In directories
            If (directoryNames(index).ToLower() = d.Directory.Name.ToLower()) Then
                Collapse(d.DirectoryNodes, directoryNames, index + 1)
                d.Collapse()
            End If
        Next
    End Sub

    ''' <summary>
    ''' Enables drive state checking timer.
    ''' </summary>
    Private Sub EnableDriveStateCheck()
        Me.m_timer = New Timer()
        Me.m_timer.Interval = Convert.ToInt32(TimeSpan.FromSeconds(3).TotalMilliseconds)
        AddHandler Me.m_timer.Tick, AddressOf Me.Timer_Tick
        Me.m_timer.Start()
    End Sub

    ''' <summary>
    ''' Disables drive state checking timer.
    ''' </summary>
    Private Sub DisableDriveStateCheck()
        Me.m_timer.Stop()
        RemoveHandler Me.m_timer.Tick, AddressOf Me.Timer_Tick
        Me.m_timer.Dispose()
        Me.m_timer = Nothing
    End Sub

    ''' <summary>
    ''' Creates directory tree for expanded drive node.
    ''' </summary>
    Private Overloads Sub CreateDirectoryTree(ByVal node As DriveNode)
        If (IsReady(node.Drive) = True) Then

            Dim remove As New List(Of DirectoryNode)

            For Each dirNode As DirectoryNode In node.DirectoryNodes
                Try
                    Dim directories = dirNode.Directory.GetDirectories()
                    CreateDirectoryTree(dirNode, directories)
                    CreateFileTree(dirNode)
                Catch ex As UnauthorizedAccessException
                    If (HideUnauthorizedDirectories = True) Then
                        remove.Add(dirNode)
                    End If
                    Continue For
                End Try
            Next

            If (HideUnauthorizedDirectories = True) Then
                For Each dirNode As DirectoryNode In remove
                    node.Nodes.Remove(dirNode)
                Next
            End If
        End If
    End Sub


    ''' <summary>
    ''' Creates directory tree for exanded directory node.
    ''' </summary>
    Private Overloads Sub CreateDirectoryTree(ByVal node As DirectoryNode)

        Dim remove As New List(Of DirectoryNode)

        For Each dirNode As DirectoryNode In node.DirectoryNodes
            Try
                Dim directories = dirNode.Directory.GetDirectories()
                CreateDirectoryTree(dirNode, directories)
                CreateFileTree(dirNode)
            Catch ex As System.IO.DirectoryNotFoundException
                If (HideUnauthorizedDirectories = True) Then
                    remove.Add(dirNode)
                End If
                Continue For
            Catch ex As UnauthorizedAccessException
                If (HideUnauthorizedDirectories = True) Then
                    remove.Add(dirNode)
                End If
                Continue For
            End Try
        Next

        If (HideUnauthorizedDirectories = True) Then
            For Each dirNode As DirectoryNode In remove
                node.Nodes.Remove(dirNode)
            Next
        End If

    End Sub

    ''' <summary>
    ''' Creates directory tree for the directory node.
    ''' </summary>
    Private Overloads Sub CreateDirectoryTree(ByVal parent As DirectoryTreeNode, ByVal directories As IEnumerable(Of DirectoryInfo))

        For Each dir As DirectoryInfo In directories

            ' skip hidden
            If (dir.Attributes.HasFlag(FileAttributes.Hidden) And ShowHidden = True) Then
                Continue For
            End If

            ' skip system
            If ((dir.Attributes.HasFlag(FileAttributes.System) Or dir.Name.StartsWith("$")) And ShowSystem = False) Then
                Continue For
            End If

            parent.Add(dir)

        Next

    End Sub

    ''' <summary>
    ''' Creates file tree for expanded directory node.
    ''' </summary>
    Private Overloads Sub CreateFileTree(ByVal node As DirectoryNode)
        If Not ShowFiles Then Exit Sub
        If (node.FileNodes.Count() = 0) Then
            Dim files = GetFiles(node.Directory)
            CreateFileTree(node, files)
        End If
    End Sub

    ''' <summary>
    ''' Creates file tree for the directory node.
    ''' </summary>
    Private Overloads Sub CreateFileTree(ByVal parent As DirectoryNode, ByVal files As IEnumerable(Of FileInfo))

        For Each fl As FileInfo In files

            ' skip hidden files if hidden should not be shown
            If (fl.Attributes.HasFlag(FileAttributes.Hidden) And ShowHidden = False) Then
                Continue For
            End If

            ' skip system files if system should not be shown
            If (fl.Attributes.HasFlag(FileAttributes.System) And ShowSystem = False) Then
                Continue For
            End If

            parent.Add(fl)

        Next

    End Sub

    ''' <summary>
    ''' Creates initial tree that contains my computer node
    ''' and drive nodes or root drive node.
    ''' </summary>
    Private Sub CreateInitialTree()

        Me.tvFiles.Nodes.Clear()

        Dim compNode = ResolveRootNode()

        Dim rootType = compNode.NodeType

        If (rootType = FileSystemNodeType.Computer) Then

            Dim drives = DriveInfo.GetDrives()

            For Each drive As DriveInfo In drives

                ' skip unknown drives
                If (drive.DriveType = DriveType.Unknown) Then
                    Continue For
                End If

                ' if inactive drive should not be shown and drive is not read, skip
                If (ShowInactiveDrives = False And drive.IsReady = False) Then
                    Continue For
                End If

                ' skip network drives if should not show
                If (drive.DriveType = DriveType.Network And ShowNetworkDrives = False) Then
                    Continue For
                End If

                ' skip ram drives if should not show
                If (drive.DriveType = DriveType.Ram And ShowRamDrives = False) Then
                    Continue For
                End If

                ' skid cd/dvd-rom and other removable if should not show
                If ((drive.DriveType = DriveType.CDRom Or drive.DriveType = DriveType.Removable) And ShowRemovableDrives = False) Then
                    Continue For
                End If


                Dim drvNode As DriveNode = CType(compNode, ComputerNode).Add(drive)

                If (ShowInactiveDrives And drive.IsReady = False) Then
                    Me.m_inactiveDrives.Add(drive.Name)
                End If

                CreateInitialDriveTree(drvNode)

            Next
        Else
            Dim drvNode As DriveNode = CType(compNode, DriveNode)
            If (ShowInactiveDrives And drvNode.Drive.IsReady = False) Then
                Me.m_inactiveDrives.Add(drvNode.Drive.Name)
            End If

            CreateInitialDriveTree(drvNode)
        End If


        Me.tvFiles.Nodes.Add(compNode)

    End Sub

    ''' <summary>
    ''' Resolves the root node of the tree. Either computer node or specified drive node.
    ''' </summary>
    Private Function ResolveRootNode() As FileSystemNode
        If (String.IsNullOrWhiteSpace(RootDrive)) Then
            Return New ComputerNode()
        End If
        Dim drives = DriveInfo.GetDrives()
        For Each drive As DriveInfo In drives
            If (drive.Name = RootDrive) Then
                Return New DriveNode(drive)
            End If
        Next
        Return New ComputerNode()
    End Function

    ''' <summary>
    ''' Creates initial drive tree by ignoring drives that are not ready.
    ''' </summary>
    Private Sub CreateInitialDriveTree(ByVal node As DriveNode)
        If (node.Drive.IsReady) Then
            Dim directories = GetDirectories(node.Drive)
            CreateDirectoryTree(node, directories)
        End If
    End Sub

    ''' <summary>
    ''' Gets directories of the provided drive.
    ''' </summary>
    Private Function GetDirectories(ByVal drive As DriveInfo) As IEnumerable(Of DirectoryInfo)

        Dim directories As New List(Of DirectoryInfo)

        Dim directoryNames() As String = Directory.GetDirectories(drive.Name)

        For Each directoryName As String In directoryNames
            Dim directory As New DirectoryInfo(directoryName)
            If (directory.Exists) Then
                directories.Add(directory)
            End If
        Next

        Return directories

    End Function

    ''' <summary>
    ''' Gets files of the provided directory.
    ''' </summary>
    Private Function GetFiles(ByVal directory As DirectoryInfo) As IEnumerable(Of FileInfo)
        If (FileExtensions = "*.*") Then
            Return directory.GetFiles()
        End If

        Dim files As New List(Of FileInfo)

        Dim extens() As String = FileExtensions.Split(";")

        If (extens.Length = 0) Then
            Return files
        End If

        For Each ex As String In extens
            If (String.IsNullOrWhiteSpace(ex)) Then
                Continue For
            End If

            Dim tmp As String = ex

            If Not (tmp.StartsWith("*.")) Then

                ' check if dot is not in extension
                If Not (tmp.StartsWith(".")) Then
                    tmp = "." + tmp
                End If

                ' check if star is not in extension
                If Not (tmp.StartsWith("*")) Then
                    tmp = "*" + tmp
                End If

            ElseIf Not (tmp.StartsWith("*")) Then
                tmp = "*" + tmp
            End If

            ' get files matching the extension
            Dim fls() As FileInfo = directory.GetFiles(tmp)

            ' append if any
            If (fls.Length > 0) Then
                files.AddRange(fls)
            End If
        Next

        Return files

    End Function

    ''' <summary>
    ''' Executed when node is expanding.
    ''' </summary>
    Private Sub OnNodeExpanding(ByVal node As FileSystemNode)
        If (node IsNot Nothing) Then

            Select Case node.NodeType
                Case FileSystemNodeType.Drive
                    CreateDirectoryTree(CType(node, DriveNode))
                Case FileSystemNodeType.Directory
                    CreateDirectoryTree(CType(node, DirectoryNode))
                    CreateFileTree(CType(node, DirectoryNode))
            End Select

            OnNodeSelected(node)

        End If
    End Sub

    ''' <summary>
    ''' Executed when node is clicked.
    ''' </summary>
    Private Sub OnNodeClicked(ByVal node As FileSystemNode)

        If (node IsNot Nothing) Then

            If (node.NodeType = FileSystemNodeType.Drive) Then

                Dim dn As DriveNode = CType(node, DriveNode)

                If Not (dn.Drive.IsReady) Then
                    If (IsReady(dn.Drive)) Then
                        CreateInitialDriveTree(dn)
                    End If
                End If

            End If

            OnNodeSelected(node)

        End If
    End Sub

    ''' <summary>
    ''' Executed when node is collapsed.
    ''' </summary>
    Private Sub OnNodeCollapsed(ByVal node As TreeNode)
        For Each n As FileSystemNode In node.Nodes
            If (n.NodeType <> FileSystemNodeType.Computer And n.NodeType <> FileSystemNodeType.Drive) Then
                If (n.Nodes.Count > 0) Then
                    n.Nodes.Clear()
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Executed when node is selected.
    ''' </summary>
    Private Sub OnNodeSelected(ByVal node As FileSystemNode)
        Select Case node.NodeType
            Case FileSystemNodeType.Drive
                Dim drive As DriveInfo = CType(node, DriveNode).Drive
                OnDriveSelected(drive)
            Case FileSystemNodeType.Directory
                Dim directory As DirectoryInfo = CType(node, DirectoryNode).Directory
                'OnDirectorySelected(directory)
            Case FileSystemNodeType.File
                Dim file As FileInfo = CType(node, FileNode).File
                OnFileSelected(file)
        End Select

    End Sub

    ''' <summary>
    ''' Checks if drive is ready and if not, prompts user to insert media.
    ''' </summary>
    Private Function IsReady(ByVal drive As DriveInfo) As Boolean

        ' drive is already active
        If (drive.IsReady) Then
            Return True
        End If

        ' message box content
        Dim msg As String = String.Format("{0} drive is not ready. Please insert media.", drive.Name)
        Dim caption As String = String.Format("{0} drive not ready", drive.Name)

        ' prompt media insert
        If (MessageBox.Show(msg, caption, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = DialogResult.OK) Then

            Try
                ' calculate time out and use stopwatch to calculate
                Dim timeout As Int64 = Convert.ToInt64(TimeSpan.FromSeconds(DriveReadTimeout).TotalMilliseconds)
                Dim watch As Stopwatch = Stopwatch.StartNew()
                Dim sleepTime As Int32 = Convert.ToInt32(timeout / 10)
                Cursor.Current = Cursors.WaitCursor

                While (True)
                    Dim elapsed As Int64 = watch.ElapsedMilliseconds

                    ' timeout is reached
                    If (elapsed >= timeout) Then
                        watch.Stop()
                        Exit While
                    End If

                    ' drive has become active
                    If (drive.IsReady) Then
                        Return True
                    End If

                    System.Threading.Thread.Sleep(sleepTime)

                End While

            Finally
                Cursor.Current = Cursors.Arrow
            End Try

            msg = String.Format("{0} drive could not be read in specified timeout.", drive.Name)

            MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End If

        Return False

    End Function

    ''' <summary>
    ''' Raises DriveActivated event.
    ''' </summary>
    Private Sub OnDriveActivated(ByVal drive As DriveInfo)
        RaiseEvent DriveActivated(Me, New DriveInfoEventArgs(drive))
    End Sub

    ''' <summary>
    ''' Raises DriveUnactivated event.
    ''' </summary>
    Private Sub OnDriveUnactivated(ByVal drive As DriveInfo)
        RaiseEvent DriveUnactivated(Me, New DriveInfoEventArgs(drive))
    End Sub

    ''' <summary>
    ''' Raises DriveSelected event.
    ''' </summary>
    Private Sub OnDriveSelected(ByVal drive As DriveInfo)
        RaiseEvent DriveSelected(Me, New DriveInfoEventArgs(drive))
    End Sub

    ''' <summary>
    ''' Raises DirectorySelected event.
    ''' </summary>
    Private Sub OnDirectorySelected(ByVal directory As DirectoryInfo)
        RaiseEvent DirectorySelected(Me, New DirectoryInfoEventArgs(directory))
    End Sub

    ''' <summary>
    ''' Raises FileSelected event.
    ''' </summary>
    Private Sub OnFileSelected(ByVal file As FileInfo)
        RaiseEvent FileSelected(Me, New FileInfoEventArgs(file))
    End Sub

    ''' <summary>
    ''' Clears the selected nodes collection.
    ''' </summary>
    Private Sub ClearSelectedNodes()
        For Each node In m_selectedNodes.Values
            node.BackColor = Color.Empty
            node.ForeColor = Color.Empty
        Next
        m_selectedNodes.Clear()
    End Sub

    ''' <summary>
    ''' Highlights the selected nodes.
    ''' </summary>
    Public Sub HighlightSelectedNodes()
        For Each node In m_selectedNodes.Values
            node.BackColor = SystemColors.Highlight
            node.ForeColor = SystemColors.HighlightText
            m_currentNode = node
        Next
    End Sub

#End Region

#Region "Event Handlers"

    Private Sub FileSystemTree_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Dock = DockStyle.Fill
        ' if drive state should be tracked,
        ' enable timer that does that
        If (TrackDriveState) Then
            EnableDriveStateCheck()
        End If
    End Sub

    Private Sub FileSystemTree_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
        ' if visible and initial tree is not created, then create it
        If (Visible And Me.m_initialTreeCreate = False) Then
            CreateInitialTree()
            Me.m_initialTreeCreate = True
        End If
    End Sub

    Private Sub FileSystemTree_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        ' if not visible, clear the tree and set so that initial tree
        ' will be recreated
        If (Visible = False) Then
            Me.tvFiles.Nodes.Clear()
            Me.m_initialTreeCreate = False
        End If
    End Sub

    Private Sub FileSystemTree_Disposed(sender As Object, e As EventArgs) Handles MyBase.Disposed

        ' if tracking drive state, disable it
        If (TrackDriveState) Then
            DisableDriveStateCheck()
        End If

    End Sub
    Public Sub RemoveNode(strpath As String)
        Expand(strpath)
        Dim nd As FileSystemNode = tvFiles.SelectedNode
        Traverse(False)
        nd.Remove()
        'Expand(tvFiles.SelectedNode.FullPath)
    End Sub
    Public Sub MoveNode(strDir As String, StrDest As String)
        tvFiles.Nodes.RemoveByKey(strDir)
        Dim nd As New DirectoryNode(New DirectoryInfo(StrDest))
        Expand(StrDest)
    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs)
        Exit Sub
        Try
            For Each node As FileSystemNode In Me.tvFiles.Nodes(0).Nodes

                ' if node is drive node
                If (node.NodeType = FileSystemNodeType.Drive) Then

                    Dim drvNode As DriveNode = CType(node, DriveNode)

                    ' check if drive has become ready or changed to not-ready
                    If (drvNode.Drive.IsReady And Me.m_inactiveDrives.Contains(drvNode.Drive.Name)) Then
                        ' when ready remove from inactive, create initial tree and raise event
                        Me.m_inactiveDrives.Remove(drvNode.Drive.Name)
                        CreateInitialDriveTree(drvNode)
                        OnDriveActivated(drvNode.Drive)
                    ElseIf (drvNode.Drive.IsReady = False And Not Me.m_inactiveDrives.Contains(drvNode.Drive.Name)) Then
                        ' when not ready add to inactive, clear content and raise event
                        Me.m_inactiveDrives.Add(drvNode.Drive.Name)
                        If (node.Nodes.Count > 0) Then
                            node.Nodes.Clear()
                        End If
                        OnDriveUnactivated(drvNode.Drive)
                    End If
                End If

            Next
        Catch ex As Exception
        End Try

    End Sub

    Private Sub tvFiles_AfterCollapse(sender As Object, e As TreeViewEventArgs) Handles tvFiles.AfterCollapse
        OnNodeCollapsed(e.Node)
    End Sub


    Private Sub tvFiles_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles tvFiles.BeforeExpand
        'MsgBox("tvFiles_BeforeExpand")
        Dim node As FileSystemNode = CType(e.Node, FileSystemNode)
        OnNodeExpanding(node)
    End Sub

    Private Sub tvFiles_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles tvFiles.NodeMouseClick

        If (e.Button = Windows.Forms.MouseButtons.Left) Then
            Dim node As FileSystemNode = CType(e.Node, FileSystemNode)

            ' check if control key is pressed to select multiple items
            If (ModifierKeys = Keys.Control) Then
                If (m_selectedNodes.ContainsKey(node.FullName) = False) Then
                    m_selectedNodes.Add(node.FullName, node)
                End If
            Else
                ' no control, clear previous selections and select current node
                ClearSelectedNodes()
                m_selectedNodes.Add(node.FullName, node)
            End If

            ' highlight selected nodes
            HighlightSelectedNodes()

            OnNodeClicked(node)
        End If

    End Sub



    Private Sub tvFiles_BeforeSelect(sender As Object, e As TreeViewCancelEventArgs) Handles tvFiles.BeforeSelect
        'MsgBox("tvFiles_BeforeSelect") 'Repeatedly called
        ClearSelectedNodes()

    End Sub

    Private Sub tvFiles_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvFiles.AfterSelect
        '    'MsgBox("tvFiles_AfterSelect")
        HighlightSelectedNodes()
        If e.Node.FullPath <> "" Then RaiseEvent DirectorySelected(Me, New DirectoryInfoEventArgs(New DirectoryInfo(NodePath(e)))) 'TODO Causing crashes

        '    ''RaiseEvent DriveSelected(Me, New DriveInfoEventArgs(New DriveInfo(NodePath(e))))

    End Sub
    Private Function NodePath(e) As String
        Dim s As String
        Dim myc As String = "My Computer\"
        s = e.node.FullPath
        If InStr(s, myc) <> 0 Then s = s.Remove(0, Len(myc))
        Return s
    End Function



    Public Sub Traverse(blnBack As Boolean)
        'Exit Sub
        Dim node As TreeNode
        Dim newnode As TreeNode
        node = tvFiles.SelectedNode
        If node Is Nothing Then
            node = tvFiles.Nodes(0)
        End If
        If Not blnBack Then

            'If node has children move to first child
            If node.Nodes.Count <> 0 Then
                newnode = node.Nodes(0)

            Else
                'Otherwise move to next sibling
                If node.NextNode IsNot Nothing Then
                    newnode = node.NextNode
                Else
                    While node.NextNode Is Nothing
                        node = node.Parent
                    End While
                    newnode = node.NextNode
                End If
            End If

        Else

            'Traverse backward
            'If there is an older sibling, move to its youngest, most removed descendant. 
            If node.PrevNode IsNot Nothing Then
                newnode = node.PrevNode
                While newnode.Nodes.Count <> 0
                    newnode = newnode.Nodes(newnode.Nodes.Count - 1)
                End While
            Else
                'Otheriwse move to parent

                newnode = node.Parent
            End If
        End If
        ClearSelectedNodes()
        tvFiles.SelectedNode = newnode
        HighlightSelectedNodes()
    End Sub

    Private Sub FileSystemTree_DirectorySelected(sender As Object, e As DirectoryInfoEventArgs) Handles Me.DirectorySelected
        newSelectedFolder = e.Directory.FullName
    End Sub

    Private Sub FileSystemTree_BackColorChanged(sender As Object, e As EventArgs) Handles Me.BackColorChanged
        tvFiles.BackColor = Me.BackColor
    End Sub


    Public Sub AfterLabelEdit(sender As Object, e As NodeLabelEditEventArgs) Handles tvFiles.AfterLabelEdit
        Dim oldlabel As String = e.Label

        e = RelabelFolder(e)
        Dim node As FileSystemNode = CType(e.Node, FileSystemNode)
        Dim pnode As FileSystemNode = CType(e.Node.Parent, FileSystemNode)
        'Try
        Dim m As New DirectoryInfo(node.FullName)
        Dim s As String = Replace(e.Node.FullPath, node.Text, e.Label)
        s = Replace(s, "My Computer\", "")
        s = Replace(s, "\\", "\")
        '      RaiseEvent LabelEdited(sender, s)
        If m.FullName <> s Then
            Dim n As New DirectoryInfo(s)
            If n.Exists Then

            Else
                m.MoveTo(s)
            End If
        End If
        '  Me.RefreshTree()

    End Sub

    Private Shared Function RelabelFolder(e As NodeLabelEditEventArgs) As NodeLabelEditEventArgs
        If Not (e.Label Is Nothing) Then
            If e.Label.Length > 0 Then
                If e.Label.IndexOfAny(New Char() {"@"c, ","c, "!"c}) = -1 Then
                    ' Stop editing without canceling the label change.
                    e.Node.EndEdit(False)

                Else
                    ' Cancel the label edit action, inform the user, and
                    ' place the node in edit mode again. 
                    e.CancelEdit = True
                    MessageBox.Show("Invalid tree node label." &
              Microsoft.VisualBasic.ControlChars.Cr &
              "The invalid characters are: '@', ',', '!'",
              "Node Label Edit")
                    e.Node.BeginEdit()
                End If
            Else
                ' Cancel the label edit action, inform the user, and
                ' place the node in edit mode again. 
                e.CancelEdit = True
                MessageBox.Show("Invalid tree node label." &
           Microsoft.VisualBasic.ControlChars.Cr &
           "The label cannot be blank", "Node Label Edit")
                e.Node.BeginEdit()
            End If
        End If

        Return e
    End Function

    Public Sub tvFiles_KeyDown(sender As Object, e As KeyEventArgs) Handles tvFiles.KeyDown
        Select Case e.KeyCode
            Case TraverseKey, TraverseKeyBack
                Traverse(e.KeyCode = TraverseKeyBack)
            Case Keys.F2
                tvFiles.LabelEdit = True
                Dim nd As TreeNode = tvFiles.SelectedNode
                If Not (nd Is Nothing) And Not (nd.Parent Is Nothing) Then
                    tvFiles.SelectedNode = nd
                    tvFiles.LabelEdit = True
                    If Not nd.IsEditing Then
                        nd.BeginEdit()
                    End If
                Else
                    MessageBox.Show("No tree node selected or selected node is a root node." &
                  Microsoft.VisualBasic.ControlChars.Cr &
                  "Editing of root nodes is not allowed.", "Invalid selection")
                End If
            Case Keys.Left, Keys.Right, Keys.Up, Keys.Down

            Case Else
                e.SuppressKeyPress = True
        End Select

    End Sub

    Private Sub tvFiles_DoubleClick(sender As Object, e As EventArgs) Handles tvFiles.DoubleClick

    End Sub

    Private Sub tvFiles_Click(sender As Object, e As EventArgs) Handles tvFiles.Click
        'MsgBox(e.ToString)
    End Sub




#End Region

End Class
