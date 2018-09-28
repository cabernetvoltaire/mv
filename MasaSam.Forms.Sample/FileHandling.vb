Imports System.IO
Imports MasaSam.Forms.Controls
Imports System.Threading
Module FileHandling
    Public blnSuppressCreate As Boolean = False
    Public blnChooseOne As Boolean = False
    Public VIDEOEXTENSIONS = ".divx.vob.webm.avi.flv.mov.m4p.mpeg.mpg.m4a.m4v.mkv.mp4.rm.ram.wmv.wav.mp3.3gp"
    Public PICEXTENSIONS = "arw.jpeg.png.jpg.bmp.gif"
    Public CurrentfilterState As FilterHandler = MainForm.CurrentFilterState
    Public Random As RandomHandler = MainForm.Random
    Public NavigateMoveState As StateHandler = MainForm.NavigateMoveState
    Public FilePumpList As New List(Of String)
    Public FP As New FilePump
    Public Event FolderMoved(Path As String)
    Public t As Thread

    Dim strFilterExtensions(6) As String
    Public Sub AssignExtensionFilters()
        strFilterExtensions(FilterHandler.FilterState.All) = ""
        strFilterExtensions(FilterHandler.FilterState.Piconly) = PICEXTENSIONS
        strFilterExtensions(FilterHandler.FilterState.PicVid) = PICEXTENSIONS & VIDEOEXTENSIONS
        strFilterExtensions(FilterHandler.FilterState.LinkOnly) = ".lnk"
        strFilterExtensions(FilterHandler.FilterState.Vidonly) = VIDEOEXTENSIONS
        strFilterExtensions(FilterHandler.FilterState.NoPicVid) = PICEXTENSIONS & VIDEOEXTENSIONS & "NOT"
    End Sub

    ''' <summary>
    ''' Fills the listbox with files from a given folder, in a given filter state
    ''' </summary>
    ''' <param name="lbx"></param>
    ''' <param name="e"></param>
    ''' <param name="flist"></param>
    ''' <param name="blnRandom"></param>
    Public Sub FillListbox(lbx As ListBox, e As IO.DirectoryInfo, ByVal flist As List(Of String), blnRandom As Boolean)
        If Not e.Exists Then Exit Sub
        If e.Name = "My Computer" Then Exit Sub

        lbx.Items.Clear()
        flist.Clear()

        If e.EnumerateFiles.Count = 0 Then
            lbx.Items.Add("If there is nothing showing here, check your filters")

            Exit Sub
        End If
        Try
            FilterListBox(e, flist)
            flist = SetPlayOrder(MainForm.PlayOrder.State, flist)
            FillShowbox(lbx, MainForm.CurrentFilterState.State, flist)

            lbx.Tag = e


            If lbx.Items.Count <> 0 Then
                If Random.OnDirChange Or Random.NextSelect Then
                    Dim s As Long = lbx.Items.Count - 1
                    lbx.SelectedIndex = Rnd() * s
                Else
                    lbx.SelectedIndex = lbx.FindString(Media.MediaPath)
                    If lbx.SelectedIndex = -1 Then lbx.SelectedIndex = 0
                End If
            End If

        Catch ex As ArgumentException
            MsgBox(ex.ToString)
        Catch ex As IOException
            Exit Try
        End Try


    End Sub
    Private Function FilterListBox(e As DirectoryInfo, ByVal lst As List(Of String))
        'Dim l As New List(Of String)
        For Each f In e.EnumerateFiles
            Dim s As String = LCase(f.Extension)

            Select Case CurrentfilterState.State
                Case FilterHandler.FilterState.All

                    lst.Add(f.FullName)
                Case FilterHandler.FilterState.NoPicVid

                    If InStr(VIDEOEXTENSIONS & PICEXTENSIONS, s) <> 0 And Len(s) > 0 Then
                    Else
                        lst.Add(f.FullName)

                    End If

                Case FilterHandler.FilterState.Piconly
                    If InStr(PICEXTENSIONS, s) <> 0 And Len(s) > 0 Then

                        lst.Add(f.FullName)
                    Else
                    End If
                Case FilterHandler.FilterState.LinkOnly
                    If s = ".lnk" Then
                        lst.Add(f.FullName)
                    End If

                Case FilterHandler.FilterState.Vidonly
                    If InStr(VIDEOEXTENSIONS, s) <> 0 And Len(s) > 0 Then
                        lst.Add(f.FullName)
                    End If

                Case FilterHandler.FilterState.PicVid
                    If InStr(VIDEOEXTENSIONS & PICEXTENSIONS, s) <> 0 And Len(s) > 0 Then
                        lst.Add(f.FullName)
                    End If
            End Select

        Next
    End Function




    ''' <summary>
    ''' Actually saves the store list
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="Dest"></param>
    Public Sub StoreList(list As List(Of String), Dest As String)
        If Dest = "" Then Exit Sub
        Dim fs As New StreamWriter(New FileStream(Dest, FileMode.Create, FileAccess.Write))
        'fs.WriteLine(list.Count)
        For Each s In list
            fs.WriteLine(s)

        Next
        fs.Close()
    End Sub
    ''' <summary>
    ''' Loads Dest into List, and adds all to lbx. Any files not found are put in notlist, which can then be removed from the lbx
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="Dest"></param>
    ''' <param name="lbx"></param>
    '''
    Public Sub Getlist(list As List(Of String), Dest As String, lbx As ListBox)

        Dim notlist As New List(Of String)
        Dim count As Long = 0


        Dim fs As New StreamReader(New FileStream(Dest, FileMode.OpenOrCreate, FileAccess.Read))
        'count = fs.ReadLine

        Do While fs.Peek <> -1
            Dim s As String = fs.ReadLine

            Try
                Dim f As New FileInfo(s)
                If f.Exists Then
                    list.Add(s)
                    count += 1
                    lbx.Items.Add(s)
                Else
                    notlist.Add(s)
                End If
            Catch ex As System.IO.PathTooLongException
                Continue Do
            Catch ex As System.ArgumentException
                ReportFault("Filehandling.Getlist", ex.Message)
                Exit Sub
            End Try


            ProgressIncrement(40)
        Loop

        If lbx.Items.Count <> 0 Then lbx.TabStop = True
        fs.Close()
        lngShowlistLines = Showlist.Count

        If notlist.Count = 0 Then Exit Sub
        If MsgBox(notlist.Count & " files were not found. Remove from list?", vbYesNo, "Metavisua") = MsgBoxResult.Yes Then
            For Each s In notlist
                list.Remove(s)
            Next
            If MsgBox("Re-save list?", vbYesNo, "Metavisua") Then
                StoreList(list, Dest)
            End If
        End If
    End Sub
    Public Sub PromoteFolder(folder As DirectoryInfo)
        MainForm.CancelDisplay()
        folder.MoveTo(folder.Parent.Parent.FullName & "\" & folder.Name)
        MainForm.tvMain2.RefreshTree(folder.FullName)
    End Sub
    Public Sub MoveFolder(strDir As String, strDest As String, blnOverride As Boolean)
        If strDest Is Nothing Then Exit Sub
        Try
            With My.Computer.FileSystem
                Dim dir = New DirectoryInfo(strDir)
                Dim s As String = dir.Name
                Dim f As New DirectoryInfo(dir.Parent.FullName)
                Dim destdir = New DirectoryInfo(strDest)
                Select Case NavigateMoveState.State
                    Case StateHandler.StateOptions.Copy
                        .CopyDirectory(strDir, strDest & "\" & s, FileIO.UIOption.OnlyErrorDialogs)
                    Case StateHandler.StateOptions.Move
                        Dim flist As FileInfo() = dir.GetFiles()
                        .MoveDirectory(strDir, strDest & "\" & s, FileIO.UIOption.OnlyErrorDialogs)

                        For Each x In flist
                            Movelink(x, strDest & "\" & s & "\" & x.Name)
                        Next

                    Case StateHandler.StateOptions.MoveLeavingLink
                        'Create link directory?
                        .MoveDirectory(strDir, strDest & "\" & s, FileIO.UIOption.OnlyErrorDialogs)

                    Case StateHandler.StateOptions.CopyLink
                        'Creat link directory?

                End Select
                UpdateButton(strDir, strDest & "\" & s) 'todo doesnt handle sub-tree
            End With
        Catch ex As Exception
            '            MsgBox(ex.Message) '
        End Try

    End Sub



    ''' <summary>
    ''' Moves files to strDest, and removes them from lbx1 asking for a subfolder in destination if more than one file. If strDest is empty, files are deleted.
    ''' </summary>
    ''' <param name="files"></param>
    ''' <param name="strDest"></param>
    ''' <param name="lbx1"></param>
    Public Sub MoveFiles(files As List(Of String), strDest As String, lbx1 As ListBox)
        Dim ind As Long = lbx1.SelectedIndex
        ''If only one file, then use the filepump
        'If files.Count = 1 And strDest <> "" Then 'TODO This causes Bundle not to work if only one file
        '    'Try
        '    '    FileIO.FileSystem.CopyFile(files(0), strDest & "\" & FileIO.FileSystem.GetName(files(0)), FileIO.UIOption.OnlyErrorDialogs)
        '    '    FileIO.FileSystem.DeleteFile(files(0))
        '    'Catch ex As Exception
        '    'End Try
        '    'Build a single filepump
        '    lbx1.Items.Remove(files(0))
        '    MainForm.tmrPumpFiles.Enabled = True
        '    FilePump(files.Item(0) & "|" & strDest, lbx1)
        '    lbx1.SelectionMode = SelectionMode.One
        '    If lbx1.Items.Count <> 0 Then lbx1.SetSelected(Math.Max(Math.Min(ind, lbx1.Items.Count - 1), 0), True)
        'Else

        Dim s As String = strDest 'if strDest is empty then delete

        If files.Count > 1 And strDest <> "" Then
            If Not blnSuppressCreate Then s = CreateNewDirectory(MainForm.tvMain2, strDest, True)
        End If
        t = New Thread(New ThreadStart(Sub() MovingFiles(files, strDest, s)))
        t.IsBackground = True
        t.SetApartmentState(ApartmentState.STA)

        t.Start()

        '  Dim file As String

        ' file = MovingFiles(files, strDest, s)
        lbx1.SelectionMode = SelectionMode.One
        For Each f In files
            If NavigateMoveState.State <> StateHandler.StateOptions.Copy Then
                lbx1.Items.Remove(f)
            End If
        Next
        '            lbx1.Refresh
        '  FillListbox(lbx1, New DirectoryInfo(Media.MediaDirectory), FileboxContents, False)
        If lbx1.Items.Count <> 0 Then lbx1.SetSelected(Math.Max(Math.Min(ind, lbx1.Items.Count - 1), 0), True)
        'End If



    End Sub


    ''' <summary>
    ''' Moves files to strDest, creating a subfolder s if specified, unless strDest is empty in which case, they are deleted.
    ''' </summary>
    ''' <param name="files"></param>
    ''' <param name="strDest"></param>
    ''' <param name="s"></param>
    ''' 
    ''' <returns></returns>
    Private Sub MovingFiles(files As List(Of String), strDest As String, s As String)

        Dim file As String = ""

        For Each file In files
            Dim m As New FileInfo(file)
            With My.Computer.FileSystem
                Dim i As Long = 0
                Dim spath As String
                If InStr(s, "\") = s.Length - 1 Or s = "" Then
                    spath = s & m.Name

                Else
                    spath = s & "\" & m.Name

                End If
                While .FileExists(spath) 'Existing path
                    Dim x = m.Extension
                    Dim b = InStr(spath, "(")
                    If b = 0 Then
                        spath = Replace(spath, x, "(" & i & ")" & x)
                    Else
                        spath = Left(spath, b - 1) & "(" & i & ")" & x
                    End If

                    i += 1
                End While
                Select Case NavigateMoveState.State
                    Case StateHandler.StateOptions.Copy
                        .CopyFile(m.FullName, spath)
                    Case StateHandler.StateOptions.Move, StateHandler.StateOptions.Navigate
                        If Not currentPicBox.Image Is Nothing Then DisposePic(currentPicBox)
                        If strDest = "" Then
                            Deletefile(m.FullName)

                        Else
                            Dim f As New IO.FileInfo(m.FullName)
                            m.MoveTo(spath)
                            Movelink(f, spath)

                        End If
                    Case StateHandler.StateOptions.MoveLeavingLink
                        'Move, and place link here
                        m.MoveTo(spath)
                        CreateLink(m.FullName, Media.MediaDirectory)

                    Case StateHandler.StateOptions.CopyLink
                        'Paste a link in destination
                        Dim fpath As New FileInfo(spath)
                        CreateLink(m.FullName, fpath.Directory.FullName)

                End Select
            End With
        Next

    End Sub
    ''' <summary>
    ''' Checks to see if f is in the favourite links, and if so, updates the link. 
    ''' </summary>
    ''' <param name="f"></param>
    ''' <param name="path"></param>

    Private Sub Movelink(f As IO.FileInfo, path As String)
        Dim links As New List(Of String)
        Dim faves As New IO.DirectoryInfo(FavesFolderPath)
        For Each j In faves.GetFiles
            If links.Contains(j.FullName) Then
            Else
                links.Add(LinkTarget(j.FullName))
            End If
        Next
        If links.Contains(f.FullName) Then
            CreateFavourite(path)
        End If
    End Sub

    Public Sub MoveFiles(filendest As List(Of String), lbx1 As ListBox)

        Dim ind As Long = lbx1.SelectedIndex

        Dim fd As String
        For Each fd In filendest
            Dim fds() As String
            fds = fd.Split("|")
            Dim file As String = fds(0)
            Dim s As String = fds(1)
            If Len(file) > 250 Then
                'TODO deal with too long file names
            End If
            Dim m As New FileInfo(file)
            With My.Computer.FileSystem
                Dim i As Long = 0
                Dim spath As String
                If InStr(s, "\") = s.Length - 1 Or s = "" Then
                    spath = s & m.Name

                Else
                    spath = s & "\" & m.Name

                End If
                While .FileExists(spath) 'Existing path
                    Dim x = m.Extension
                    Dim b = InStr(spath, "(")
                    If b = 0 Then
                        spath = Replace(spath, x, "(" & i & ")" & x)
                    Else
                        spath = Left(spath, b - 1) & "(" & i & ")" & x
                    End If

                    i += 1
                End While

                'Exit For
                Select Case NavigateMoveState.State
                    Case StateHandler.StateOptions.Copy
                        .CopyFile(m.FullName, spath)
                    Case StateHandler.StateOptions.Move
                        If Not currentPicBox.Image Is Nothing Then DisposePic(currentPicBox)
                        lbx1.Items.Remove(m.FullName)
                        If s = "" Then
                            Deletefile(m.FullName)
                        Else
                            Try
                                .MoveFile(m.FullName, spath, FileIO.UIOption.OnlyErrorDialogs, FileIO.UICancelOption.ThrowException)
                                Movelink(New IO.FileInfo(spath & "\" & m.Name), spath)

                            Catch ex As Exception
                                Exit For
                            End Try
                        End If
                    Case StateHandler.StateOptions.MoveLeavingLink
                    Case StateHandler.StateOptions.CopyLink
                    Case StateHandler.StateOptions.Navigate

                End Select
            End With
        Next
        lbx1.SelectionMode = SelectionMode.One
        'If lbx1.Items.Count <> 0 Then lbx1.SetSelected(Math.Max(Math.Min(ind, lbx1.Items.Count - 1), 0), True)

    End Sub

    Public Sub CreateNewDirectory(strDir As String)
        AddFolders.Show()
        AddFolders.Folder = strDir
        MainForm.tvMain2.RefreshTree(strDir)

    End Sub
    Public Function CreateNewDirectory(tv As FileSystemTree, strDest As String, blnAsk As Boolean) As String
        Dim blnCreate As Boolean = True
        Dim blnAssign As Boolean = False
        Dim s As String = ""
        If blnAsk Then
            s = InputBox("Name of folder to create? (Blank means none)", "Create sub-folder", lastselection)
            If s = "" Then blnCreate = False
            s = strDest & "\" & s
        Else
            s = strDest & "\"
        End If
        If blnCreate Then
            Try
                IO.Directory.CreateDirectory(s)
                tv.RefreshTree(strDest)
            Catch ex As IO.DirectoryNotFoundException
            End Try
            'If MsgBox("Assign new folder to button?", MsgBoxStyle.YesNo, "Assign folder") = MsgBoxResult.Yes Then '
            '    Dim i As Integer = -1
            '    While i > 4 + 8 Or i < 5
            '        i = InputBox("Number?",, "f")
            '    End While
            '    i = i - 5
            '    AssignButton(i, iCurrentAlpha, 1, s, True)
            'End If
        End If
        Return s
    End Function
    Public Sub AddCurrentType(Recurse As Boolean)
        AddFilesToCollection(Showlist, strFilterExtensions(CurrentfilterState.State), Recurse)
        FillShowbox(MainForm.lbxShowList, FilterHandler.FilterState.All, Showlist)

    End Sub
    Public Sub Addpics(Recurse As Boolean)
        AddFilesToCollection(Showlist, PICEXTENSIONS, Recurse)
        FillShowbox(MainForm.lbxShowList, CurrentfilterState.State, Showlist)
    End Sub
    Public Sub AddFilesToCollection(ByVal list As List(Of String), extensions As String, blnRecurse As Boolean)
        Dim s As String
        Dim d As New DirectoryInfo(Media.MediaDirectory)

        s = InputBox("Only include files containing? (Leave empty to add all)")
        If blnChooseOne Then
        Else
        End If
        MainForm.Cursor = Cursors.WaitCursor
        ProgressBarOn(1000)
        FindAllFilesBelow(d, list, extensions, False, s, blnRecurse, blnChooseOne)

        MainForm.Cursor = Cursors.Default

    End Sub
    Public Sub AddSingleFileToList(ByVal list As List(Of String), strPath As String)
        list.Add(strPath)

    End Sub
    Public Sub Deletefile(s As String)

        With My.Computer.FileSystem
            If .FileExists(s) Then
                Try
                    .DeleteFile(s, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    If MainForm.lbxFiles.Items.Contains(s) Then
                        MainForm.lbxFiles.Items.Remove(s)
                    End If
                Catch ex As Exception
                    'Exit Sub
                    'Catch ex As
                End Try
            End If

        End With
    End Sub
    ''' <summary>
    ''' Adds all files in d of given extension, or removes them, to the list, only including strSearch 
    ''' </summary>
    ''' <param name="d"></param>
    ''' <param name="list"></param>
    ''' <param name="extensions"></param>
    ''' <param name="blnRemove"></param>
    ''' <param name="strSearch"></param>
    ''' <param name="blnRecurse"></param>
    Public Sub FindAllFilesBelow(d As DirectoryInfo, list As List(Of String), extensions As String, blnRemove As Boolean, strSearch As String, blnRecurse As Boolean, blnOneOnly As Boolean)

        For Each file In d.EnumerateFiles
            Application.DoEvents()
            ProgressIncrement(1)
            Try
                If InStr(LCase(extensions), LCase("NOT")) <> 0 Then
                    If InStr(extensions, LCase(file.Extension)) = 0 And file.Extension <> "" Then
                        'Only include if NOT the given extension
                        AddRemove(list, blnRemove, strSearch, file)
                    End If
                Else

                    If InStr(extensions, LCase(file.Extension)) <> 0 And file.Extension <> "" Then 'File has an extension, and an appropriate one
                        AddRemove(list, blnRemove, strSearch, file)

                    Else
                        If extensions = "" Then
                            AddRemove(list, blnRemove, strSearch, file)

                        End If
                    End If
                End If
                If blnOneOnly Then Exit For 'Exits each folder when one file has been found matching the condition.
            Catch ex As PathTooLongException
                ReportFault("FindAllFilesBelow", ex.Message)
            End Try

            ProgressIncrement(1)
        Next
        If blnRecurse Then

            For Each di In d.EnumerateDirectories
                Try

                    FindAllFilesBelow(di, list, extensions, blnRemove, strSearch, blnRecurse, blnOneOnly)
                Catch ex As UnauthorizedAccessException
                    Continue For
                Catch ex As DirectoryNotFoundException
                    Continue For
                End Try
            Next
        End If
    End Sub

    Private Sub AddRemove(list As List(Of String), blnRemove As Boolean, strSearch As String, file As FileInfo)
        If InStr(LCase(file.FullName), LCase(strSearch)) <> 0 Or strSearch = "" Then
            If blnRemove Then
                list.Remove(file.FullName)
            Else
                list.Add(file.FullName)
            End If
        End If
    End Sub

    Public Sub FindAllFoldersBelow(d As DirectoryInfo, list As List(Of String), blnRecurse As Boolean, blnNonEmptyOnly As Boolean)


        For Each di In d.EnumerateDirectories
            Try
                If blnNonEmptyOnly AndAlso di.EnumerateFiles.Count > 0 Or Not blnNonEmptyOnly Then
                    list.Add(di.FullName)
                End If

            Catch ex As PathTooLongException
                ReportFault("FindAllFoldersBelow", ex.Message)
            Catch ex As System.UnauthorizedAccessException
                Continue For
            End Try

        Next
        If blnRecurse Then

            For Each di In d.EnumerateDirectories
                Try

                    FindAllFoldersBelow(di, list, blnRecurse, blnNonEmptyOnly)
                Catch ex As UnauthorizedAccessException
                    Continue For
                Catch ex As DirectoryNotFoundException
                    Continue For
                End Try
            Next
        End If
    End Sub
    ''' <summary>
    ''' Moves all the files in d to its parent, if it exists. Returns true if successful
    ''' </summary>
    ''' <param name="d"></param>
    Public Function Promotefiles(d As DirectoryInfo) As Boolean
        If Not IO.Directory.Exists(d.FullName) Then
            Return False
        End If
        For Each file In d.EnumerateFiles
            If Len(file.FullName) <= 247 Then
                Try
                    My.Computer.FileSystem.MoveFile(file.FullName, d.Parent.FullName & "\" & file.Name)
                Catch ex As IOException
                    Continue For


                End Try
            End If

        Next

        Return True
    End Function
    Public Function DeleteEmptyFolders(d As DirectoryInfo, blnRecurse As Boolean) As Boolean



        For Each di In d.EnumerateDirectories
            ProgressBarOn(d.EnumerateDirectories.Count)
            If blnRecurse Then DeleteEmptyFolders(di, True)
            If di.EnumerateDirectories.Count = 0 And di.EnumerateFiles.Count = 0 Then
                di.Delete()
                'TODO What to do if node doesn't exist?
                MainForm.tvMain2.RemoveNode(di.FullName)
                ProgressIncrement(1)
            End If

        Next

        Return True
    End Function

    Public Function FolderCount(d As DirectoryInfo, count As Integer, blnRecurse As Boolean) As Long
        Try
            count = count + d.EnumerateDirectories.Count

        Catch ex As UnauthorizedAccessException
            Return 0
            Exit Function
        End Try
        If blnRecurse Then
            For Each di In d.EnumerateDirectories
                count = FolderCount(di, count, True)
            Next
        End If
        Return count
    End Function
    Public Function FileCount(d As DirectoryInfo, count As Integer, blnRecurse As Boolean) As Long
        Try
            count = count + d.EnumerateFiles.Count


        Catch ex As Exception
            Return 0
            Exit Function
        End Try
        If blnRecurse Then
            For Each di In d.EnumerateDirectories
                count = FileCount(di, count, True)
            Next
        End If
        Return count
    End Function
    ''' <summary>
    ''' Takes files in d and places them in target, recursively for subfolders if selected
    ''' </summary>
    ''' <param name="Directory"></param>
    ''' <param name="Target"></param>
    ''' <param name="Recurse"></param>
    Public Sub HarvestFolder(Directory As DirectoryInfo, Target As DirectoryInfo, Recurse As Boolean)
        'HarvestFolder(d, blnRecurse)
        'Exit Sub
        Dim i
        i = InputBox("Harvest folders with no more than how many files in?")
        If i = "" Then i = 0

        Dim s As New List(Of String) '= Nothing
        ' s.Add("Test")
        s = AppendToListFromFolder(s, Directory, i)

        If Recurse Then
            For Each di In Directory.EnumerateDirectories
                s = AppendToListFromFolder(s, di, i)
            Next
        End If
        blnSuppressCreate = True 'Prevent request make folder for plural files
        MoveFiles(s, Target.FullName, MainForm.lbxFiles)
        blnSuppressCreate = False
        DeleteEmptyFolders(Directory, True)

    End Sub
    Public Sub HarvestFolder(d As DirectoryInfo, Recurse As Boolean)
        If Recurse Then

            For Each di In d.EnumerateDirectories
                HarvestFolder(di, Recurse)
            Next
        End If
        blnSuppressCreate = True
        For Each f In d.EnumerateFiles
            f.MoveTo(Media.MediaDirectory & "\" & f.Name)
        Next
        DeleteEmptyFolders(d, True)
    End Sub
    ''' <summary>
    ''' Adds all the files in d to s, provided there are no more than icount, or icount is zero
    ''' If icount is zero, then all files are added.
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="d"></param>
    ''' <param name="icount"></param>
    ''' <returns></returns>
    Private Function AppendToListFromFolder(ByVal s As List(Of String), d As DirectoryInfo, icount As Short) As List(Of String)
        Dim l As New List(Of String)
        l = s
        If d.EnumerateFiles.Count <= icount Or icount = 0 Then
            For Each file In d.EnumerateFiles
                l.Add(file.FullName)
            Next
        End If
        Return s
    End Function

    Public Sub BurstFolder(d As DirectoryInfo)
        HarvestFolder(d, d.Parent, True)
    End Sub

End Module
