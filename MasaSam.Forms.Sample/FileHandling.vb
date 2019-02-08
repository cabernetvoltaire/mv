Imports System.IO
Imports System.Threading
Imports MasaSam.Forms.Controls
Module FileHandling
    Public blnSuppressCreate As Boolean = False
    Public blnChooseOne As Boolean = False

    Public CurrentfilterState As FilterHandler = MainForm.CurrentFilterState
    Public Random As RandomHandler = MainForm.Random
    Public NavigateMoveState As StateHandler = MainForm.NavigateMoveState
    ' Public FP As New FilePump
    Public Listbox As ListBox = MainForm.lbxFiles
    Public Event FolderMoved(Path As String)
    Public Event FileMoved(Files As List(Of String), lbx As ListBox)
    Public t As Thread
    Public WithEvents MS As New MovieSwapper(MainForm.MainWMP, MainForm.MainWMP2)
    Public Media As MediaHandler = MainForm.Media
    Public fm As New FavouritesMinder("Q:\Favourites")

    Public Sub OnMediaShown(MP As AxWMPLib.AxWindowsMediaPlayer, M As MediaHandler) Handles MS.MediaShown
        MainForm.currentWMP = MP
        Media = M
        Media.StartPoint = MainForm.StartPoint
        Media.MediaJumpToMarker(Media.StartPoint)
        '        MainForm.tmrJumpVideo.Enabled = True

    End Sub
    Public Sub OnMediaLoaded(MP As AxWMPLib.AxWindowsMediaPlayer, M As MediaHandler) Handles MS.LoadedMedia
        MainForm.currentWMP = MP
        Media = M

        '        MainForm.tmrJumpVideo.Enabled = True
    End Sub
    Public Sub OnfileMoved(f As List(Of String), lbx As ListBox)
        MainForm.OnFileMoved(f, lbx)
    End Sub

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

    Private Function FilterLBList(e As DirectoryInfo, ByRef lst As List(Of String)) As List(Of String)
        lst.Clear()

        For Each f In e.EnumerateFiles
            lst.Add(f.FullName)
        Next

        Select Case CurrentfilterState.State
            Case FilterHandler.FilterState.All
            Case FilterHandler.FilterState.NoPicVid
                For Each f In e.EnumerateFiles
                    If InStr(PICEXTENSIONS & VIDEOEXTENSIONS, f.Extension) = 0 And Len(f.Extension) <> 0 Then
                    Else
                        lst.Remove(f.FullName)
                    End If
                Next
            Case FilterHandler.FilterState.LinkOnly

                For Each f In e.EnumerateFiles
                    If f.Extension = ".lnk" Then
                    Else
                        lst.Remove(f.FullName)
                    End If
                Next
            Case FilterHandler.FilterState.Piconly

                For Each f In e.EnumerateFiles
                    If InStr(PICEXTENSIONS, f.Extension) = 0 And Len(f.Extension) <> 0 Then
                        lst.Remove(f.FullName)
                    Else
                    End If
                Next
            Case FilterHandler.FilterState.PicVid

                For Each f In e.EnumerateFiles
                    If InStr(PICEXTENSIONS & VIDEOEXTENSIONS, f.Extension) = 0 And Len(f.Extension) <> 0 Then
                        lst.Remove(f.FullName)
                    Else
                    End If
                Next
            Case FilterHandler.FilterState.Vidonly

                For Each f In e.EnumerateFiles
                    If InStr(VIDEOEXTENSIONS, f.Extension) = 0 And Len(f.Extension) <> 0 Then
                        lst.Remove(f.FullName)
                    Else
                    End If
                Next

        End Select
        Return lst
    End Function

    Public Function FilterListBoxList(e As DirectoryInfo, ByVal lst As List(Of String))
        lst = FilterLBList(e, lst)
        Exit Function
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
    Public Sub LevelAllFolders()
        Dim x As New DirectoryInfo(Media.MediaDirectory)
        For Each m In x.EnumerateDirectories("*", SearchOption.AllDirectories)
            If m.Parent.FullName <> Media.MediaDirectory Then
                MoveFolder(m.FullName, Media.MediaDirectory)
            End If
        Next
    End Sub


    Public Sub MoveFolderNew(Dir As String, Dest As String)
        Dim TargetDir As New DirectoryInfo(Dest)
        Dim SourceDir As New DirectoryInfo(Dir)

        My.Computer.FileSystem.CreateDirectory(TargetDir.FullName & "\" & SourceDir.Name)
        Dim flist As New List(Of String)
        For Each f In SourceDir.GetFiles
            flist.Add(f.FullName)
        Next
        blnSuppressCreate = True
        MoveFiles(flist, TargetDir.FullName & "\" & SourceDir.Name, Listbox)

    End Sub

    Public Sub MoveFolder(strDir As String, strDest As String)
        MoveFolderNew(strDir, strDest)
        Exit Sub
        If strDest = "" Then
            Dim k As New DirectoryInfo(strDir)
            k.Delete(True)

        End If
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
                        Dim flist As New List(Of String)
                        GetFiles(dir, flist)

                        fm.DestinationPath = strDest
                        fm.CheckFiles(flist)

                        .MoveDirectory(strDir, strDest & "\" & s, FileIO.UIOption.OnlyErrorDialogs)


                    Case StateHandler.StateOptions.MoveLeavingLink
                        'Create link directory?
                        .MoveDirectory(strDir, strDest & "\" & s, FileIO.UIOption.OnlyErrorDialogs)

                    Case StateHandler.StateOptions.CopyLink
                        'Creat link directory?

                End Select
                UpdateButton(strDir, strDest & "\" & s) 'todo doesnt handle sub-tree
            End With
        Catch ex As Exception
            MsgBox(ex.Message) '
        End Try

    End Sub

    Private Sub GetFiles(dir As DirectoryInfo, flist As List(Of String))
        For Each m In dir.EnumerateFiles()
            flist.Add(m.FullName)
        Next
        For Each x In dir.EnumerateDirectories
            GetFiles(x, flist)
        Next
    End Sub


    ''' <summary>
    ''' Moves files to strDest, and removes them from lbx1 asking for a subfolder in destination if more than one file. If strDest is empty, files are deleted.
    ''' </summary>
    ''' <param name="files"></param>
    ''' <param name="strDest"></param>
    ''' <param name="lbx1"></param>
    Public Sub MoveFiles(files As List(Of String), strDest As String, lbx1 As ListBox)


        Dim s As String = strDest 'if strDest is empty then delete
        If files.Count > 1 And strDest <> "" Then
            If Not blnSuppressCreate Then s = CreateNewDirectory(MainForm.tvMain2, strDest, True)
        End If
        Select Case NavigateMoveState.State
            Case StateHandler.StateOptions.Copy, StateHandler.StateOptions.CopyLink
            Case Else

                MainForm.CancelDisplay()
        End Select


        t = New Thread(New ThreadStart(Sub() MovingFiles(files, strDest, s)))
        t.IsBackground = True
        t.SetApartmentState(ApartmentState.STA)
        t.Start()

        'Deal with the list box
        RaiseEvent FileMoved(files, lbx1)


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
        If strDest <> "" Then
            Dim dinfo As New IO.DirectoryInfo(strDest)
            If dinfo.Exists = False Then dinfo.Create()
        End If
        Dim file As String = ""

        For Each file In files
            '   If Media.Player.URL = file Then Media.Player.URL = ""

            If Not FileLengthCheck(file) Then Continue For
            Dim m As New IO.FileInfo(file)
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
                            fm.DestinationPath = spath
                            fm.CheckFile(f)
                            m.MoveTo(spath)
                            '    Movelink(f, spath)

                        End If
                    Case StateHandler.StateOptions.MoveLeavingLink
                        'Move, and place link here
                        Dim f As New IO.FileInfo(m.FullName)
                        fm.DestinationPath = spath
                        fm.CheckFile(f)
                        m.MoveTo(spath)
                        CreateLink(m.FullName, Media.MediaDirectory, f.Name)

                    Case StateHandler.StateOptions.CopyLink
                        'Paste a link in destination
                        Dim fpath As New FileInfo(spath)
                        CreateLink(m.FullName, fpath.Directory.FullName, m.Name)
                    Case StateHandler.StateOptions.ExchangeLink
                        'Only works on links
                        Dim sh As New ShortcutHandler

                        If m.Extension = ".lnk" Then
                            Dim f As New IO.FileInfo(LinkTarget(m.FullName))
                            spath = f.FullName
                            f.MoveTo(Media.MediaDirectory & "\" & f.Name)
                            CreateLink(f.FullName, New FileInfo(spath).Directory.FullName, m.Name)
                            m.Delete()
                        End If
                        'Move link target here. 
                        'Create link to this file in link target directory
                        'Delite this link


                End Select
            End With
        Next
        'RaiseEvent FileMoved(files, MainForm.lbxFiles)
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


    Public Sub MoveFiles(list As List(Of String), Destination As String)
        With My.Computer.FileSystem
            For Each f In list
                If Len(f) < 200 Then
                    Dim m = New IO.FileInfo(f)
                    Try
                        .MoveFile(m.FullName, Destination & "\" & m.Name)

                    Catch ex As Exception
                        Continue For
                    End Try
                End If
            Next
        End With
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

        End If
        Return s
    End Function
    Public Sub AddCurrentType(Recurse As Boolean)
        AddFilesToCollection(Showlist, strFilterExtensions(CurrentfilterState.State), Recurse)
        FillShowbox(MainForm.lbxShowList, FilterHandler.FilterState.All, Showlist)

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
                    'If MainForm.lbxFiles.FindString(s) <> -1 Then
                    'MainForm.lbxFiles.Items.Remove(s)
                    'End If
                Catch ex As Exception
                    'Exit Sub
                    'Catch ex As
                End Try
            End If

        End With
    End Sub
    Public Function CountSubFiles(strPath As String) As Integer
        Dim dir As New DirectoryInfo(strPath)
        Dim i As Integer
        If dir.EnumerateDirectories.Count <> 0 Then
            For Each f In dir.EnumerateDirectories
                i = i + CountSubFiles(f.FullName)
            Next
        End If
        i = i + dir.EnumerateFiles.Count
        Return i
    End Function

    Public Function ListSubFiles(strPath As String) As List(Of String)
        Dim dir As New DirectoryInfo(strPath)
        Dim ls As New List(Of String)
        If dir.EnumerateDirectories.Count <> 0 Then
            For Each f In dir.EnumerateDirectories
                For Each n In f.EnumerateFiles
                    ls.Add(n.FullName)
                Next

            Next
        End If
        For Each n In dir.EnumerateFiles
            ls.Add(n.FullName)
        Next
        Return ls
    End Function

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
        '  MsgBox(CountSubFiles(d.FullName) & " files below")
        Dim x As IEnumerable(Of FileInfo)
        If blnRecurse Then
            x = d.EnumerateFiles("*", SearchOption.AllDirectories)
        Else
            x = d.EnumerateFiles()
        End If
        For Each file In x
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
        'If blnRecurse Then

        '    For Each di In d.EnumerateDirectories
        '        Try

        '            FindAllFilesBelow(di, list, extensions, blnRemove, strSearch, blnRecurse, blnOneOnly)
        '        Catch ex As UnauthorizedAccessException
        '            Continue For
        '        Catch ex As DirectoryNotFoundException
        '            Continue For
        '        End Try
        '    Next
        'End If
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

    Public Function DeleteEmptyFolders(d As DirectoryInfo, blnRecurse As Boolean) As Boolean

        If blnRecurse Then
            For Each di In d.EnumerateDirectories

                DeleteEmptyFolders(di, True)
            Next
        End If
        If d.EnumerateDirectories.Count = 0 And d.EnumerateFiles.Count = 0 Then
            Dim s As String = d.Parent.FullName
            MainForm.tvMain2.RemoveNode(d.FullName)
            Try
                d.Delete()

            Catch ex As Exception
                Exit Function
            End Try
            ' MainForm.tvMain2.RefreshTree(s)
        End If


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

    Public Sub HarvestBelow(d As DirectoryInfo)
        For Each di In d.EnumerateDirectories
            BurstFolder(di)
        Next
    End Sub
    Public Sub HarvestFolder(d As DirectoryInfo, Recurse As Boolean, Parent As Boolean)
        If Recurse Then
            For Each di In d.EnumerateDirectories
                HarvestFolder(di, Recurse, Parent)
            Next
        End If
        blnSuppressCreate = True
        Dim l As New List(Of String)
        For Each f In d.EnumerateFiles
            l.Add(f.FullName)
        Next
        If Parent Then
            MoveFiles(l, d.Parent.FullName, MainForm.lbxFiles)
        Else
            MoveFiles(l, Media.MediaDirectory, MainForm.lbxFiles)

            blnSuppressCreate = False
        End If

        'For Each f In d.EnumerateFiles
        '    If Parent Then
        '        Dim m As String = d.Parent.FullName & "\" & f.Name
        '        Dim fi As New FileInfo(m)
        '        If fi.Exists Then
        '        Else
        '            'Use an encapsulated move routine
        '            f.MoveTo(m)
        '        End If
        '    Else
        '        f.MoveTo(Media.MediaDirectory & "\" & f.Name)
        '    End If
        'Next
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

    Public Sub BurstFolder(d As DirectoryInfo)
        HarvestFolder(d, True, True)

    End Sub

End Module
