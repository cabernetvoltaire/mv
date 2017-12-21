Imports System.IO
Module FileHandling
    Public blnSuppressCreate As Boolean = False
    Public blnChooseOne As Boolean = False
    Public strVideoExtensions = ".vob .webm.avi.flv.mov.mpeg.mpg.m4v.mkv.mp4.wmv.wav.mp3.3gp"
    Public strPicExtensions = ".jpeg.png.jpg.bmp.gif"

    Public FilePumpList As New List(Of String)

    Dim strFilterExtensions(6) As String
    Public Sub AssignExtensionFilters()
        strFilterExtensions(FilterState.All) = ""
        strFilterExtensions(FilterState.Piconly) = strPicExtensions
        strFilterExtensions(FilterState.PicVid) = strPicExtensions & strVideoExtensions
        strFilterExtensions(FilterState.LinkOnly) = ".lnk"
        strFilterExtensions(FilterState.Vidonly) = strVideoExtensions
        strFilterExtensions(FilterState.NoPicVid) = strPicExtensions & strVideoExtensions & "NOT"
    End Sub
    Public Sub SaveButtonlist()
        Dim path As String
        With frmMain.SaveFileDialog1
            .DefaultExt = "msb"
            .Filter = "Metavisua button files|*.msb|All files|*.*"
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                path = .FileName
            Else
                Exit Sub

            End If
        End With
        KeyAssignmentsStore(path)
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
        'flist.Clear()

        If e.EnumerateFiles.Count = 0 Then Exit Sub
        Try
            For Each f In e.EnumerateFiles
                Dim s As String = LCase(f.Extension)

                Select Case CurrentFilterState
                    Case FilterState.All

                        lbx.Items.Add(f.FullName)
                        'FileboxContents.Add(f.FullName)
                    Case FilterState.NoPicVid

                        If InStr(strVideoExtensions & strPicExtensions, s) <> 0 And Len(s) > 0 Then
                        Else
                            lbx.Items.Add(f.FullName)
                            'FileboxContents.Add(f.FullName)

                        End If

                    Case FilterState.Piconly
                        If InStr(strPicExtensions, s) <> 0 And Len(s) > 0 Then

                            lbx.Items.Add(f.FullName)
                            'FileboxContents.Add(f.FullName)
                        Else
                        End If
                    Case FilterState.LinkOnly
                        If s = ".lnk" Then
                            lbx.Items.Add(f.FullName)
                            'FileboxContents.Add(f.FullName)
                        End If

                    Case FilterState.Vidonly
                        If InStr(strVideoExtensions, s) <> 0 And Len(s) > 0 Then
                            lbx.Items.Add(f.FullName)
                            'FileboxContents.Add(f.FullName)
                        End If

                    Case FilterState.PicVid
                        If InStr(strVideoExtensions & strPicExtensions, s) <> 0 And Len(s) > 0 Then
                            lbx.Items.Add(f.FullName)
                            'FileboxContents.Add(f.FullName)
                        End If
                End Select

            Next
            If lbx.Name = "lbxFiles" Then
                CopyList(FileboxContents, lbx)
            Else
                CopyList(Showlist, lbx)
            End If

            lbx.Tag = e
            'If lbx.Items.Count <> 0 Then
            '    If blnRandom Then
            '        Dim s As Long = lbx.Items.Count - 1
            '        lbx.SelectedIndex = Rnd() * s
            '    Else
            '        lbx.SelectedIndex = 0
            '    End If
            'End If

        Catch ex As ArgumentException
            MsgBox(ex.ToString)
        Catch ex As IOException
            Exit Try
        End Try


    End Sub
    Public Sub DisposeLists(list As Object)
        For Each item In list
            item.Dispose()

        Next
    End Sub
    Public Function FilePump(strFileDest As String, lbx1 As ListBox)

        FilePumpList.Add(strFileDest)
        If FilePumpList.Count = 5 Then

            MoveFiles(FilePumpList, lbx1)
            FilePumpList.Clear()
        End If
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
                MsgBox(ex.Message)
                Exit Sub
            End Try


            '            frmMain.tbLastFile.Text = s & "(" & list.Count & ")"
            ProgressIncrement(40)
            '           frmMain.TSPB.Value = Math.Min(count * 40, frmMain.TSPB.Maximum)
            '          frmMain.Update()
        Loop
        '     frmMain.TSPB.Visible = False

        If lbx.Items.Count <> 0 Then lbx.TabStop = True
        fs.Close()
        lngShowlistLines = Showlist.Count

        If notlist.Count = 0 Then Exit Sub
        If MsgBox(notlist.Count & " files were not found. Remove from list?", vbYesNo, "Metavisua") = MsgBoxResult.Yes Then
            For Each s In notlist
                list.Remove(s)
            Next
            If MsgBox("Re-save list?") Then
                StoreList(list, Dest)
            End If
        End If
    End Sub
    Public Function MakeSubList(list As List(Of String), str As String) As List(Of String)
        Dim s As New List(Of String)
        For Each file In list
            If InStr(UCase(file), UCase(str)) <> 0 Then
                s.Add(file)
            End If
        Next
        Return s
    End Function
    Public Sub MoveFolder(strDir As String, strDest As String, tvw As MasaSam.Forms.Controls.FileSystemTree, blnOverride As Boolean)
        If strDest Is Nothing Then Exit Sub

        If Not blnOverride Then
            If MsgBox("This will move current folder to " & strDest & ". Are you sure?", MsgBoxStyle.OkCancel) <> MsgBoxResult.Ok Then Exit Sub
        End If
        Try
            With My.Computer.FileSystem
                Dim s As String = .GetDirectoryInfo(strDir).Name
                .MoveDirectory(strDir, strDest & "\" & s, FileIO.UIOption.AllDialogs)
                UpdateButton(strDir, strDest & "\" & s)
                tvw.RemoveNode(strDir)
            End With


        Catch ex As Exception
            MsgBox(ex.Message)
                End Try

    End Sub
    Private Sub UpdateButton(strPath As String, strDest As String)
        For i = 0 To 25
            For j = 0 To 7
                Dim s = strButtonFilePath(j, i, 1)
                If s = strPath Then
                    strButtonFilePath(j, i, 1) = strDest
                    Exit For
                    Exit For
                End If
            Next
        Next

    End Sub
    ''' <summary>
    ''' Moves files to strDest, and removes them from lbx1 asking for a subfolder in destination if more than one file. If strDest is empty, files are deleted.
    ''' </summary>
    ''' <param name="files"></param>
    ''' <param name="strDest"></param>
    ''' <param name="lbx1"></param>
    Public Sub MoveFiles(files As List(Of String), strDest As String, lbx1 As ListBox)
        Dim ind As Long = lbx1.SelectedIndex
        If files.Count = 1 Then
            'Try
            '    FileIO.FileSystem.CopyFile(files(0), strDest & "\" & FileIO.FileSystem.GetName(files(0)), FileIO.UIOption.OnlyErrorDialogs)
            '    FileIO.FileSystem.DeleteFile(files(0))
            'Catch ex As Exception
            'End Try

            lbx1.Items.Remove(files(0))
            frmMain.tmrPumpFiles.Enabled = True
            FilePump(files.Item(0) & "|" & strDest, lbx1)
            lbx1.SelectionMode = SelectionMode.One
            If lbx1.Items.Count <> 0 Then lbx1.SetSelected(Math.Max(Math.Min(ind, lbx1.Items.Count - 1), 0), True)

            Exit Sub
        End If
        Dim s As String = strDest 'if strDest is empty then delete

        If files.Count > 1 And strDest <> "" Then
            If Not blnSuppressCreate Then s = CreateNewDirectory(strDest)
        End If
        Dim file As String

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

                'Exit For
                If blnCopyMode Then
                    .CopyFile(m.FullName, spath)
                Else
                    'Deal with existing files
                    Try
                        ' m = Nothing
                        ' currentPicBox.Image = Nothing

                        If Not currentPicBox.Image Is Nothing Then DisposePic(currentPicBox)

                        lbx1.Items.Remove(m.FullName)

                        If strDest = "" Then

                            Deletefile(m.FullName)
                        Else
                            'MsgBox("Start")
                            Try
                                .MoveFile(m.FullName, spath, FileIO.UIOption.AllDialogs)
                            Catch ex As Exception
                                Exit For
                            End Try
                            '.DeleteFile(m.FullName)
                            'MsgBox("Finish")
                        End If


                    Catch ex As IOException
                        Continue For
                        'MsgBox(ex.Message)
                    End Try
                End If
            End With
        Next
        lbx1.SelectionMode = SelectionMode.One
        If lbx1.Items.Count <> 0 Then lbx1.SetSelected(Math.Max(Math.Min(ind, lbx1.Items.Count - 1), 0), True)

    End Sub
    Public Sub MoveFiles(filendest As List(Of String), lbx1 As ListBox)

        Dim ind As Long = lbx1.SelectedIndex

        Dim fd As String
        For Each fd In filendest
            Dim fds() As String
            fds = fd.Split("|")
            Dim file As String = fds(0)
            Dim s As String = fds(1)
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
                If blnCopyMode Then
                    .CopyFile(m.FullName, spath)
                Else
                    'Deal with existing files
                    Try
                        ' m = Nothing
                        ' currentPicBox.Image = Nothing

                        If Not currentPicBox.Image Is Nothing Then DisposePic(currentPicBox)
                        lbx1.Items.Remove(m.FullName)

                        If s = "" Then

                            Deletefile(m.FullName)
                        Else
                            'MsgBox("Start")
                            Try
                                .MoveFile(m.FullName, spath, FileIO.UIOption.AllDialogs, FileIO.UICancelOption.ThrowException)
                            Catch ex As Exception
                                Exit For
                            End Try
                            '.DeleteFile(m.FullName)
                            'MsgBox("Finish")
                        End If


                    Catch ex As IOException
                        Continue For
                        'MsgBox(ex.Message)
                    End Try
                End If
            End With
        Next
        lbx1.SelectionMode = SelectionMode.One
        If lbx1.Items.Count <> 0 Then lbx1.SetSelected(Math.Max(Math.Min(ind, lbx1.Items.Count - 1), 0), True)

    End Sub

    Public Function CreateNewDirectory(strDest As String) As String
        Dim s As String = InputBox("Name of folder to create? (Blank means none)", "Create sub-folder", lastselection)
        s = strDest & "\" & s
        Try
            IO.Directory.CreateDirectory(s)
        Catch ex As IO.DirectoryNotFoundException
        End Try

        Return s
    End Function
    Public Sub AddCurrentType(blnRecurse As Boolean)
        AddFilesToCollection(Showlist, strFilterExtensions(CurrentFilterState), blnRecurse)
        FillShowbox(frmMain.lbxShowList, FilterState.All, Showlist)

    End Sub
    Public Sub Addpics(blnRecurse As Boolean)
        AddFilesToCollection(Showlist, strPicExtensions, blnRecurse)
        FillShowbox(frmMain.lbxShowList, CurrentFilterState, Showlist)
    End Sub
    Public Sub AddFilesToCollection(ByVal list As List(Of String), extensions As String, blnRecurse As Boolean)
        Dim s As String
        Dim d As New DirectoryInfo(CurrentFolderPath)

        s = InputBox("Only include files containing? (Leave empty to add all)")
        If blnChooseOne Then
        Else
            PreFindAllFiles(blnRecurse, d)
        End If
        frmMain.Cursor = Cursors.WaitCursor

        FindAllFilesBelow(d, list, extensions, False, s, blnRecurse, blnChooseOne)

        frmMain.Cursor = Cursors.Default

    End Sub
    Private Sub PreFindAllFiles(blnRecurse As Boolean, d As DirectoryInfo)
        If blnChooseOne Then
            Dim l As Long = FolderCount(d, 0, blnRecurse)
            ProgressBarOn(l)
        Else
            Dim l As Long = FileCount(d, 0, blnRecurse)
            ProgressBarOn(l)
        End If
    End Sub
    Public Sub Deletefile(s As String)

        With My.Computer.FileSystem
            If .FileExists(s) Then
                Try
                    .DeleteFile(s, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

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
    ''' <param name="DontInclude"></param>
    ''' <param name="extensions"></param>
    ''' <param name="blnRemove"></param>
    ''' <param name="strSearch"></param>
    ''' <param name="blnRecurse"></param>
    Public Sub FindAllFilesBelow(d As DirectoryInfo, list As List(Of String), extensions As String, blnRemove As Boolean, strSearch As String, blnRecurse As Boolean, blnOneOnly As Boolean)


        For Each file In d.EnumerateFiles
            Try
                If InStr(LCase(extensions), LCase("NOT")) <> 0 Then
                    If InStr(extensions, LCase(file.Extension)) = 0 And file.Extension <> "" Then
                        'Only include if NOT the given extension
                        If InStr(LCase(file.FullName), LCase(strSearch)) = 0 Or strSearch = "" Then
                            If blnRemove Then
                                list.Remove(file.FullName)
                            Else

                                list.Add(file.FullName)

                            End If

                        End If
                    End If
                Else

                    If InStr(extensions, LCase(file.Extension)) <> 0 And file.Extension <> "" Then 'File has an extension, and an appropriate one
                        If InStr(LCase(file.FullName), LCase(strSearch)) <> 0 Or strSearch = "" Then 'Either search is empty, or matches search
                            If blnRemove Then
                                list.Remove(file.FullName)
                            Else
                                list.Add(file.FullName)

                            End If

                        End If
                    Else
                        If extensions = "" Then
                            If InStr(LCase(file.FullName), LCase(strSearch)) <> 0 Or strSearch = "" Then
                                If blnRemove Then
                                    list.Remove(file.FullName)
                                Else
                                    list.Add(file.FullName)

                                End If
                            End If

                        End If
                    End If
                End If
                If blnOneOnly Then Exit For 'Exits each folder when one file has been found matching the condition.
            Catch ex As PathTooLongException
                MsgBox(ex.Message)
            End Try

            ProgressIncrement(1)
        Next
        ProgressBarOff()
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
    Public Sub FindAllFoldersBelow(d As DirectoryInfo, list As List(Of String), blnRecurse As Boolean, blnNonEmptyOnly As Boolean)


        For Each di In d.EnumerateDirectories
            Try
                If blnNonEmptyOnly AndAlso di.EnumerateFiles.Count > 0 Or Not blnNonEmptyOnly Then
                    list.Add(di.FullName)
                End If

            Catch ex As PathTooLongException
                MsgBox(ex.Message)
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
                frmMain.tvMain2.RemoveNode(di.FullName)
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

        Catch ex As UnauthorizedAccessException
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
    ''' <param name="d"></param>
    ''' <param name="target"></param>
    ''' <param name="blnRecurse"></param>
    Public Sub HarvestFolder(d As DirectoryInfo, target As DirectoryInfo, blnRecurse As Boolean)
        Dim s As New List(Of String)
        For Each file In d.EnumerateFiles

            s.Add(file.FullName)
        Next
        If blnRecurse Then
            For Each di In d.EnumerateDirectories
                HarvestFolder(di, target, True) 'TODO look at this. I'm not sure it works.
                For Each file In di.EnumerateFiles

                    s.Add(file.FullName)
                Next
            Next
        End If
        blnSuppressCreate = True 'Prevent request make folder for plural files
        MoveFiles(s, target.FullName, frmMain.lbxFiles)
        blnSuppressCreate = False
        DeleteEmptyFolders(d, True)
    End Sub
    Public Sub BurstFolder(d As DirectoryInfo)
        HarvestFolder(d, d.Parent, True)
    End Sub

End Module
