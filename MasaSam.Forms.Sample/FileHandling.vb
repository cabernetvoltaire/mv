Imports System.IO
Module FileHandling
    Public blnSuppressCreate As Boolean = False
    Public strVideoExtensions = ".webm.avi.flv.mov.mpeg.mpg.m4v.mkv.mp4.wmv.wav.mp3.3gp"
    Public strPicExtensions = ".jpeg.png.jpg.bmp.gif"

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
        If IsNothing(e) Then Exit Sub
        If e.Name = "My Computer" Then Exit Sub
        If lbx.Name = "lbxFiles" Then
            lbx.BackColor = FilterColours(CurrentFilterState)
        End If

        lbx.Items.Clear()
        '   flist.Clear()

        Try
            For Each f In e.EnumerateFiles
                Dim s As String = LCase(f.Extension)

                Select Case CurrentFilterState
                    Case FilterState.All

                        lbx.Items.Add(f.FullName)
                        'FileboxContents.Add(f.FullName)
                    Case FilterState.NoPicVid

                        'Select Case s
                        If InStr(strVideoExtensions & strPicExtensions, s) <> 0 And Len(s) > 0 Then
                        Else
                            lbx.Items.Add(f.FullName)
                            'FileboxContents.Add(f.FullName)

                        End If

                    Case FilterState.Piconly
                        'Select Case s
                        If InStr(strPicExtensions, s) <> 0 And Len(s) > 0 Then

                            lbx.Items.Add(f.FullName)
                            'FileboxContents.Add(f.FullName)
                        Else
                        End If
                    Case FilterState.LinkOnly
                        Select Case s
                            Case ".lnk"
                                lbx.Items.Add(f.FullName)
                                'FileboxContents.Add(f.FullName)
                            Case Else
                        End Select

                    Case FilterState.Vidonly
                        'Select Case s
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
            If lbx.Equals(frmMain.lbxFiles) Then

                CopyList(FileboxContents, lbx)
            Else
                CopyList(Showlist, lbx)

            End If

            lbx.Tag = e
                If lbx.Items.Count <> 0 Then
                If blnRandom Then
                    Dim s As Long = lbx.Items.Count - 1
                    lbx.SelectedIndex = Rnd() * s
                Else
                    lbx.SelectedIndex = 0
                End If
            End If

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
        If Not blnOverride Then
            If MsgBox("This will move current folder to " & strDest & ". Are you sure?", MsgBoxStyle.OkCancel) <> MsgBoxResult.Ok Then Exit Sub
        End If
        Try

                    My.Computer.FileSystem.MoveDirectory(strDir, strDest)
                    tvw.RemoveNode(strDir)
                Catch ex As Exception
                    MsgBox(ex.Message)
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
        Dim s As String = strDest 'if strDest is empty then delete
        If files.Count > 1 And strDest <> "" Then
            If Not blnSuppressCreate Then s = CreateNewDirectory(strDest)
        End If
        Dim file As String
        frmMain.CancelDisplay()

        For Each file In files
            Dim m As New FileInfo(file)
            With My.Computer.FileSystem
                Dim i As Int16 = 0
                Dim spath As String = s & "\" & m.Name
                While .FileExists(spath)
                    Dim x = m.Extension
                    Dim b = InStr(spath, "(")
                    If b = 0 Then
                        Replace(spath, x, "(" & i & ")")
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

                            .MoveFile(m.FullName, spath)
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

        'With current file(s)
        'Move file to destination folder
        'If more than one file
        'Create a new folder to put them in.
    End Sub

    Public Function CreateNewDirectory(strDest As String) As String
        Dim s As String = InputBox("Name of folder to create? (Blank means none)", "Create sub-folder", frmMain.LastSelection)
        s = strDest & "\" & s
        Try
            IO.Directory.CreateDirectory(s)
        Catch ex As IO.DirectoryNotFoundException
        End Try

        Return s
    End Function


    Public Sub AddCurrentType(blnRecurse As Boolean)
        AddFilesToCollection(Showlist, FBCShown, strFilterExtensions(CurrentFilterState), blnRecurse)
        FillShowbox(frmMain.lbxShowList, FilterState.All, Showlist)

    End Sub

    Public Sub Addpics(blnRecurse As Boolean)
        AddFilesToCollection(Showlist, FBCShown, strPicExtensions, blnRecurse)
        FillShowbox(frmMain.lbxShowList, CurrentFilterState, Showlist)
    End Sub

    Public Sub AddFilesToCollection(ByVal list As List(Of String), dontinclude As List(Of Boolean), extensions As String, blnRecurse As Boolean)
        Dim s As String
        Dim d As New DirectoryInfo(CurrentFolderPath)

        s = InputBox("Only include files containing? (Leave empty to add all)")
        PreFindallFiles(blnRecurse, d)

        FindAllFilesBelow(d, list, dontinclude, extensions, False, s, blnRecurse)


    End Sub
    Private Sub PreFindAllFiles(blnRecurse As Boolean, d As DirectoryInfo)
        Dim l As Long = FileCount(d, 0, blnRecurse)
        ProgressBarOn(l)
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
    Public Sub FindAllFilesBelow(d As DirectoryInfo, list As List(Of String), ByRef DontInclude As List(Of Boolean), extensions As String, blnRemove As Boolean, strSearch As String, blnRecurse As Boolean)


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

                                DontInclude.Add(False)
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

                                DontInclude.Add(False)
                            End If

                        End If
                    Else
                        If extensions = "" Then
                            If InStr(LCase(file.FullName), LCase(strSearch)) <> 0 Or strSearch = "" Then
                                If blnRemove Then
                                    list.Remove(file.FullName)
                                Else
                                    list.Add(file.FullName)

                                    DontInclude.Add(False)
                                End If
                            End If

                        End If
                    End If
                End If

            Catch ex As PathTooLongException
                MsgBox(ex.Message)
            End Try

            ProgressIncrement(1)
        Next
        ProgressBarOff()
        If blnRecurse Then

            For Each di In d.EnumerateDirectories
                Try

                    FindAllFilesBelow(di, list, DontInclude, extensions, blnRemove, strSearch, blnRecurse)
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


        ProgressBarOn(1000)

        For Each di In d.EnumerateDirectories
            If blnRecurse Then DeleteEmptyFolders(di, True)
            If di.EnumerateDirectories.Count = 0 And di.EnumerateFiles.Count = 0 Then
                di.Delete()

                ProgressIncrement(1)
            End If

        Next

        Return True
    End Function
    Public Function FileCount(d As DirectoryInfo, count As Integer, blnRecurse As Boolean) As Long

        count = count + d.EnumerateFiles.Count
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
                HarvestFolder(di, target, True)
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
        DeleteEmptyFolders(d, True)
    End Sub

End Module
