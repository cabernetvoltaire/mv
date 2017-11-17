Imports System.IO
Module FileHandling
    Public strVideoExtensions = ".webm.avi.flv.mov.mpeg.mpg.m4v.mkv.mp4.wmv.wav.mp3"
    Public strPicExtensions = ".jpeg.png.jpg.bmp.gif"


    ''' <summary>
    ''' Fills the listbox with files from a given folder, in a given filter state
    ''' </summary>
    ''' <param name="lbx"></param>
    ''' <param name="e"></param>
    ''' <param name="Currentfilterstate"></param>
    ''' <param name="flist"></param>
    ''' <param name="blnRandom"></param>
    Public Sub FillListbox(lbx As ListBox, e As IO.DirectoryInfo, Currentfilterstate As Integer, ByVal flist As List(Of String), blnRandom As Boolean)
        If IsNothing(e) Then Exit Sub
        If e.Name = "My Computer" Then Exit Sub


        Try
            lbx.Items.Clear()
            For Each f In e.EnumerateFiles
                '                If Not IsNothing(f.FullName) Then 'frmMain.FileBoxContents.Add(f.FullName)
                Dim s As String = LCase(f.Extension)

                Select Case Currentfilterstate
                    Case FilterState.All

                        lbx.Items.Add(f.FullName)
                        FileboxContents.Add(f.FullName)
                    Case FilterState.NoPicVid

                        'Select Case s
                        If InStr(strVideoExtensions & strPicExtensions, s) <> 0 And Len(s) > 0 Then
                        Else
                            lbx.Items.Add(f.FullName)
                            FileboxContents.Add(f.FullName)

                        End If

                    Case FilterState.Piconly
                        'Select Case s
                        If InStr(strPicExtensions, s) <> 0 And Len(s) > 0 Then

                            lbx.Items.Add(f.FullName)
                            FileboxContents.Add(f.FullName)
                        Else
                        End If
                    Case FilterState.LinkOnly
                        Select Case s
                            Case ".lnk"
                                lbx.Items.Add(f.FullName)
                                FileboxContents.Add(f.FullName)
                            Case Else
                        End Select

                    Case FilterState.Vidonly
                        'Select Case s
                        If InStr(strVideoExtensions, s) <> 0 And Len(s) > 0 Then
                            lbx.Items.Add(f.FullName)
                            FileboxContents.Add(f.FullName)
                        End If


                    Case FilterState.PicVid

                        If InStr(strVideoExtensions & strPicExtensions, s) <> 0 And Len(s) > 0 Then
                            lbx.Items.Add(f.FullName)
                            FileboxContents.Add(f.FullName)


                        End If
                End Select

            Next
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

        For Each s In list
            fs.WriteLine(s)

        Next
        fs.Close()
    End Sub
    ''' <summary>
    ''' Loads Dest into List, and adds all to lbx
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="Dest"></param>
    ''' <param name="lbx"></param>
    '''
    Public Sub Getlist(list As List(Of String), Dest As String, lbx As ListBox)

        Dim notlist As New List(Of String)
        Dim count As Long = 0
        Dim fs As New StreamReader(New FileStream(Dest, FileMode.OpenOrCreate, FileAccess.Read))
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
                MsgBox(ex.Message, MsgBoxStyle.OkOnly)
                Exit Sub
            End Try


            frmMain.tbLastFile.Text = s & "(" & list.Count & ")"
            ProgressIncrement(40, frmMain.TSPB.Maximum)
            frmMain.TSPB.Value = Math.Min(count * 40, frmMain.TSPB.Maximum)
            frmMain.Update()
        Loop
        frmMain.TSPB.Visible = False

        If lbx.Items.Count <> 0 Then lbx.TabStop = True
        fs.Close()
        lngShowlistLines = Showlist.Count

        If notlist.Count = 0 Then Exit Sub
        If MsgBox(notlist.Count & " files were not found. Remove from list?", vbYesNo, "Metavisua") = MsgBoxResult.Yes Then
            For Each s In notlist
                list.Remove(s)
            Next
        End If
    End Sub
    Public Sub MoveFolder(strDir As String, strDest As String, tvw As MasaSam.Forms.Controls.FileSystemTree)
        If MsgBox("This will move current folder to " & strDest & ". Are you sure?") = MsgBoxResult.Ok Then
            Try

                '    My.Computer.FileSystem.MoveDirectory(strDir,strDest)
                tvw.RemoveNode(strDir)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Moves files to strDest, and removes them from lbx1 asking for a subfolder in destination if more than one file
    ''' </summary>
    ''' <param name="files"></param>
    ''' <param name="strDest"></param>
    ''' <param name="lbx1"></param>
    Public Sub MoveFiles(files As List(Of String), strDest As String, lbx1 As ListBox)
        If files.Count > 1 Then
            Dim s As String
            s = InputBox("Name of folder to create? (Blank means none)", "Create sub-folder", "")
            s = strDest & "\" & s
            Try
                IO.Directory.CreateDirectory(s)
            Catch ex As IO.DirectoryNotFoundException
            End Try
        End If
        Dim file As String
        For Each file In files

            Dim m As New FileInfo(file)
            With My.Computer.FileSystem
                Dim spath As String = strDest & "\" & m.Name
                If .FileExists(spath) Then Exit For
                If blnCopyMode Then
                    .CopyFile(m.FullName, spath)

                Else
                    ' currentPicBox.Image.Dispose()
                    'Deal with existing files
                    Try

                        .MoveFile(m.FullName, spath)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try

                    lbx1.Items.Remove(m.FullName)
                End If
            End With

        Next

        'With current file(s)
        'Move file to destination folder
        'If more than one file
        'Create a new folder to put them in.
    End Sub
    Public Sub AddFilesToCollection(ByVal list As List(Of String), dontinclude As List(Of Boolean), extensions As String, blnRecurse As Boolean)
        Dim s As String

        s = InputBox("Search for?")
        Dim d As New DirectoryInfo(CurrentFolderPath)

        FindAllFilesBelow(d, list, dontinclude, extensions, False, s, blnRecurse)

        lngShowlistLines = list.Count

    End Sub
    Public Sub FindAllFilesBelow(d As DirectoryInfo, list As List(Of String), ByRef DontInclude As List(Of Boolean), extensions As String, blnRemove As Boolean, strSearch As String, blnRecurse As Boolean)
        ProgressBarOn(1000)

        For Each file In d.EnumerateFiles
            Try
                If InStr(extensions, LCase(file.Extension)) <> 0 And file.Extension <> "" Then
                    If InStr(LCase(file.FullName), LCase(strSearch)) <> 0 Or strSearch = "" Then

                        If blnRemove Then
                            list.Remove(file.FullName)
                        Else
                            list.Add(file.FullName)

                            DontInclude.Add(False)
                        End If
                        'tbFiles.Text = file.FullName & " (" & list.Count.ToString & " files.)"

                    End If
                Else
                    If extensions = "" Then
                        If blnRemove Then
                            list.Remove(file.FullName)
                        Else
                            list.Add(file.FullName)

                            DontInclude.Add(False)
                        End If
                    End If
                End If

            Catch ex As PathTooLongException
                MsgBox(ex.Message)
            End Try

            ProgressIncrement(1, 1000)
        Next
        ProgressBarOff()

        For Each di In d.EnumerateDirectories
            Try
                FindAllFilesBelow(di, list, DontInclude, extensions, blnRemove, strSearch, blnRecurse)
            Catch ex As UnauthorizedAccessException
                Continue For
            Catch ex As DirectoryNotFoundException
                Continue For
            End Try
        Next

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
                My.Computer.FileSystem.MoveFile(file.FullName, d.Parent.FullName & "\" & file.Name)
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
                ProgressIncrement(1, 1000)
            End If
        Next

        Return True
    End Function
    Public Sub HarvestFolder(d As DirectoryInfo, blnRecurse As Boolean)
        If blnRecurse Then
            For Each di In d.EnumerateDirectories
                HarvestFolder(di, True)
                Promotefiles(di)
            Next
        End If
        Promotefiles(d)

    End Sub
    ''' <summary>
    ''' Takes files in d and places them in target, recursively for subfolders if selected
    ''' </summary>
    ''' <param name="d"></param>
    ''' <param name="target"></param>
    ''' <param name="blnRecurse"></param>
    Public Sub HarvestFolder(d As DirectoryInfo, target As DirectoryInfo, blnRecurse As Boolean)
        If blnRecurse Then
            For Each di In d.EnumerateDirectories
                HarvestFolder(di, True)
                Dim s As New List(Of String)
                For Each file In di.EnumerateFiles

                    s.Add(file.FullName)
                Next
                MoveFiles(s, target.FullName, frmMain.lbxFiles)
            Next
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

End Module
