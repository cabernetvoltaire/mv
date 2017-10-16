Imports System.IO
Module FileHandling
    Public strVideoExtensions = ".webm.avi.flv.mov.mpeg.mpg.m4v.mkv.mp4.wmv.wav.mp3"
    Public strPicExtensions = ".jpeg.png.jpg.bmp.gif"



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
                    Case FilterState.NoPicVid

                        'Select Case s
                        If InStr(strVideoExtensions & strPicExtensions, s) <> 0 And Len(s) > 0 Then
                        Else
                            lbx.Items.Add(f.FullName)
                        End If

                    Case FilterState.Piconly
                        'Select Case s
                        If InStr(strPicExtensions, s) <> 0 And Len(s) > 0 Then

                            lbx.Items.Add(f.FullName)
                        Else
                        End If
                    Case FilterState.LinkOnly
                        Select Case s
                            Case ".lnk"
                                lbx.Items.Add(f.FullName)
                            Case Else
                        End Select

                    Case FilterState.Vidonly
                        'Select Case s
                        If InStr(strVideoExtensions, s) <> 0 And Len(s) > 0 Then
                            lbx.Items.Add(f.FullName)
                        End If


                    Case FilterState.PicVid

                        If InStr(strVideoExtensions & strPicExtensions, s) <> 0 And Len(s) > 0 Then
                            lbx.Items.Add(f.FullName)


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

    Public Sub StoreList(list As List(Of String), Dest As String)
        If Dest = "" Then Exit Sub
        Dim fs As New StreamWriter(New FileStream(Dest, FileMode.Create, FileAccess.Write))
        For Each s In list
            fs.WriteLine(s)
        Next
        fs.Close()
    End Sub
    Public Sub Getlist(list As List(Of String), Dest As String, lbx As ListBox)
        Dim notlist As New List(Of String)
        Dim fs As New StreamReader(New FileStream(Dest, FileMode.OpenOrCreate, FileAccess.Read))
        Do While fs.Peek <> -1
            Dim s As String = fs.ReadLine

            Try
                Dim f As New FileInfo(s)
                If f.Exists Then
                    list.Add(s)
                    lbx.Items.Add(s)
                Else
                    notlist.Add(s)
                End If
            Catch ex As System.IO.PathTooLongException
                Continue Do
            End Try


            frmMain.tsslblLastfile.Text = s & "(" & list.Count & ")"
            frmMain.Update()
        Loop
        If lbx.Items.Count <> 0 Then lbx.TabStop = True
        fs.Close()
        If notlist.Count = 0 Then Exit Sub
        If MsgBox(notlist.Count & " files were not found. Remove from list?", vbYesNo, "Metavisua") = MsgBoxResult.Yes Then
            For Each s In notlist
                list.Remove(s)
            Next
        End If
    End Sub
End Module
