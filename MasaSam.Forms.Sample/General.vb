Option Explicit On
Imports System.IO
Imports System.Drawing.Imaging
Public Module General
    Public Enum ExifOrientations As Byte
        Unknown = 0
        TopLeft = 1
        TopRight = 2
        BottomRight = 3
        BottomLeft = 4
        LeftTop = 5
        RightTop = 6
        RightBottom = 7
        LeftBottom = 8
    End Enum

    Public Enum CtrlFocus As Byte
        Tree = 0
        Files = 1
        ShowList = 2
    End Enum

    Public lngShowlistLines As Long = 0

    Public Orientation() As String = {"Unknown", "TopLeft", "TopRight", "BottomRight", "BottomLeft", "LeftTop", "RightTop", "RightBottom", "LeftBottom"}
    Public Enum Filetype As Byte
        Pic
        Movie
        Doc
        Gif
        Xcel
        Browsable
        Link
        Unknown
    End Enum


#Region "Controls"

    Public Sub ProgressBarOn(max As Long)
        With MainForm.TSPB
            .Value = 0
            .Maximum = max 'Math.Max(lngListSizeBytes, 100)
            .Visible = True
        End With

    End Sub
    Public Sub ProgressIncrement(st As Integer)
        With MainForm.TSPB
            '   .Maximum = max
            .Value = (.Value + st) Mod .Maximum
        End With
        'MainForm.Update()
    End Sub
    Public Sub ProgressBarOff()
        With MainForm.TSPB
            .Visible = False
        End With

    End Sub
#End Region

#Region "Links"
    Public Sub SelectDeadLinks(lbx As ListBox)
        HighlightList(lbx, GetDeadLinks(lbx))
    End Sub
    Public Function GetDeadLinks(lbx As ListBox) As List(Of String)
        Dim ls As New List(Of String)
        'ls = SelectFromListbox(lbx, ".lnk", False)
        lbx.SelectedItems.Clear()
        For Each fl In lbx.Items
            If InStr(fl, ".lnk") <> 0 Then
                ls.Add(fl)
            End If
        Next
        Dim deadlinks As New List(Of String)
        For Each f In ls
            If Not LinkTargetExists(f) Then
                deadlinks.Add(f)
            End If
        Next
        Return deadlinks
    End Function
    Public Sub CreateFavourite(Filepath As String)
        CreateLink(Filepath, FavesFolderPath, "")
        Exit Sub
        Dim sh As New ShortcutHandler()
        Dim f As New FileInfo(Filepath)
        sh.Bookmark = Media.Position
        sh.TargetPath = Filepath
        sh.ShortcutPath = FavesFolderPath
        sh.ShortcutName = f.Name
        sh.Create_ShortCut(sh.Bookmark)

    End Sub
    Public Sub CreateLink(Filepath As String, DestinationDirectory As String, Name As String)
        Dim sh As New ShortcutHandler
        Dim f As New FileInfo(Filepath)
        sh.Bookmark = Media.Position

        sh.TargetPath = Filepath
        sh.ShortcutPath = DestinationDirectory
        If Name = "" Then
            sh.ShortcutName = f.Name
        Else
            sh.ShortcutName = Name
        End If
        sh.Create_ShortCut(sh.Bookmark)
        If DestinationDirectory = Media.MediaDirectory Then MainForm.UpdatePlayOrder(False)
    End Sub

    Public Function GetAllFilesBelow(DirectoryPath As String, ByVal FileList As List(Of String))
        If DirectoryPath.Contains("RECYCLE") Then
            Return FileList
            Exit Function
        End If
        Dim m As New DirectoryInfo(DirectoryPath)
        Try
            For Each k In m.EnumerateDirectories
                FileList = GetAllFilesBelow(k.FullName, FileList)
            Next

        Catch ex As System.UnauthorizedAccessException
            Return FileList
            Exit Function
        End Try
        For Each f In m.EnumerateFiles
            FileList.Add(f.FullName)
        Next
        Return FileList
    End Function

    ''' <summary>
    ''' Returns the path of the link defined in str
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Public Function LinkTarget(str As String) As String
        Try
            str = CreateObject("WScript.Shell").CreateShortcut(str).TargetPath
            Return str
        Catch ex As Exception
            Return str
        End Try

    End Function
#End Region


#Region "Rotation Functions"

    Public Function ImageOrientation(ByVal img As Image) As ExifOrientations
        ' Get the index of the orientation property.
        Dim orientation_index As Integer = Array.IndexOf(img.PropertyIdList, OrientationId)

        ' If there is no such property, return Unknown.
        If (orientation_index < 0) Then Return ExifOrientations.Unknown

        ' Return the orientation value.
        Return DirectCast(img.GetPropertyItem(OrientationId).Value(0), ExifOrientations)
    End Function
#End Region




    Public Sub ExtractMetaData(theImage As Image)

        ' Try
        'Create an Image object. 

        'Get the PropertyItems property from image.
        Dim propItems As PropertyItem() = theImage.PropertyItems

            'Set up the display.
            Dim font As New Font("Arial", 10)
            Dim blackBrush As New SolidBrush(Color.Black)
            Dim X As Integer = 0
            Dim Y As Integer = 0

            'For each PropertyItem in the array, display the id, type, and length.
            Dim count As Integer = 0
            Dim propItem As PropertyItem
            Dim des As String = ""

            For Each propItem In propItems
                des = des + vbCrLf & "Property Item " + count.ToString()
                des = des & vbTab & "iD: 0x" & propItem.Id.ToString("x")
                des = des & vbTab & "  type" & propItem.Type.ToString()
                des = des & vbTab & "Length" & propItem.Len.ToString()

                ' e.Graphics.DrawString("Property Item " + count.ToString(),
                'font, blackBrush, X, Y)
                ' Y += font.Height

                ' e.Graphics.DrawString("   iD: 0x" & propItem.Id.ToString("x"),
                'font, blackBrush, X, Y)
                ' Y += font.Height

                ' e.Graphics.DrawString("   type: " & propItem.Type.ToString(),
                'font, blackBrush, X, Y)
                ' Y += font.Height

                ' e.Graphics.DrawString("   length: " & propItem.Len.ToString() &
                ' " bytes", font, blackBrush, X, Y)
                ' Y += font.Height

                count += 1
            Next propItem
            MsgBox(des)
        'MsgBox(PropertyItems(theImage))
        'Catch ex As ArgumentException
        'MessageBox.Show("There was an error. Make sure the path to the image file is valid.")
        'End Try

    End Sub

    Public Function TimeOperation(blnStart As Boolean) As TimeSpan
        Static StartTime As Date
        If blnStart Then
            StartTime = Now
            Return Now - StartTime
        Else
            Return Now - StartTime
        End If
    End Function
#Region "List functions"

    ''' <summary>
    ''' Copies list from a lbx
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="lbx"></param>
    Public Sub CopyList(ByVal list As List(Of String), lbx As ListBox)
        list.Clear()
        For Each m In lbx.Items
            list.Add(m)
        Next
    End Sub
    ''' <summary>
    ''' Copies list from a sorted list2
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="list2"></param>
    Private Sub CopyList(ByVal list As List(Of String), ByVal list2 As SortedList(Of String, String))
        list.Clear()
        For Each m As KeyValuePair(Of String, String) In list2
            list.Add(m.Value)
        Next
    End Sub
    Public Function Duplicatelist(ByVal inList As List(Of String)) As List(Of String)
        Dim out As New List(Of String)
        For Each i In inList
            out.Add(i)
        Next
        Return out
    End Function
    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Long, String))
        list.Clear()
        For Each m As KeyValuePair(Of Long, String) In list2
            list.Add(m.Value)
        Next
    End Sub
    Public Function ListfromListbox(lbx As ListBox) As List(Of String)
        Dim s As New List(Of String)
        For Each l In lbx.SelectedItems
            s.Add(l)
        Next
        Return s
    End Function
    ''' <summary>
    ''' Fill a listbox with a list, ignores the filter - dunno why
    ''' </summary>
    ''' <param name="lbx"></param>
    ''' <param name="Filter"></param>
    ''' <param name="lst"></param>
    '''
    Public Sub FillShowbox(lbx As ListBox, Filter As Byte, ByVal lst As List(Of String))

        If lst.Count = 0 Then Exit Sub
        If lst.Count > 1000 Then
            ProgressBarOn(lst.Count)
        End If

        If lbx.Name = "lbxShowList" Then
            MainForm.CollapseShowlist(False)

        End If

        lbx.Items.Clear()

        For Each s In lst
            lbx.Items.Add(s)
            'ProgressIncrement(1)
        Next
        '        lbx.TabStop = True
        ProgressBarOff()

        'frmMain.FilterShowBox()

        'MainForm.UpdateFileInfo()
    End Sub
    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Date, String))
        list.Clear()
        For Each m As KeyValuePair(Of Date, String) In list2
            list.Add(m.Value)
        Next
    End Sub
#End Region


    Public Sub ReportFault(routinename As String, msg As String, Optional box As Boolean = True)
        If box Then
            MsgBox("Exception in " & routinename & vbCrLf & msg)
        Else
            Console.Write("Exception in " & routinename & vbCrLf & msg)

        End If
    End Sub

    Public Sub ReportTime(str As String)
        Debug.Print(Int(Now().Second) & "." & Int(Now().Millisecond) & " " & str)
    End Sub

    Public Sub ChangeFolder(strPath As String)
        If strPath = Media.MediaDirectory Then
        Else
            If Not LastFolder.Contains(Media.MediaDirectory) Then
                LastFolder.Push(Media.MediaDirectory)

            End If
            MainForm.FNG.Clear()

            Media.MediaDirectory = strPath 'Switch to this folder
            ChangeWatcherPath(Media.MediaDirectory)

            ReDim FBCShown(0)
            NofShown = 0
            '   My.Computer.Registry.CurrentUser.SetValue("File", Media.MediaPath)
        End If

    End Sub
    Public Sub ChangeWatcherPath(path As String)
        Dim d As New DirectoryInfo(path)
        If d.Parent Is Nothing Then
        Else

            MainForm.WatchStart(d.Parent.FullName)
        End If


    End Sub

    Public Function FileLengthCheck(file As String) As Boolean
        Dim m As New FileInfo(file)
        If Len(m.FullName) > 247 Then
            If MsgBox("Filename too long - truncate?", MsgBoxStyle.YesNo, "Filename too long") = MsgBoxResult.Yes Then
                Dim i As Integer = Len(m.FullName)
                Dim l As Integer = Len(m.Directory.FullName)
                If l > 247 Then
                    ReportFault("FileLengthCheck", "Unsuccessful - folder name alone is too long")
                    Return False
                    Exit Function
                Else
                    Dim str As String = Right(m.Name, 247 - l)
                    m.MoveTo(m.Directory.FullName & "\" & str)
                    Return True
                End If
            End If
        End If
        Return True
    End Function



    Public Function SetPlayOrderOld(Order As Byte, ByVal List As List(Of String)) As List(Of String)
        'Return SetPlayOrderNew(Order, List)
        Exit Function

        Dim NewListS As New SortedList(Of String, String)
        Dim NewListL As New SortedList(Of Long, String)
        Dim NewListD As New SortedList(Of DateTime, String)
        For Each f In List
            If Len(f) > 247 Then Continue For
            Dim file As New FileInfo(f)
            'frmMain.ListBox1.BringToFront()

            Try
                Select Case Order
                    Case SortHandler.Order.Name
                        Dim l As Long = 0
                        Dim s As String
                        s = file.Name & Str(l)
                        While NewListS.ContainsKey(s)
                            l += 1
                            s = file.Name & Str(l)
                            '               frmMain.ListBox1.Items.Add(s)

                        End While
                        NewListS.Add(s, file.FullName)
                    Case SortHandler.Order.Size
                        Try
                            Dim l As Long
                            l = file.Length
                            While NewListL.ContainsKey(l)
                                l += 1
                                'MsgBox(l)
                            End While
                            NewListL.Add(l, file.FullName)

                        Catch ex As ArgumentException
                            MsgBox("Fail")

                        End Try
                    Case SortHandler.Order.DateTime
                        'MsgBox(time)
                        Dim time = GetDate(file)
                        While NewListD.ContainsKey(time)
                            time = time.AddSeconds(1)
                        End While
                        NewListD.Add(time, file.FullName)
                    Case SortHandler.Order.PathName
                        Dim l As Long = 0
                        Dim s As String
                        s = file.FullName & Str(l)
                        While NewListS.ContainsKey(s)
                            l += 1
                            s = file.FullName & Str(l)
                            '               frmMain.ListBox1.Items.Add(s)

                        End While
                        '                        MsgBox(file.FullName)
                        NewListS.Add(s, file.FullName)

                    Case SortHandler.Order.Type
                        NewListS.Add(file.Extension & file.Name & Str(Rnd() * (100)), file.FullName)

                    Case SortHandler.Order.Random
                        Dim l As Long
                        l = Int(Rnd() * (100 * List.Count))
                        While NewListS.ContainsKey(Str(l))
                            l = Int(Rnd() * (100 * List.Count))
                            '                       frmMain.ListBox1.Items.Add(l)

                        End While
                        NewListS.Add(Str(l), file.FullName)
                    Case Else

                End Select
            Catch ex As System.ArgumentException 'TODO could do better than this. 
                ReportFault("General.SetPlayOrder", ex.Message)
                Continue For
            Catch ex As IO.FileNotFoundException
                Continue For
            Catch ex As System.IO.PathTooLongException
                Continue For
            End Try
        Next

        If NewListD.Count <> 0 Then
            CopyList(List, NewListD)
        ElseIf NewListS.Count <> 0 Then
            CopyList(List, NewListS)
        ElseIf NewListL.Count <> 0 Then
            CopyList(List, NewListL)
        End If

        If MainForm.PlayOrder.ReverseOrder Then
            List = ReverseListOrder(List)
        End If

        Return List




    End Function

    Public Function SetPlayOrder(Order As Byte, ByVal List As List(Of String)) As List(Of String)
        Dim NewListS As New SortedList(Of String, String)
        Dim NewListL As New SortedList(Of Long, String)
        Dim NewListD As New SortedList(Of DateTime, String)
        'frmMain.ListBox1.BringToFront()
        Try
            Select Case Order
                Case SortHandler.Order.Name
                    For Each f In List
                        If Len(f) > 247 Then Continue For
                        Dim file As New FileInfo(f)

                        Dim l As Long = 0
                        Dim s As String
                        s = file.Name & Str(l)
                        While NewListS.ContainsKey(s)
                            l += 1
                            s = file.Name & Str(l)
                            '               frmMain.ListBox1.Items.Add(s)

                        End While
                        NewListS.Add(s, file.FullName)
                    Next
                Case SortHandler.Order.Size
                    For Each f In List
                        If Len(f) > 247 Then Continue For
                        Dim file As New FileInfo(f)
                        Try
                            Dim l As Long
                            l = file.Length
                            While NewListL.ContainsKey(l)
                                l += 1
                                'MsgBox(l)
                            End While
                            NewListL.Add(l, file.FullName)

                        Catch ex As ArgumentException
                            MsgBox("Fail")

                        End Try
                    Next

                Case SortHandler.Order.DateTime
                    For Each f In List
                        If Len(f) > 247 Then Continue For
                        Dim file As New FileInfo(f)
                        'MsgBox(time)
                        Dim time = GetDate(file)
                        While NewListD.ContainsKey(time)
                            time = time.AddSeconds(1)
                        End While
                        NewListD.Add(time, file.FullName)
                    Next
                Case SortHandler.Order.PathName
                    For Each f In List
                        If Len(f) > 247 Then Continue For
                        Dim file As New FileInfo(f)
                        Dim l As Long = 0
                        Dim s As String
                        s = file.FullName & Str(l)
                        While NewListS.ContainsKey(s)
                            l += 1
                            s = file.FullName & Str(l)
                            '               frmMain.ListBox1.Items.Add(s)

                        End While
                        '                        MsgBox(file.FullName)
                        NewListS.Add(s, file.FullName)
                    Next

                Case SortHandler.Order.Type
                    For Each f In List
                        If Len(f) > 247 Then Continue For
                        Dim file As New FileInfo(f)
                        NewListS.Add(file.Extension & file.Name & Str(Rnd() * (100)), file.FullName)
                    Next

                Case SortHandler.Order.Random
                    For Each f In List
                        If Len(f) > 247 Then Continue For
                        Dim file As New FileInfo(f)

                        Dim l As Long
                        l = Int(Rnd() * (100 * List.Count))
                        While NewListS.ContainsKey(Str(l))
                            l = Int(Rnd() * (100 * List.Count))
                            '                       frmMain.ListBox1.Items.Add(l)

                        End While
                        NewListS.Add(Str(l), file.FullName)
                    Next

                Case Else

            End Select
        Catch ex As System.ArgumentException 'TODO could do better than this. 
            ReportFault("General.SetPlayOrder", ex.Message)
        End Try

        If NewListD.Count <> 0 Then
            CopyList(List, NewListD)
        ElseIf NewListS.Count <> 0 Then
            CopyList(List, NewListS)
        ElseIf NewListL.Count <> 0 Then
            CopyList(List, NewListL)
        End If

        If MainForm.PlayOrder.ReverseOrder Then
            List = ReverseListOrder(List)
        End If

        Return List




    End Function

    Function GetDate(f As FileInfo) As DateTime
        Dim time As DateTime = f.CreationTime
        Dim time2 As DateTime = f.LastAccessTime
        Dim time3 As DateTime = f.LastWriteTime
        If time2 < time Then time = time2
        If time3 < time Then time = time3
        Return time
    End Function




    Public Function LoadImage(fname As String) As Image
        Dim FileStream1 As New System.IO.FileStream(fname, IO.FileMode.Open, IO.FileAccess.Read)
        Try
            Dim MyImage As Image = Image.FromStream(FileStream1)
            FileStream1.Close()
            FileStream1.Dispose()
            Return MyImage
        Catch ex As System.ArgumentException
            FileStream1.Close()
            FileStream1.Dispose()
            Return Nothing
        End Try


    End Function
    Public Function SelectFromListbox(lbx As ListBox, s As String, blnRegex As Boolean) As List(Of String)
        Dim ls As New List(Of String)
        Dim i As Long
        lbx.SelectedItem = Nothing
        lbx.SelectionMode = SelectionMode.MultiExtended
        For i = 0 To lbx.Items.Count - 1
            If blnRegex Then
                Dim r As New System.Text.RegularExpressions.Regex(s)
                If r.Matches(lbx.Items(i)).Count > 0 Then
                    lbx.SetSelected(i, True)
                    ls.Add(lbx.Items(i))
                End If
            Else
                If InStr(UCase(lbx.Items(i)), UCase(s)) <> 0 Then
                    lbx.SetSelected(i, True)
                    ls.Add(lbx.Items(i))
                End If
            End If
        Next
        Return ls
    End Function

    Public Sub HighlightList(lbx As ListBox, ls As List(Of String))
        lbx.SelectedItems.Clear()
        lbx.Refresh()
        lbx.SelectionMode = SelectionMode.MultiExtended
        For Each f In ls
            Dim i = lbx.FindString(f)
            If i >= 0 Then lbx.SetSelected(i, True)
        Next
    End Sub
    Public Function ReverseListOrder(m As List(Of String)) As List(Of String)
        Dim k As New List(Of String)
        For Each x In m
            k.Insert(0, x)

        Next
        Return k
    End Function

    Public Function ReturnListOfDirectories(ByVal list As List(Of String), strPath As String) As List(Of String)
        Dim d As New DirectoryInfo(strPath)
        For Each di In d.EnumerateDirectories
            list.Add(di.Name)
            list = ReturnListOfDirectories(list, di.Name)
        Next
        Return list
    End Function
End Module
