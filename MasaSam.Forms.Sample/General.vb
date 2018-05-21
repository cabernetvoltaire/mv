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
    'Public Enum PlayOrder As Byte
    '    Original
    '    Random
    '    Name
    '    PathName
    '    Time
    '    Length
    '    Type
    'End Enum


    Public lngShowlistLines As Long = 0

    Public Orientation() As String = {"Unknown", "TopLeft", "TopRight", "BottomRight", "BottomLeft", "LeftTop", "RightTop", "RightBottom", "LeftBottom"}
    Public Enum Filetype As Byte
        Pic
        Movie
        Doc
        Gif
        Xcel
        Browsable
        Unknown
    End Enum


#Region "Controls"

    Public Sub ProgressBarOn(max As Long)
        With frmMain.TSPB
            .Value = 0
            .Maximum = max 'Math.Max(lngListSizeBytes, 100)
            .Visible = True
        End With

    End Sub
    Public Sub ProgressIncrement(st As Integer)
        With frmMain.TSPB
            '   .Maximum = max
            .Value = (.Value + st) Mod .Maximum 'TODO this could be causing the stalling

        End With
        frmMain.Update()
    End Sub
    Public Sub ProgressBarOff()
        With frmMain.TSPB
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
        ls = SelectFromListbox(lbx, ".lnk", False)
        lbx.SelectedItems.Clear()
        Dim deadlinks As New List(Of String)
        For Each f In ls
            Dim Finfo = New FileInfo(CreateObject("WScript.Shell").CreateShortcut(f).TargetPath)

            If Not Finfo.Exists Then
                deadlinks.Add(f)
            End If
        Next
        Return deadlinks
    End Function

    ''' <summary>
    ''' Returns the path of the link defined in str
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Public Function LinkTarget(str As String) As String
        Try
            Return CreateObject("WScript.Shell").CreateShortcut(str).TargetPath
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
            frmMain.CollapseShowlist(False)

        End If

        lbx.Items.Clear()

        For Each s In lst
            lbx.Items.Add(s)
            ProgressIncrement(1)
        Next
        '        lbx.TabStop = True
        ProgressBarOff()
        'frmMain.UpdateFileInfo()
    End Sub
    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Date, String))
        list.Clear()
        For Each m As KeyValuePair(Of Date, String) In list2
            list.Add(m.Value)
        Next
    End Sub
#End Region
    Public Function FindType(file As String) As Filetype
        blnLink = False
        Try
            Dim info As New FileInfo(file)
            If info.Extension = "" Then
                Return Filetype.Unknown
                Exit Function
            End If
            'If it's a .lnk, find the file
            If LCase(info.Extension) = ".lnk" Then
                blnLink = True
                strCurrentFilePath = LinkTarget(info.FullName) ' CreateObject("WScript.Shell").CreateShortcut(info.FullName).TargetPath
                Try
                    If My.Computer.FileSystem.FileExists(strCurrentFilePath) Then

                        info = New FileInfo(strCurrentFilePath)
                    Else
                        'TODO: Ask whether to delete link, or fix it. 
                        Return Filetype.Unknown
                        Exit Function
                    End If

                Catch ex As Exception
                End Try
            End If
            strExt = LCase(info.Extension)
            'Select Case LCase(strExt)
            If InStr(strVideoExtensions, strExt) <> 0 Then
                Return Filetype.Movie
            ElseIf InStr(strPicExtensions, strExt) <> 0 Then
                Return Filetype.Pic
            ElseIf InStr(".txt.prn.sty.doc", strExt) <> 0 Then
                Return Filetype.Doc
            Else
                Return Filetype.Unknown


            End If
        Catch ex As PathTooLongException
            Return Filetype.Unknown
        End Try
    End Function




    Public Sub ChangeFolder(strPath As String, blnSHow As Boolean)

        If strPath = CurrentFolderPath Then

        Else
            If Not LastFolder.Contains(CurrentFolderPath) Then
                LastFolder.Push(CurrentFolderPath)

            End If
            CurrentFolderPath = strPath 'Switch to this folder
            ' FilterMoveFiles(strPath)
        End If
        ChangeWatcherPath(CurrentFolderPath)
        ReDim FBCShown(0)
        NofShown = 0
        'frmMain.SetControlColours(blnMoveMode)
    End Sub
    Public Sub ChangeWatcherPath(path As String)
        Dim d As New DirectoryInfo(path)
        If d.Parent Is Nothing Then
        Else

            frmMain.WatchStart(d.Parent.FullName)
        End If


    End Sub


    Public Sub MouseHoverInfo(lbx As ListBox, tt As ToolTip)
        Dim cc, sc As Point
        cc = New Point
        sc = New Point
        sc = lbx.MousePosition()
        cc = lbx.PointToClient(sc)
        Dim i As Int32 = lbx.IndexFromPoint(cc)
        If i >= 0 AndAlso i < lbx.Items.Count - 1 Then tt.SetToolTip(lbx, New FileInfo(lbx.Items(i)).Length)
    End Sub
    Public Function SetPlayOrder(Order As Byte, ByVal List As List(Of String)) As List(Of String)
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
                        Dim time As DateTime = file.CreationTime
                        Dim time2 As DateTime = file.LastAccessTime
                        If time2 < time Then time = time2
                        'MsgBox(time)
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

                End Select
            Catch ex As System.ArgumentException 'TODO could do better than this. 
                MsgBox(ex.Message)
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


        Return List




    End Function
    ''' <summary>
    ''' Removes an item from a listbox, and its associated list, advances selected to next. 
    ''' </summary>
    ''' <param name="item"></param>
    ''' <param name="lbx"></param>
    ''' <param name="lst"></param>
    Public Sub RemoveFromListBox(item As String, lbx As ListBox, lst As List(Of String))
        Exit Sub
        If item = "" Then Exit Sub
        Dim s As Integer = lbx.FindString(item)
        If s = "" Then Exit Sub
        lbx.SelectedItem = lbx.Items((s) Mod lbx.Items.Count) 'actually an increment, because of 0 start. 
        lbx.Items.Remove(item)
        lst.Remove(item)


    End Sub

    Public Sub MediaAdvance(wmp As AxWMPLib.AxWindowsMediaPlayer, stp As Long)
        wmp.Ctlcontrols.step(stp)
        wmp.Refresh()


    End Sub
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


    Public Function ReturnListOfDirectories(ByVal list As List(Of String), strPath As String) As List(Of String)
        Dim d As New DirectoryInfo(strPath)
        For Each di In d.EnumerateDirectories
            list.Add(di.Name)
            list = ReturnListOfDirectories(list, di.Name)
        Next
        Return list
    End Function
End Module
