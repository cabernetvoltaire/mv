Option Explicit On
Imports System.IO

Public Module General
    Public Property blnSecondScreen As Boolean = True
    Public Property LastShowList As String
    Public Property blnLink As Boolean
    Public Property blnMoveMode As Boolean = False
    Public Property lastselection As String
    Public Property blnJumpToMark As Boolean = True


    Public Sub ProgressBarOn(max As Long)
        With frmMain.TSPB
            .Value = 0
            .Maximum = max 'Math.Max(lngListSizeBytes, 100)
            .Visible = True
        End With

    End Sub
    Public Sub ProgressBarOff()
        With frmMain.TSPB
            .Visible = False
        End With

    End Sub
    Public Function LinkTarget(str As String) As String
        Return CreateObject("WScript.Shell").CreateShortcut(str).TargetPath
    End Function
    Public Sub ProgressIncrement(st As Integer)
        With frmMain.TSPB
            '   .Maximum = max
            .Value = (.Value + st) Mod .Maximum 'TODO this could be causing the stalling

        End With
        frmMain.Update()
    End Sub

    Public Function ImageOrientation(ByVal img As Image) As ExifOrientations
        ' Get the index of the orientation property.
        Dim orientation_index As Integer = Array.IndexOf(img.PropertyIdList, OrientationId)

        ' If there is no such property, return Unknown.
        If (orientation_index < 0) Then Return ExifOrientations.Unknown

        ' Return the orientation value.
        Return DirectCast(img.GetPropertyItem(OrientationId).Value(0), ExifOrientations)
    End Function

    Public strPlayOrder() As String = {"Original", "Random", "Name", "Path Name", "Date/Time", "Size", "Type"}
    Public Property lngListSizeBytes As Long
    Public btnDest() As Button = {frmMain.btn1, frmMain.btn2, frmMain.btn3, frmMain.btn4, frmMain.btn5, frmMain.btn6, frmMain.btn7, frmMain.btn8}

    Public lblDest() As Label = {frmMain.lbl1, frmMain.lbl2, frmMain.lbl3, frmMain.lbl4, frmMain.lbl5, frmMain.lbl6, frmMain.lbl7, frmMain.lbl8}
    Public Enum PlayOrder As Byte
        Original
        Random
        Name
        PathName
        Time
        Length
        Type
    End Enum
    Public lngShowlistLines As Long = 0

    Public blnRandomStartPoint = False
    Public PlaybackSpeed As Double = 1
    Public lngInterval = 50
    Public lngMediaDuration As Long
    Public lngMark As Long
    Public iPropjump As Integer = 15
    Public iQuickJump As Integer = 20
    Public strCurrentFilePath As String = ""
    Public strExt As String
    Public Property CurrentFolderPath As String = "E:\"
    Public FileboxContents As New List(Of String)
    Public FBCShown As New List(Of Boolean)
    Public blnDontShowRepeats As Boolean = True
    Public lCurrentDisplayIndex As Long = 0
    Public fType As Filetype
    Public blnRandom As Boolean = False
    Public Showlist As New List(Of String)
    Public Oldlist As New List(Of String)

    Public Sublist As New List(Of String)
    Public currentPicBox As New PictureBox
    Public Autozoomrate As Decimal = 0.4
    Public strVisibleButtons(8) As String

    Public blnButtonsLoaded As Boolean = False
    Public ssspeed As Integer = 200

    Public blnFullScreen As Boolean

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
    Public SlideShowSpeeds() As Integer = {50, 100, 500, 1000, 2000, 9000}
    Public VideoSpeeds() As Integer = {10, 20, 50, 75, 100}
    Public Function TimeOperation(blnStart As Boolean) As TimeSpan
        Static StartTime As Date
        If blnStart Then
            StartTime = Now
            Return Now - StartTime
        Else
            Return Now - StartTime
        End If
    End Function

    ''' <summary>
    ''' Copies list from sorted list2
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="list2"></param>
    Public Sub CopyList(list As List(Of String), list2 As ListBox)
        list.Clear()
        For Each m In list2.Items
            list.Add(m)
        Next
    End Sub

    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of String, String))
        list.Clear()
        For Each m As KeyValuePair(Of String, String) In list2
            list.Add(m.Value)
        Next
    End Sub
    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Long, String))
        list.Clear()
        For Each m As KeyValuePair(Of Long, String) In list2
            list.Add(m.Value)
        Next
    End Sub
    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Date, String))
        list.Clear()
        For Each m As KeyValuePair(Of Date, String) In list2
            list.Add(m.Value)
        Next
    End Sub

    ''' <summary>
    ''' Only load the labels of the current set. 
    ''' </summary>
    Public Sub Buttons_Load()
        For i As Byte = 0 To 7
            lblDest(i).Font = New Font(lblDest(i).Font, FontStyle.Bold)
        Next
        ' blnButtonsLoaded = True
        InitialiseButtons()
    End Sub

    ''' <summary>
    ''' Assigns paths to buttons, moves files, or switches to folders
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Public Sub HandleFunctionKeyDown(sender As Object, e As KeyEventArgs)
        Dim i As Byte = e.KeyCode - Keys.F5
        'Move files
        If e.Control Xor blnMoveMode Then 'Move files if CTRL held
            frmMain.CancelDisplay()

            Select Case PFocus
                Case CtrlFocus.Files
                    MoveFiles(ListfromListbox(frmMain.lbxFiles), strVisibleButtons(i), frmMain.lbxFiles)
                Case CtrlFocus.Tree
                    MoveFolder(CurrentFolderPath, strVisibleButtons(i), frmMain.tvMain2, blnMoveMode)
                Case CtrlFocus.ShowList
                    ' If MsgBox("This will move all the showlist files to the folder " & strVisibleButtons(i) & ". Is this what you want?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes Then Exit Sub
                    MoveFiles(ListfromListbox(frmMain.lbxShowList), strVisibleButtons(i), frmMain.lbxShowList)
            End Select

        Else
            'Assign buttons
            If e.Shift Or strVisibleButtons(i) = "" Then
                AssignButton(i, iCurrentAlpha, CurrentFolderPath) 'Just assign
                If My.Computer.FileSystem.FileExists(strButtonFile) Then
                    KeyAssignmentsStore(strButtonFile)
                Else
                    SaveButtonlist()
                End If
            Else
                'SWITCH folder
                If strVisibleButtons(i) <> CurrentFolderPath Then
                    ChangeFolder(strVisibleButtons(i), True)
                    frmMain.CancelDisplay()
                    frmMain.tvMain2.SelectedFolder = CurrentFolderPath
                End If
            End If

        End If
        frmMain.SetControlColours(blnMoveMode)
    End Sub

    Public Sub ChangeFolder(strPath As String, blnSHow As Boolean)
        If strPath = CurrentFolderPath Then
        Else
            If Not LastFolder.Contains(CurrentFolderPath) Then
                LastFolder.Push(CurrentFolderPath)

                'If blnSHow Then frmMain.tvMain2.SelectedFolder = CurrentFolderPath
            End If
            CurrentFolderPath = strPath 'Switch to this folder
        End If

        frmMain.SetControlColours(blnMoveMode)
    End Sub

    Public Function ListfromListbox(lbx As ListBox) As List(Of String)
        Dim s As New List(Of String)
        For Each l In lbx.SelectedItems
            s.Add(l)
        Next
        Return s
    End Function
    ''' <summary>
    ''' Fill a listbox with a list, according to a filter
    ''' </summary>
    ''' <param name="lbx"></param>
    ''' <param name="Filter"></param>
    ''' <param name="lst"></param>
    Public Sub FillShowbox(lbx As ListBox, Filter As Byte, ByVal lst As List(Of String))
        If lst.Count = 0 Then Exit Sub
        ProgressBarOn(lst.Count)
        If lbx.Equals(frmMain.lbxShowList) Then
            frmMain.CollapseShowlist(False)
        Else
            lbx.Items.Clear()

        End If


        For Each s In lst
            lbx.Items.Add(s)
            ProgressIncrement(1)
        Next
        '        lbx.TabStop = True
        ProgressBarOff()
        frmMain.UpdateFileInfo()
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
                    Case PlayOrder.Name
                        Dim l As Long = 0
                        Dim s As String
                        s = file.Name & Str(l)
                        While NewListS.ContainsKey(s)
                            l += 1
                            s = file.Name & Str(l)
                            '               frmMain.ListBox1.Items.Add(s)

                        End While
                        NewListS.Add(s, file.FullName)
                    Case PlayOrder.Length
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
                    Case PlayOrder.Time
                        Dim time As DateTime = file.LastWriteTime
                        Dim time2 As DateTime = file.LastAccessTime
                        If time2 < time Then time = time2
                        'MsgBox(time)
                        While NewListD.ContainsKey(time)
                            time = time.AddSeconds(1)
                        End While
                        NewListD.Add(time, file.FullName)
                    Case PlayOrder.PathName
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

                    Case PlayOrder.Type
                        NewListS.Add(file.Extension & file.Name & Str(Rnd() * (100)), file.FullName)

                    Case PlayOrder.Random
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
    Public Function SelectFromListbox(lbx As ListBox, s As String) As List(Of String)
        Dim ls As New List(Of String)
        Dim i As Long
        lbx.SelectedItem = Nothing
        lbx.SelectionMode = SelectionMode.MultiExtended
        For i = 0 To lbx.Items.Count - 1

            If InStr(UCase(lbx.Items(i)), UCase(s)) <> 0 Then
                lbx.SetSelected(i, True)
                ls.Add(lbx.Items(i))
            End If
        Next
        Return ls
    End Function
    Public Function HighlightList(lbx As ListBox, ls As List(Of String))
        lbx.SelectedItems.Clear()
        lbx.Refresh()
        lbx.SelectionMode = SelectionMode.MultiExtended
        For Each f In ls
            Dim i = lbx.FindString(f)
            If i >= 0 Then lbx.SetSelected(i, True)
        Next
    End Function
    Public Function SelectDeadLinks(lbx As ListBox)
        HighlightList(lbx, GetDeadLinks(lbx))
    End Function
    Public Function GetDeadLinks(lbx As ListBox) As List(Of String)
        Dim ls As New List(Of String)
        ls = SelectFromListbox(lbx, ".lnk")
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


    Public Function ReturnListOfDirectories(ByVal list As List(Of String), strPath As String) As List(Of String)
        Dim d As New DirectoryInfo(strPath)
        For Each di In d.EnumerateDirectories
            list.Add(di.Name)
            list = ReturnListOfDirectories(list, di.Name)
        Next
        Return list
    End Function
End Module
