Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports AxWMPLib
Imports MasaSam.Forms.Controls


Public Class frmMain
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
    Private Const OrientationId As Integer = &H112
    Public blnSpeedRestart As Boolean = False
    Public iSSpeeds() As Integer = {2000, 500, 50}
    Public iPlaybackSpeed() As Decimal = {0.07, 0.4, 0.65}
    Public currentWMP As New AxWMPLib.AxWindowsMediaPlayer
    Public LastPlayed As New Stack(Of String)
    Public blnAutoAdvanceFolder As Boolean = True




    Public Function ImageOrientation(ByVal img As Image) As ExifOrientations
        ' Get the index of the orientation property.
        Dim orientation_index As Integer = Array.IndexOf(img.PropertyIdList, OrientationId)

        ' If there is no such property, return Unknown.
        If (orientation_index < 0) Then Return ExifOrientations.Unknown

        ' Return the orientation value.
        Return DirectCast(img.GetPropertyItem(OrientationId).Value(0), ExifOrientations)
    End Function
    Public Enum PlayOrder As Byte
        Original
        Random
        Name
        PathName
        Time
        Length
        Type
    End Enum
    Private Sub SaveShowlist()
        Dim path As String
        If SaveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            path = SaveFileDialog1.FileName
            Me.Show()
        End If
        path = SaveFileDialog1.FileName
        FileHandling.StoreList(Showlist, path)
    End Sub
    Public Sub HandleKeys(sender As Object, e As KeyEventArgs)
        Select Case e.KeyCode
            Case KeyEscape
                currentWMP.Ctlcontrols.pause()
                tmrSlideShow.Enabled = False

            Case KeyRandomize
                blnRandom = Not blnRandom
                If blnRandom Then
                    tsslblRandom.Text = "RANDOM"
                Else
                    tsslblRandom.Text = "ORDERED"

                End If


            Case KeyNextFile, KeyPreviousFile
                AdvanceFile(e.KeyCode = KeyNextFile)
                e.Handled = True
                tmrSlideShow.Enabled = False

            Case KeySmallJumpDown, KeySmallJumpUp, KeySmallJumpDownAlt, KeySmallJumpUpAlt
                MediaSmallJump(e)
                e.Handled = True

                Case KeyBigJumpOn, KeyBigJumpBack
                MediaLargeJump(e)
                e.Handled = True
            Case KeyMuteToggle
                currentWMP.settings.mute = Not currentWMP.settings.mute
                e.Handled = True
            Case KeyJumpToPoint
                MediaJumpToMarker()
                tmrJumpVideo.Enabled = True
                e.Handled = True
            Case KeyMarkPoint, KeyMarkPointAlt
                MediaMarker = currentWMP.Ctlcontrols.currentPosition
                e.Handled = True

            Case KeyRotate
                RotatePic(currentPicBox)
            Case KeyJumpAutoT
                JumpRandom(e.Shift)

            Case KeyTraverseTree, KeyTraverseTreeBack
                'tvwMain2_KeyDown(sender, e)
                'TraverseTree(tvwMain2.tvFiles, e.KeyCode = KeyTraverseTree)
            Case KeyToggleSpeed
                With currentWMP
                    If .playState = WMPLib.WMPPlayState.wmppsPaused Or PlaybackSpeed <> 1 Then
                        .Ctlcontrols.play()
                        PlaybackSpeed = 1
                        .settings.rate = PlaybackSpeed
                    Else
                        .Ctlcontrols.pause()
                    End If
                End With
                blnSpeedRestart = True
                tsslblSPEED.Text = "SPEED (" & PlaybackSpeed * 100 & "%)"
            Case KeySpeed1, KeySpeed2, KeySpeed3
                If e.KeyCode = KeySpeed1 Then
                    If e.Control Then 'increase the extremes if Control held
                        If e.Shift Then 'decrease if Shift
                            iSSpeeds(0) = iSSpeeds(0) * 0.9
                        Else
                            iSSpeeds(0) = iSSpeeds(0) / 0.9
                        End If
                    End If

                End If
                If e.KeyCode = KeySpeed3 Then
                    If e.Control Then 'increase the extremes if Control held
                        If e.Shift Then 'decrease if Shift
                            iSSpeeds(2) = iSSpeeds(2) * 0.9
                        Else
                            iSSpeeds(2) = iSSpeeds(2) / 0.9
                        End If
                    End If

                End If

                Dim Choice As Byte = e.KeyCode - KeySpeed1
                PlaybackSpeed = iPlaybackSpeed(Choice) 'TODO Options
                If tmrSlideShow.Enabled Then tmrSlideShow.Interval = iSSpeeds(Choice)
                If Choice = KeyToggleSpeed - KeySpeed1 Then PlaybackSpeed = 1

                currentWMP.settings.rate = PlaybackSpeed
                blnSpeedRestart = True
                tsslblSPEED.Text = "SPEED (" & PlaybackSpeed * 100 & "%)"
                e.Handled = True
            Case KeyFilter 'Cycle through listbox filters
                CurrentFilterState = (CurrentFilterState + 1) Mod (GetMaxValue(Of FilterState)() + 1)
                'MsgBox(CurrentFilterState.ToString)
                FillListbox(lbxFiles, lbxFiles.Tag, CurrentFilterState, Showlist)
                tsslblFilter.Text = UCase(Filterstates(CurrentFilterState))
                e.Handled = True

            Case KeyFullscreen
                GoFullScreen(Not blnFullScreen)
                e.Handled = True

            Case KeyLoopToggle
                currentWMP.settings.setMode("loop", Not currentWMP.settings.getMode("loop"))
                e.Handled = True

            Case KeyJumpToPoint
            Case KeyBackUndo
                'strCurrentFilePath = LastPlayed.Pop
                'tmrPicLoad.Enabled = True

            Case KeyReStartSS
                tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
        End Select
        'e.Handled = True
    End Sub
    Public Sub GoFullScreen(blnGo As Boolean)

        'Dim numofMon As Integer = Screen.AllScreens.Length
        'If numofMon > 1 Then
        '    FullScreen.Bounds = Screen.AllScreens(0).Bounds
        'End If
        If blnGo Then
            SetWMP(FullScreen.FSWMP)
            SetPB(FullScreen.fullScreenPicBox)
            'FullScreen.PictureBox1.Image = FullScreen.fullScreenPicBox.Image
            ' SetPB(FullScreen.PictureBox1)
            FullScreen.Show()
        Else
            SetWMP(MainWMP)
            SetPB(PictureBox1)
            FullScreen.Close()
        End If
        blnFullScreen = Not blnFullScreen
        tmrPicLoad.Enabled = True
    End Sub
    Public Sub AdvanceFile(blnForward As Boolean)
        Dim diff As Integer
        If blnForward Then
            diff = 1
        Else
            diff = -1
        End If
        If blnRandom Then

            SelectRandomToPlay(Showlist)
        Else
            ' MsgBox("Advance the slideshow somehow")
            If Showlist.Count = 0 Then
                'If no showlist, use the filelist to advance. 
                If lbxFiles.Items.Count = 0 Then Exit Sub 'if not filelist, then give up.

                If lbxFiles.SelectedIndex = 0 And Not blnForward Then
                    lbxFiles.SelectedIndex = lbxFiles.Items.Count - 1
                Else
                    lbxFiles.SelectedIndex = (lbxFiles.SelectedIndex + diff) Mod lbxFiles.Items.Count

                End If
                '                StartToShow(Showlist(lCurrentDisplayIndex))
                'showing file handled by lbx change event.
            Else
                'Otherwise, advance the playlist. 
                'TODO what happens when diff is negative and 0 is reached?
                If lCurrentDisplayIndex = 0 And Not blnForward Then
                    lCurrentDisplayIndex = Showlist.Count - 1
                Else

                    lCurrentDisplayIndex = (lCurrentDisplayIndex + diff) Mod Showlist.Count
                End If
                StartToShow(Showlist(lCurrentDisplayIndex))
            End If
        End If

    End Sub
    Private Sub StartToShow(showpath As String)
        strCurrentFilePath = showpath
        tmrPicLoad.Enabled = True
    End Sub

    Public Sub MediaSmallJump(e As KeyEventArgs)
        Dim iJumpFactor As Integer
        Dim blnBack As Boolean = e.KeyCode < (KeySmallJumpUp + KeySmallJumpDown) / 2
        If e.Control Then
            iJumpFactor = 5

        ElseIf e.Shift Then
            'Changes the jumpsize while Shift held
            If Not blnBack Then
                iQuickJump = iQuickJump / 2
            Else
                iQuickJump = iQuickJump * 2
            End If
        Else
            'Ordinary jump
            iJumpFactor = 1
        End If
        If blnBack Then
            NewPosition = currentWMP.Ctlcontrols.currentPosition - iQuickJump / iJumpFactor
        Else
            NewPosition = currentWMP.Ctlcontrols.currentPosition + iQuickJump / iJumpFactor
        End If
        blnRandomStartPoint = False
        tmrJumpVideo.Enabled = True
    End Sub

    Public Sub MediaLargeJump(e As KeyEventArgs)
        Dim iJumpFactor As Integer

        If e.Control Then  'Holding down ctrl reduces the jump distance by a factor
            iJumpFactor = 5
        Else
            iJumpFactor = 1
        End If

        With currentWMP
            .Ctlcontrols.currentPosition = Math.Min(.currentMedia.duration, .Ctlcontrols.currentPosition + .currentMedia.duration * Math.Sign(e.KeyCode - (KeyBigJumpOn + KeyBigJumpBack) / 2) / (iJumpFactor * iPropjump))
        End With


    End Sub
    Public Sub JumpRandom(blnAutoTrail As Boolean)
        If Not blnAutoTrail Then
            blnRandomStartPoint = True
            tmrJumpVideo.Enabled = True
            tsslblSTART.Text = "START:RANDOM"
        Else
            MsgBox("Autotrail")
        End If
    End Sub
    Private Function FindType(file As String) As Filetype
        strCurrentFilePath = file
        Try
            Dim info As New FileInfo(file)
            'If it's a .lnk, find the file
            If info.Extension = ".lnk" Then
                strCurrentFilePath = CreateObject("WScript.Shell").CreateShortcut(info.FullName).TargetPath
                Try

                    info = My.Computer.FileSystem.GetFileInfo(strCurrentFilePath)
                Catch ex As Exception

                End Try
            End If
            strExt = LCase(info.Extension)
            'Select Case LCase(strExt)
            If InStr(strVideoExtensions, strExt) <> 0 Then
                currentWMP.settings.rate = PlaybackSpeed
                Return Filetype.Movie
            ElseIf InStr(strPicExtensions, strExt) <> 0 Then
                'CurrentWMP.Ctlcontrols.paused = True
                Return Filetype.Pic
            Else
                Return Filetype.Unknown


            End If
        Catch ex As PathTooLongException
            Return Filetype.Unknown
        End Try
        '  MessageBox.Show(info.Extension)
    End Function
    Private Sub FillShowbox(lbxShowList As ListBox, all As Byte, showlist As List(Of String))
        lbxShowList.Items.Clear()

        For Each s In showlist
            lbxShowList.Items.Add(s)
        Next
    End Sub
    Private Sub SetWMP(tWMP As AxWindowsMediaPlayer)
        With currentWMP
            tWMP.URL = .URL
            'MsgBox(.Ctlcontrols.currentPositionString)
            tWMP.Ctlcontrols.currentPosition = .Ctlcontrols.currentPosition
            .URL = ""

            '   .Visible = False
            currentWMP = tWMP
            .uiMode = "FULL"
            .stretchToFit = True
            .Visible = True
        End With

    End Sub
    Private Sub SetPB(tPB As PictureBox)
        tPB.Image = currentPicBox.Image
        currentPicBox.Visible = False

        currentWMP.Visible = False
        currentWMP.URL = ""

        tPB.Visible = True
        currentPicBox = tPB

    End Sub

    Private Sub TraverseTree(Tree As TreeView, blnForward As Boolean)

        If Tree.SelectedNode Is Nothing Then Exit Sub

        With Tree.SelectedNode

            If blnForward Then
                If .Nodes.Count <> 0 Then
                    If .Index = 0 Then
                        .FirstNode.Expand()
                    Else
                        .NextNode.Expand()
                    End If

                Else
                    .Parent.NextNode.Expand()
                End If

            End If
        End With
    End Sub
    Private Sub HighlightCurrent(strPath As String)
        Dim finfo As New IO.FileInfo(strPath)
        Dim dr As New DriveInfo(strPath)
        Dim fldr As New DirectoryInfo(strPath)
        'tvwMain2.Collapse()
        tvwMain.Expand(strPath)
        'Tvw.SelectNode(strPath, 0, False)
        ' FillListbox(lbxFiles, fldr, CurrentFilterState, LboxFiles) 'TODO
        'Highlight the file in lbxFiles
        For i = 0 To lbxFiles.Items.Count - 1
            If lbxFiles.Items(i) = strPath Then
                lbxFiles.SetSelected(i, True)
                Exit For
            End If
        Next
        'Highlight the file in lbxShowlist
        For i = 0 To lbxShowList.Items.Count - 1
            If lbxShowList.Items(i) = strPath Then
                lbxShowList.SetSelected(i, True)
                Exit For
            End If
        Next

    End Sub
    Private Sub LoadShowList()
        Dim path As String
        tsslblLastfile.Text = TimeOperation(True).TotalMilliseconds
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            path = OpenFileDialog1.FileName
            Me.Show()
        End If
        If path = "" Then Exit Sub
        Dim finfo As New FileInfo(path)
        Dim size As Long = finfo.Length
        If path <> "" Then
            Getlist(Showlist, path, lbxShowList)
        End If
        Dim time As TimeSpan = TimeOperation(False)
        Dim loadrate As Single = size / time.TotalMilliseconds
        tsslblLastfile.Text = time.TotalMilliseconds.ToString & "to load. (" & loadrate & "bytes/millisecond)"

        tsslblShowfile.Text = path
    End Sub

    Private Sub AddMovies(blnRecurse As Boolean)
        AddFilesToCollection(Showlist, FBCShown, strVideoExtensions, True)
        FillShowbox(lbxShowList, FilterState.All, Showlist)
    End Sub
    Private Sub SelectRandomToPlay(flist As List(Of String))
        If flist.Count = 0 Then Exit Sub
        Randomize()
        Static played As Integer
        Dim Number As Long
        'Number = (Rnd(1) * (flist.Count - 1))
        '    While FBCShown(Number) And flist.Count >= played
        Number = (Rnd(1) * (flist.Count - 1)) 'HACK Danger of infinite loop with small folders
        'End While 'TODO implement the Don't show again feature
        strCurrentFilePath = flist(Number)
        played += 1
        'FBCShown(Number) = True
        tsslblFiles.Text = flist.Count & "/" & FBCShown.Count & "/" & Number
        tsslblFiles.Text = "PLAYING " & Number & " OUT OF " & flist.Count & ". " & played & "ALREADY PLAYED"
        tmrPicLoad.Enabled = True

    End Sub
    Private Function SetPlayOrder(Order As Byte, List As List(Of String)) As List(Of String)
        Dim NewListS As New SortedList(Of String, String)
        Dim NewListL As New SortedList(Of Long, String)
        Dim NewListD As New SortedList(Of Date, String)
        Select Case Order
            Case PlayOrder.Name
                For Each f In List
                    Dim file As New FileInfo(f)
                    NewListS.Add(file.Name & file.FullName, file.FullName)


                Next
            Case PlayOrder.Length
                For Each f In List
                    Dim file As New FileInfo(f)
                    Try
                        NewListL.Add(file.Length + Len(file.FullName), file.FullName)
                    Catch ex As Exception

                    End Try
                Next
            Case PlayOrder.Time
                For Each f In List
                    Dim file As New FileInfo(f)
                    Dim time = file.CreationTimeUtc.AddMilliseconds(Rnd(100))
                    Try
                        NewListD.Add(time, file.FullName)
                    Catch ex As System.ArgumentException 'TODO could do better than this. 
                        Continue For
                    End Try
                Next
            Case PlayOrder.PathName
                For Each f In List
                    Dim file As New FileInfo(f)
                    NewListS.Add(file.FullName, file.FullName)

                Next
            Case PlayOrder.Type
                For Each f In List
                    Dim file As New FileInfo(f)
                    NewListS.Add(file.Extension & file.Name & Str(Rnd(100)), file.FullName)
                Next
            Case PlayOrder.Random
                For Each f In List
                    Try
                        Dim file As New FileInfo(f)
                        NewListS.Add(Str(Rnd(List.Count)), file.FullName)

                    Catch ex As System.ArgumentException
                        Dim file As New FileInfo(f)
                        NewListS.Add(Str(Rnd(List.Count)), file.FullName)

                    Catch ex As System.IO.PathTooLongException
                        Continue For
                    End Try
                Next
        End Select
        If NewListD.Count <> 0 Then
            CopyList(List, NewListD)

        ElseIf NewListS.Count <> 0 Then
            CopyList(List, NewListS)
        ElseIf NewListL.Count <> 0 Then
            CopyList(List, NewListL)

        End If

        Return List




    End Function

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

    Private Sub Addpics(blnRecurse As Boolean)
        AddFilesToCollection(Showlist, FBCShown, strPicExtensions, blnRecurse)
        FillShowbox(lbxShowList, FilterState.All, Showlist)
    End Sub
    Private Sub AddFilesToCollection(ByVal list As List(Of String), dontinclude As List(Of Boolean), extensions As String, blnRecurse As Boolean)
        Dim s As String

        s = InputBox("Search for?")
        Dim d As New System.IO.DirectoryInfo(CurrentFolderPath)

        FindAllFilesBelow(d, list, dontinclude, extensions, False, s, blnRecurse)



    End Sub

    Private Sub RemoveFilesFromCollection(ByVal list As List(Of String), dontinclude As List(Of Boolean), extensions As String)
        Dim d As New System.IO.DirectoryInfo(CurrentFolderPath)
        FindAllFilesBelow(d, list, dontinclude, extensions, True, "", True)
        tsslblFiles.Text = list.Count & " files and " & dontinclude.Count & "nots"


    End Sub
    Public Sub CtrlText(ctl As Control, str As String)
        ctl.Text = str
    End Sub

    'Form Controls
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        PreferencesGet()
        'HighlightCurrent(strCurrentFolderPath)
        tmrPicLoad.Interval = lngInterval
        tmrJumpVideo.Interval = lngInterval / 50
        currentWMP = MainWMP
        currentWMP.stretchToFit = True
        currentWMP.uiMode = "FULL"
        currentWMP.Dock = DockStyle.Fill
        currentPicBox = PictureBox1

    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control
        'GiveKey(e.KeyCode)
        HandleKeys(sender, e)
        ' e.Handled = True
    End Sub
    Private Sub frmMain_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        ShiftDown = e.Shift
        CtrlDown = e.Control
    End Sub
    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Settings.PreferencesSave()
    End Sub

    Private Sub tvwMain2_DirectorySelected(sender As Object, e As DirectoryInfoEventArgs)
        PreferencesSave()
        CurrentFolderPath = e.Directory.FullName
        FillListbox(lbxFiles, e.Directory, CurrentFilterState, Showlist)
        '  ControlSetFocus(lbxFiles)
    End Sub
    Private Sub tvwMain2_NodeSelected(sender As Object, e As TreeViewEventArgs)
        If e.Node.ToolTipText = "My Computer" Then Exit Sub

        If e.Node.ToolTipText = "" Then Exit Sub
        ' MsgBox(e.Node.ToolTipText)
        Dim di = New IO.DirectoryInfo(e.Node.ToolTipText)
        CurrentFolderPath = di.FullName
        FillListbox(lbxFiles, di, CurrentFilterState, Showlist)

    End Sub
    Private Sub tvwMain2_DriveSelected(sender As Object, e As DriveInfoEventArgs)
        If e.Drive.Name = "" Then Exit Sub
        Try
            If e.Drive.DriveFormat = "" Then Exit Sub
        Catch ex As System.IO.IOException
            Exit Sub
        End Try
        '  MsgBox(e.Node.ToolTipText)
        Dim di = New IO.DirectoryInfo(e.Drive.Name)
        CurrentFolderPath = di.FullName
        FillListbox(lbxFiles, di, CurrentFilterState, Showlist)
    End Sub
    Private Sub lbxFiles_SelectedIndexChanged(sender As Object, e As EventArgs)
        'Loads a new media file by triggering the PicLoad Timer

        strCurrentFilePath = lbxFiles.SelectedItem

        tmrPicLoad.Interval = lngInterval
        tmrPicLoad.Enabled = True
    End Sub



    'Timers

    Private Sub tmrInitialise_Tick(sender As Object, e As EventArgs) Handles tmrInitialise.Tick
        ToolStripButton3_Click(Me, e)
        tmrInitialise.Enabled = False
    End Sub
    Private Sub tmrPicLoad_Tick(sender As Object, e As EventArgs) Handles tmrPicLoad.Tick
        fType = FindType(strCurrentFilePath)
        HighlightCurrent(strCurrentFilePath)
        LastPlayed.Push(strCurrentFilePath)
        tsslblLastfile.Text = strCurrentFilePath
        Select Case fType
            Case Filetype.Movie
                If blnRandomStartPoint Then
                    tsslblSTART.Text = "START:RANDOM"
                Else
                    tsslblSTART.Text = "START:NORMAL"

                End If
                currentPicBox.Visible = False

                currentWMP.URL = strCurrentFilePath
                If PlaybackSpeed <> 1 Then currentWMP.settings.rate = PlaybackSpeed
                currentWMP.Visible = True
                tmrSlideShow.Enabled = False 'Slideshow stops if movie. Create separate timer for movie slideshows. 

            Case Filetype.Pic
                Dim img As Image
                If Not currentPicBox.Image Is Nothing Then
                    currentPicBox.Image.Dispose()
                    GC.SuppressFinalize(Me)
                End If
                If Not My.Computer.FileSystem.FileExists(strCurrentFilePath) Then Exit Select

                img = GetImage(strCurrentFilePath)
                'If blnFullScreen Then FullScreen.PictureBox1.Image = img
                tsslblZOOM.Text = UCase("Orientation -" & Orientation(ImageOrientation(img)))
                Select Case ImageOrientation(img)
                    Case ExifOrientations.BottomRight
                        img.RotateFlip(RotateFlipType.Rotate180FlipNone)
                    Case ExifOrientations.RightTop
                        img.RotateFlip(RotateFlipType.Rotate90FlipNone)
                    Case ExifOrientations.LeftBottom
                        img.RotateFlip(RotateFlipType.Rotate270FlipNone)

                End Select

                '  currentPicBox.Image = img

                currentWMP.Visible = False
                currentWMP.URL = ""
                PreparePic(currentPicBox, pbxBlanker, img)
                currentPicBox.Visible = True
                currentWMP.Visible = False
            Case Filetype.Unknown
                tsslblLastfile.Text = "Unhandled file " & strCurrentFilePath
                tmrPicLoad.Enabled = False
                Exit Sub
        End Select
        'MainWMP.fullScreen = blnFullScreen

        Me.Text = "Metavisua - " & strCurrentFilePath
        tmrPicLoad.Enabled = False
    End Sub
    Private Sub tmrJumpVideo_Tick(sender As Object, e As EventArgs) Handles tmrJumpVideo.Tick
        'ControlSetFocus(currentWMP)
        MediaDuration = currentWMP.currentMedia.duration
        If blnRandomStartPoint Then
            NewPosition = (Rnd(1) * (currentWMP.currentMedia.duration))
        End If
        currentWMP.Ctlcontrols.currentPosition = NewPosition
        tmrJumpVideo.Enabled = False
    End Sub


    'Tool strip buttons

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs)
        tvwMain.Expand(strCurrentFilePath)
        'Tvw.SelectNode(strCurrentFilePath, 0, False)
        ToolStripButton6_Click(sender, e)



    End Sub
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles tsbFullscreen.Click
        GoFullScreen(True)
        'FullScreen.Show()
        'SetPB(FullScreen.fullScreenPicBox)
        'SetPB(FullScreen.PictureBox1)
        'tmrPicLoad.Enabled = True
    End Sub
    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs)
        SelectRandomToPlay(Showlist)
    End Sub
    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles tsbClear.Click
        Showlist.Clear()
        lbxShowList.Items.Clear()
        FBCShown.Clear()
        ' tsslblFilter.Text = "0"
    End Sub
    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles tsbAddMovies.Click
        AddMovies(True)
    End Sub
    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        Addpics(True)
    End Sub
    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        blnRandomStartPoint = Not blnRandomStartPoint

    End Sub
    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click

        FindDuplicates.Show()
    End Sub
    Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs) Handles ToolStripButton11.Click
        SaveShowlist()
        'currentPicBox.Tag = currentPicBox.ImageLocation
        'FullScreen.Show()
    End Sub
    Private Sub ToolStripButton12_Click(sender As Object, e As EventArgs) Handles ToolStripButton12.Click
        LoadShowList()
    End Sub
    Private Sub ToolStripButton13_Click(sender As Object, e As EventArgs) Handles ToolStripButton13.Click

    End Sub
    Private Sub ToolStripButton13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripButton13.SelectedIndexChanged
        Dim i As Integer
        i = ToolStripButton13.SelectedIndex
        Showlist = SetPlayOrder(i, Showlist)
        FillShowbox(lbxShowList, FilterState.All, Showlist)
    End Sub
    Private Sub ToolStripButton14_Click(sender As Object, e As EventArgs) Handles TSBRemoveMovies.Click
        RemoveFilesFromCollection(Showlist, FBCShown, strVideoExtensions)
        FillShowbox(lbxShowList, FilterState.All, Showlist)

    End Sub
    Private Sub ToolStripButton15_Click(sender As Object, e As EventArgs) Handles ToolStripButton15.Click
        tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
    End Sub
    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Dim i As Integer
        i = ToolStripComboBox1.SelectedIndex
        ssspeed = CInt(ToolStripComboBox1.Items(i))
        tmrSlideShow.Interval = ssspeed

    End Sub

    Private Sub lbxShowList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxShowList.SelectedIndexChanged, lbxFiles.SelectedIndexChanged
        With sender
            Dim i As Long = .SelectedIndex
            strCurrentFilePath = .Items(i)
            lCurrentDisplayIndex = i
            tmrPicLoad.Enabled = True
        End With
    End Sub





    Private Sub ShowListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowListToolStripMenuItem.Click
        LoadShowList()
    End Sub
    Private Sub tmrSlideShow_Tick(sender As Object, e As EventArgs) Handles tmrSlideShow.Tick
        AdvanceFile(True)
        'tmrSlideShow.Interval = ssspeed



    End Sub
    ' Orientations.



    Private Sub RotatePic(currentPicBox As PictureBox)
        If currentPicBox.Image Is Nothing Then Exit Sub
        With currentPicBox.Image
            .RotateFlip(RotateFlipType.Rotate270FlipNone)
            '.SetPropertyItem(OrientationId).value = ExifOrientations.TopRight
            currentPicBox.Refresh()
            .Save(strCurrentFilePath)
        End With
    End Sub
    Private Sub ToolStripButton16_Click(sender As Object, e As EventArgs) Handles ToolStripButton16.Click
        RotatePic(currentPicBox)
    End Sub
    Private Function StringList(List As List(Of String), strSearch As String) As List(Of String)
        Dim Newlist As New List(Of String)
        For Each s In List
            If InStr(LCase(s), LCase(strSearch)) <> 0 Then
                Newlist.Add(s)
            End If
        Next
        Return Newlist
    End Function

    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_LostFocus(sender As Object, e As EventArgs) Handles ToolStripTextBox1.LostFocus
        Dim str = ToolStripTextBox1.Text
        If str = "" Then
            Showlist = Sublist
        Else
            Sublist = Showlist
            Showlist = StringList(Showlist, ToolStripTextBox1.Text)
        End If

        Showlist = SetPlayOrder(0, Showlist)
        FillShowbox(lbxShowList, FilterState.All, Showlist)

    End Sub

    Private Sub tsbTestShow_Click(sender As Object, e As EventArgs) Handles tsbTestShow.Click
        Test.Show()
    End Sub



    Private Sub SaveListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveListToolStripMenuItem.Click
        SaveShowlist()
    End Sub

    Private Sub IncludingSubfoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IncludingSubfoldersToolStripMenuItem.Click
        'Addpics(True)
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub IncludingSubsetsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IncludingSubsetsToolStripMenuItem.Click
        AddMovies(True)
    End Sub

    Private Sub CurrentOnlyToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CurrentOnlyToolStripMenuItem1.Click
        AddMovies(False)
    End Sub

    Private Sub CurrentOnlyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CurrentOnlyToolStripMenuItem.Click
        Addpics(False)
    End Sub

    Private Sub tvwMain2_Load(sender As Object, e As EventArgs)

    End Sub



    Private Sub tvwMain2_Enter(sender As Object, e As EventArgs) Handles lbxShowList.Enter, lbxFiles.Enter
        sender.backcolor = Color.Aquamarine
    End Sub

    Private Sub tvwMain2_Leave(sender As Object, e As EventArgs) Handles lbxShowList.Leave, lbxFiles.Leave
        sender.backcolor = Color.White

    End Sub

    Private Sub tvwMain2_BackColorChanged(sender As Object, e As EventArgs)
        MsgBox("Yes - " & sender.backcolor.ToString)
    End Sub



    Private Sub FileSystemWatcher1_Changed(sender As Object, e As FileSystemEventArgs)

    End Sub

    Private Sub showButtons_Click(sender As Object, e As EventArgs) Handles showButtons.Click
        Buttons.Show()
    End Sub


    Private Sub MainWMP_PlayStateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles MainWMP.PlayStateChange
        PlaystateChange(sender, e)
    End Sub

    Private Sub AddPicturesBGW_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Addpics(True)
    End Sub

    Private Sub tvwMain2_MouseClick(sender As Object, e As MouseEventArgs)

    End Sub
End Class
