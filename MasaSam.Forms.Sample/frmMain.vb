Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports AxWMPLib
Imports MasaSam
Imports MasaSam.Forms.Controls

Public Class frmMain

    Public Property blnSecondScreen As Boolean
    Public Property LastShowList As String
    Dim newThread As New Threading.Thread(AddressOf LoadShowList)


    Private Sub SaveButtonlist()
        Dim path As String
        With SaveFileDialog1
            .DefaultExt = "msb"
            .Filter = "Metavisua button files|*.msb|All files|*.*"
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                path = .FileName
                Me.Show()
            End If
            path = .FileName
        End With
        KeyAssignmentsStore(path)
    End Sub
    Private Sub SaveShowlist()
        Dim path As String
        With SaveFileDialog1
            .DefaultExt = "msl"
            .Filter = "Metavisua list files|*.msl|All files|*.*"
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                path = .FileName
                Me.Show()
            End If
            path = .FileName
        End With
        StoreList(Showlist, path)
    End Sub
    Public Function LoadButtonList() As String
        Dim path As String = ""
        With OpenFileDialog1
            .DefaultExt = "msb"
            .Filter = "Metavisua button files|*.msb|All files|*.*"

            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                path = .FileName
            End If
        End With
        Return path

    End Function
    Private Sub LoadShowList()
        Dim path As String = ""
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            path = OpenFileDialog1.FileName
            Me.Show()
        End If
        If path = "" Then Exit Sub
        LastShowList = path
        Dim finfo As New FileInfo(path)
        Dim size As Long = finfo.Length
        Dim time As TimeSpan
        Dim loadrate As Single
        lngListSizeBytes = size
        If path <> "" Then
            CollapseShowlist(False)
            tbLastFile.Text = TimeOperation(True).TotalMilliseconds
            ProgressBarOn(lngListSizeBytes)
            Getlist(Showlist, path, lbxShowList)

            time = TimeOperation(False)
            loadrate = size / time.TotalMilliseconds
        End If
        ProgressBarOff()
        tbShowfile.Text = "SHOWFILE LOADED:" & path
    End Sub
    Public Sub HandleKeys(sender As Object, e As KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                HandleFunctionKeyDown(sender, e)
                e.Handled = True
            Case Keys.A To Keys.Z
                ChangeButtonLetter(e)
            Case KeyToggleButtons
                ToggleButtons()
            Case KeyEscape
                currentWMP.Ctlcontrols.pause()
                currentWMP.URL = ""
                tmrSlideShow.Enabled = False
                currentPicBox.Image = Nothing
                'currentPicBox.Image.Dispose()

            Case KeyRandomize
                'Cycle through play orders
                ChosenPlayOrder = (ChosenPlayOrder + 1) Mod (PlayOrder.Type)
                tbRandom.Text = "ORDER:" & UCase(strPlayOrder(ChosenPlayOrder))
                Showlist = SetPlayOrder(ChosenPlayOrder, Showlist)
                FillShowbox(lbxShowList, CurrentFilterState, Showlist)
                FileboxContents = SetPlayOrder(ChosenPlayOrder, FileboxContents)
                FillShowbox(lbxFiles, CurrentFilterState, FileboxContents)


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
                RotatePic(currentPicBox, e.Shift)
            Case KeyJumpAutoT
                If e.Control Then
                    ToggleRandomStartPoint()
                End If
                JumpRandom(e.Shift)

            Case KeyTraverseTree, KeyTraverseTreeBack
                'Handled by Treeview behaviour unless focus is elsewhere. 
                'We want the traverse keys always to work. 
                If PFocus <> CtrlFocus.Tree Then
                    tvMain2_KeyDown(sender, e)
                    e.Handled = True
                End If
               ' TraverseTree(tvMain2.tvFiles, e.KeyCode = KeyTraverseTree)
            Case KeyToggleSpeed
                With currentWMP
                    If .playState = WMPLib.WMPPlayState.wmppsPaused Or PlaybackSpeed <> 1 Then
                        PlaybackSpeed = 1
                        tmrSlowMo.Enabled = False
                        .settings.rate = PlaybackSpeed
                        .Ctlcontrols.play()
                    Else
                        .Ctlcontrols.pause()
                    End If
                End With
                blnSpeedRestart = True
                tbSpeed.Text = "SPEED (" & PlaybackSpeed * 100 & "%)"
            Case KeySpeed1, KeySpeed2, KeySpeed3
                e = SpeedChange(e)
            Case KeyFilter 'Cycle through listbox filters
                CurrentFilterState = (CurrentFilterState + 1) Mod (GetMaxValue(Of FilterState)() + 1)
                'MsgBox(CurrentFilterState.ToString)
                FillListbox(lbxFiles, New DirectoryInfo(CurrentFolderPath), CurrentFilterState, Showlist, blnChooseRandomFile)
                tbFilter.Text = UCase(Filterstates(CurrentFilterState))
                e.Handled = True
            Case KeySelect
                SelectSubList()

            Case KeyFullscreen
                GoFullScreen(Not blnFullScreen)
                e.Handled = True

            Case KeyLoopToggle
                currentWMP.settings.setMode("loop", Not currentWMP.settings.getMode("loop"))
                e.Handled = True

            Case KeyJumpToPoint
            Case KeyBackUndo
                ChangeFolder(LastFolder.Pop, True)
                'strCurrentFilePath = LastPlayed.Pop
                'tmrPicLoad.Enabled = True

            Case KeyReStartSS
                tmrSlideShow.Enabled = Not tmrSlideShow.Enabled

        End Select

        ' e.Handled = True
    End Sub


    'Private Function SpeedChange(e As KeyEventArgs, blnTrue As Boolean)
    '   SetMotion(e.KeyCode) 'Alternative speed. Doesn't work at the moment. 
    'End Function
    Private Function SpeedChange(e As KeyEventArgs) As KeyEventArgs
        If e.KeyCode = KeySpeed1 Then
            If e.Control Then 'increase the extremes if Control held TODO Don't know if this works. 
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
        If (Not currentWMP.playState = WMPLib.WMPPlayState.wmppsPlaying) Then

            tmrSlideShow.Interval = iSSpeeds(Choice)
        End If
        PlaybackSpeed = iPlaybackSpeed(Choice) 'TODO Options

        'TODO This means you can't change Slideshow speed while watching a movie, which makes sense. 

        If Choice = KeyToggleSpeed - KeySpeed1 Then
            PlaybackSpeed = 1
            currentWMP.settings.rate = PlaybackSpeed
            currentWMP.Ctlcontrols.play()
            tmrSlowMo.Enabled = False
        Else
            '        currentWMP.settings.rate = PlaybackSpeed

            currentWMP.Ctlcontrols.pause()
            tmrSlowMo.Interval = 1000 / ((PlaybackSpeed) * 30)
            tmrSlowMo.Enabled = True

        End If
        'Report
        blnSpeedRestart = True
        tbSpeed.Text = "SPEED (" & PlaybackSpeed * 100 & "%)"
        e.Handled = True
        Return e
    End Function

    Private Sub ToggleButtons()
        Buttons_Load()
        blnButtonsLoaded = Not blnButtonsLoaded
        ctrPicAndButtons.Panel2Collapsed = Not blnButtonsLoaded
        LoadCurrentButtonSet()
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
            Dim screen As Screen
            If blnSecondScreen Then 'TODO This doesn't work. Why not?
                screen = Screen.AllScreens(1)
            Else
                screen = Screen.AllScreens(0)
            End If
            FullScreen.StartPosition = FormStartPosition.Manual
            FullScreen.Location = screen.Bounds.Location + New Point(100, 100)
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
            'Advance using whichever control has focus 
            Select Case PFocus
                Case CtrlFocus.Tree, CtrlFocus.Files

                    If lbxFiles.Items.Count = 0 Then Exit Sub 'if not filelist, then give up.

                    If lbxFiles.SelectedIndex = 0 And Not blnForward Then
                        lbxFiles.SelectedIndex = lbxFiles.Items.Count - 1
                    Else
                        lbxFiles.SelectedIndex = (lbxFiles.SelectedIndex + diff) Mod lbxFiles.Items.Count

                    End If

                Case Else
                    'Otherwise, advance the playlist. 
                    If Showlist.Count > 0 Then

                        If lCurrentDisplayIndex = 0 And Not blnForward Then
                            lCurrentDisplayIndex = Showlist.Count - 1
                        Else

                            lCurrentDisplayIndex = (lCurrentDisplayIndex + diff) Mod Showlist.Count
                        End If
                        StartToShow(Showlist(lCurrentDisplayIndex))
                    End If
            End Select
        End If

        'If Main is the current form
        'Perform according to which box has the focus
        'Otherwise
        'Work with lists. 
        'Actually rewrite with lists. 
    End Sub
    Private Sub StartToShow(showpath As String)
        strCurrentFilePath = showpath
        tmrPicLoad.Enabled = True
    End Sub
    Public Sub CollapseShowlist(blnCollapse As Boolean)
        MasterContainer.Panel2Collapsed = blnCollapse
        MasterContainer.SplitterDistance = MasterContainer.Height / 3
        '        ctrTreeandFiles.Panel1Collapsed = Not ctrTreeandFiles.Panel1Collapsed
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
            tbStartpoint.Text = "START:RANDOM"


        Else
            MsgBox("Autotrail")
        End If
    End Sub
    Private Function FindType(file As String) As Filetype
        Try
            Dim info As New FileInfo(file)
            If info.Extension = "" Then
                Return Filetype.Unknown
                Exit Function
            End If
            'If it's a .lnk, find the file
            If info.Extension = ".lnk" Then
                strCurrentFilePath = CreateObject("WScript.Shell").CreateShortcut(info.FullName).TargetPath
                Try

                    info = New FileInfo(strCurrentFilePath)
                Catch ex As Exception

                End Try
            End If
            strExt = LCase(info.Extension)
            'Select Case LCase(strExt)
            If InStr(strVideoExtensions, strExt) <> 0 Then
                ' currentWMP.settings.rate = PlaybackSpeed
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

    Private Sub SetWMP(tWMP As AxWindowsMediaPlayer)
        With currentWMP
            tWMP.URL = .URL
            tWMP.Ctlcontrols.currentPosition = .Ctlcontrols.currentPosition
            .URL = ""
            currentWMP = tWMP
            .stretchToFit = True
            .Visible = True
            .BringToFront()
        End With

    End Sub
    Private Sub SetPB(tPB As PictureBox)
        tPB.Image = currentPicBox.Image
        'currentPicBox.Image.Dispose()
        currentPicBox.Visible = False
        currentWMP.Visible = False
        currentWMP.URL = ""
        tPB.Visible = True
        currentPicBox = tPB

    End Sub

    'Private Sub TraverseTree(Tree As TreeView, blnForward As Boolean)

    '    If Tree.SelectedNode Is Nothing Then Exit Sub

    '    With Tree.SelectedNode

    '        If blnForward Then
    '            If .Nodes.Count <> 0 Then
    '                If .Index = 0 Then
    '                    .FirstNode.Expand()
    '                Else
    '                    .NextNode.Expand()
    '                End If

    '            Else
    '                .Parent.NextNode.Expand()
    '            End If

    '        End If
    '    End With
    'End Sub
    Private Sub HighlightCurrent(strPath As String)
        If strPath = "" Then Exit Sub
        If Len(strPath) > 247 Then Exit Sub
        Dim finfo As New FileInfo(strPath)
        UpdateFileInfo(finfo)
        If InStr(":\", strPath) = Len(strPath) - 2 Then

        Else
            Dim s As String = Path.GetDirectoryName(strPath)
            If tvMain2.SelectedFolder <> s Then tvMain2.SelectedFolder = s 'Only change tree if it needs changing
        End If


        HighlightListboxSelected(strPath, lbxFiles) 'Highlight the current file in lbxFiles
        If Not MasterContainer.Panel2Collapsed Then
            'Highlight the file in lbxShowlist, but only if it has the focus
            If PFocus = CtrlFocus.ShowList Then
                HighlightListboxSelected(strPath, lbxShowList)
            End If
        End If

    End Sub

    Private Sub HighlightListboxSelected(strPath As String, ctrl As ListBox)
        With ctrl
            For i = 0 To .Items.Count - 1
                If .Items(i) = strPath Then
                    .SetSelected(i, True)
                    Exit For
                End If
            Next
        End With

    End Sub




    Private Sub AddMovies(blnRecurse As Boolean)
        AddFilesToCollection(Showlist, FBCShown, strVideoExtensions, blnRecurse)
        FillShowbox(lbxShowList, FilterState.All, Showlist)
    End Sub
    Private Sub AddFiles(blnRecurse As Boolean)
        AddFilesToCollection(Showlist, FBCShown, "", blnRecurse)
        FillShowbox(lbxFiles, FilterState.All, Showlist)
    End Sub
    Private Sub SelectRandomToPlay(flist As List(Of String))
        If flist.Count = 0 Then Exit Sub
        Randomize()
        Static played As Integer
        Dim Number As Long
        'Number = (Rnd(1) * (flist.Count - 1))
        '    While FBCShown(Number) And flist.Count >= played
        'End While 'TODO implement the Don't show again feature
        strCurrentFilePath = flist(Number)
        played += 1
        'FBCShown(Number) = True
        tbFiles.Text = flist.Count & "/" & FBCShown.Count & "/" & Number
        tbFiles.Text = "PLAYING " & Number & " OUT OF " & flist.Count & ". " & played & "ALREADY PLAYED"
        tmrPicLoad.Enabled = True

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


    Private Sub RemoveFilesFromCollection(ByVal list As List(Of String), dontinclude As List(Of Boolean), extensions As String)
        Dim d As New System.IO.DirectoryInfo(CurrentFolderPath)
        FindAllFilesBelow(d, list, dontinclude, extensions, True, "", True)
        tbFiles.Text = list.Count & " files and " & dontinclude.Count & "nots"


    End Sub
    Public Sub CtrlText(ctl As Control, str As String)
        ctl.Text = str
    End Sub

    'Form Controls
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        GlobalInitialise()

    End Sub

    Private Sub GlobalInitialise()
        CollapseShowlist(True)
        PreferencesGet()
        Randomize()
        AssignExtensionFilters()
        tmrLoadLastFolder.Enabled = True

        tmrPicLoad.Interval = lngInterval
        tmrJumpVideo.Interval = lngInterval / 50
        currentWMP = MainWMP
        currentWMP.stretchToFit = True
        currentWMP.uiMode = "FULL"
        currentWMP.Dock = DockStyle.Fill
        currentPicBox = PictureBox1
        'Exit Sub
        Try
            'KeyAssignmentsRestore(strButtonFile)
            If Not blnButtonsLoaded Then ToggleButtons()

        Catch ex As FileNotFoundException

            Exit Try
        Catch ex As DirectoryNotFoundException
            Exit Try
        End Try
    End Sub

    Private Sub LoadDefaultShowList()
        Dim path As String = "Q:\Camgirls.msl"

        Getlist(Showlist, path, lbxShowList)
        CollapseShowlist(False)


    End Sub

    Private Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control
        'GiveKey(e.KeyCode)
        HandleKeys(sender, e)
        'e.Handled = True
        '  MsgBox(sender.ToString & " " & e.ToString)
    End Sub
    Private Sub frmMain_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        ShiftDown = e.Shift
        CtrlDown = e.Control
        ' MsgBox("Ring")
    End Sub
    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Settings.PreferencesSave()
    End Sub

    Private Sub tvMain2_DriveSelected(sender As Object, e As DriveInfoEventArgs)
        If e.Drive.Name = "" Then Exit Sub
        Try
            If e.Drive.DriveFormat = "" Then Exit Sub
        Catch ex As System.IO.IOException
            Exit Sub
        End Try
        '  MsgBox(e.Node.ToolTipText)
        Dim di = New IO.DirectoryInfo(e.Drive.Name)
        ChangeFolder(di.FullName, True)
        FillListbox(lbxFiles, di, CurrentFilterState, Showlist, blnChooseRandomFile)
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
        tbLastFile.Text = strCurrentFilePath
        Select Case fType
            Case Filetype.Movie
                HandleMovie()

            Case Filetype.Pic
                Dim img As Image
                If Not currentPicBox.Image Is Nothing Then
                    DisposePic(currentPicBox)
                End If
                If Not My.Computer.FileSystem.FileExists(strCurrentFilePath) Then Exit Select

                img = GetImage(strCurrentFilePath)
                If img Is Nothing Then Exit Sub
                'If blnFullScreen Then FullScreen.PictureBox1.Image = img
                OrientPic(img)
                'Resume if in middle of slideshow
                If blnRestartSlideShowFlag Then
                    tmrSlideShow.Enabled = True
                    blnRestartSlideShowFlag = False
                End If

                MovietoPic(img)
            Case Filetype.Unknown
                tbLastFile.Text = "Unhandled file " & strCurrentFilePath
                tmrPicLoad.Enabled = False
                Exit Sub
        End Select
        'MainWMP.fullScreen = blnFullScreen

        Me.Text = "Metavisua - " & strCurrentFilePath
        tmrPicLoad.Enabled = False
    End Sub

    Private Sub MovietoPic(img As Image)
        currentWMP.Visible = False
        currentWMP.URL = ""
        PreparePic(currentPicBox, pbxBlanker, img)
        currentPicBox.Visible = True
        currentPicBox.BringToFront()
        currentWMP.Visible = False
    End Sub

    Private Sub OrientPic(img As Image)
        tbZoom.Text = UCase("Orientation -" & Orientation(ImageOrientation(img)))
        Select Case ImageOrientation(img)
            Case ExifOrientations.BottomRight
                img.RotateFlip(RotateFlipType.Rotate180FlipNone)
            Case ExifOrientations.RightTop
                img.RotateFlip(RotateFlipType.Rotate90FlipNone)
            Case ExifOrientations.LeftBottom
                img.RotateFlip(RotateFlipType.Rotate270FlipNone)

        End Select
    End Sub

    Private Sub HandleMovie()
        If blnRandomStartPoint Then
            tbStartpoint.Text = "START:RANDOM"
        Else
            tbStartpoint.Text = "START:NORMAL"

        End If
        currentPicBox.Visible = False

        currentWMP.URL = strCurrentFilePath
        If PlaybackSpeed <> 1 Then currentWMP.settings.rate = PlaybackSpeed
        currentWMP.Visible = True
        currentWMP.BringToFront()

        If tmrSlideShow.Enabled Then
            blnRestartSlideShowFlag = True
            tmrSlideShow.Enabled = False 'Slideshow stops if movie. Create separate timer for movie slideshows. 
        End If
    End Sub

    Private Sub tmrJumpVideo_Tick(sender As Object, e As EventArgs) Handles tmrJumpVideo.Tick
        'ControlSetFocus(currentWMP)
        MediaDuration = currentWMP.currentMedia.duration
        If blnRandomStartPoint Then
            NewPosition = (Rnd(1) * (currentWMP.currentMedia.duration))
        End If
        currentWMP.Ctlcontrols.currentPosition = NewPosition
        tmrJumpVideo.Enabled = False
        If Not blnRandomStartAlways Then blnRandomStartPoint = False
    End Sub


    'Tool strip buttons

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs)
        'tvwMain.Expand(strCurrentFilePath)
        'Tvw.SelectNode(strCurrentFilePath, 0, False)
        ToolStripButton6_Click(sender, e)



    End Sub
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles tsbFullscreen.Click
        GoFullScreen(True)

    End Sub
    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs)
        SelectRandomToPlay(Showlist)
    End Sub
    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles tsbClear.Click
        Showlist.Clear()
        lbxShowList.Items.Clear()
        lbxShowList.TabStop = False
        FBCShown.Clear()
        'CollapseShowlist(True)
        ' tsslblFilter.Text = "0"
    End Sub
    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs)
        AddMovies(True)
    End Sub
    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs)
        Addpics(True)
    End Sub
    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        ToggleRandomStartPoint()
    End Sub

    Private Sub ToggleRandomStartPoint()
        blnRandomStartAlways = Not blnRandomStartAlways
        blnRandomStartPoint = blnRandomStartAlways
    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click

        FindDuplicates.Show()
    End Sub
    Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs)
        SaveShowlist()
        'currentPicBox.Tag = currentPicBox.ImageLocation
        'FullScreen.Show()
    End Sub
    Private Sub ToolStripButton12_Click(sender As Object, e As EventArgs)
        LoadShowList()
        'newThread.Start()
    End Sub
    Private Sub ToolStripButton13_Click(sender As Object, e As EventArgs) Handles ToolStripButton13.Click

    End Sub
    Private Sub ToolStripButton13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripButton13.SelectedIndexChanged
        Dim i As Integer
        i = ToolStripButton13.SelectedIndex
        Showlist = SetPlayOrder(i, Showlist)
        FillShowbox(lbxShowList, FilterState.All, Showlist)
    End Sub
    Private Sub ToolStripButton14_Click(sender As Object, e As EventArgs)
        RemoveFilesFromCollection(Showlist, FBCShown, strVideoExtensions)
        FillShowbox(lbxShowList, FilterState.All, Showlist)

    End Sub
    Private Sub ToolStripButton15_Click(sender As Object, e As EventArgs) Handles ToolStripButton15.Click
        tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
        If tmrSlideShow.Enabled = False Then blnRestartSlideShowFlag = False
    End Sub
    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Dim i As Integer
        i = ToolStripComboBox1.SelectedIndex
        ssspeed = CInt(ToolStripComboBox1.Items(i))
        tmrSlideShow.Interval = ssspeed

    End Sub

    Private Sub Listbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxShowList.SelectedIndexChanged, lbxFiles.SelectedIndexChanged
        With sender
            Dim i As Long = .SelectedIndex
            If .Items.count = 0 Then
                .items.add("If there is nothing showing here, check your filters")
            ElseIf i >= 0 Then
                tmrPicLoad.Enabled = False
                'strOldPath = strCurrentFilePath
                strCurrentFilePath = .Items(i)
                lCurrentDisplayIndex = i
                tmrPicLoad.Enabled = True
            End If
        End With
    End Sub





    Private Sub ShowListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowListToolStripMenuItem.Click

        LoadShowList()
        'newThread.Start()
    End Sub
    Private Sub tmrSlideShow_Tick(sender As Object, e As EventArgs) Handles tmrSlideShow.Tick
        AdvanceFile(True)
        'tmrSlideShow.Interval = ssspeed



    End Sub
    ' Orientations.



    Private Sub RotatePic(currentPicBox As PictureBox, blnLeft As Boolean)
        If currentPicBox.Image Is Nothing Then Exit Sub
        With currentPicBox.Image
            If blnLeft Then
                .RotateFlip(RotateFlipType.Rotate90FlipNone)
            Else
                .RotateFlip(RotateFlipType.Rotate270FlipNone)

            End If
            '.SetPropertyItem(OrientationId).value = ExifOrientations.TopRight
            currentPicBox.Refresh()
            Dim finfo As New FileInfo(strCurrentFilePath)
            Dim dt As New Date
            'avoid the updating of the write time
            dt = finfo.LastWriteTime
            .Save(strCurrentFilePath)
            finfo.LastWriteTime = dt

        End With
    End Sub
    Private Sub ToolStripButton16_Click(sender As Object, e As EventArgs) Handles ToolStripButton16.Click
        RotatePic(currentPicBox, True)
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
        Dim str = ToolStripTextBox1.Text 'TODO fix this
        If str = "" Then
            Showlist = Sublist
        Else
            Sublist = Showlist
            Showlist = StringList(Showlist, ToolStripTextBox1.Text)
        End If

        Showlist = SetPlayOrder(0, Showlist)
        FillShowbox(lbxShowList, FilterState.All, Showlist)

    End Sub

    Private Sub tsbTestShow_Click(sender As Object, e As EventArgs)
        Test.Show()
    End Sub



    Private Sub SaveListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveListToolStripMenuItem.Click
        SaveShowlist()
    End Sub

    Private Sub IncludingSubfoldersToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'Addpics(True)

        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub IncludingSubsetsToolStripMenuItem_Click(sender As Object, e As EventArgs)
        AddMovies(True)
    End Sub

    Private Sub CurrentOnlyToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        AddMovies(False)
    End Sub

    Private Sub CurrentOnlyToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Addpics(False)
    End Sub




    Private Sub tvMain2_Enter(sender As Object, e As EventArgs) Handles lbxShowList.Enter, lbxFiles.Enter, tvMain2.Enter
        sender.backcolor = Color.Aquamarine
    End Sub

    Private Sub tvMain2_Leave(sender As Object, e As EventArgs) Handles lbxShowList.Leave, lbxFiles.Leave, tvMain2.Leave
        sender.backcolor = Color.White

    End Sub




    Private Sub AddPicturesBGW_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Addpics(True)
    End Sub



    Private Sub tvMain2_NodeSelected(sender As Object, e As TreeViewEventArgs)
        ' Exit Sub
        If e.Node.ToolTipText = "My Computer" Then Exit Sub

        If e.Node.ToolTipText = "" Then Exit Sub
        ' MsgBox(e.Node.ToolTipText)
        Dim di = New IO.DirectoryInfo(e.Node.ToolTipText)
        ChangeFolder(di.FullName, True)
        tmrListbox.Interval = 750
        tmrListbox.Enabled = True
    End Sub



    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub AddAllFilesToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub CurrentOnlyToolStripMenuItem2_Click(sender As Object, e As EventArgs)
        AddFiles(False)
    End Sub

    Private Sub AllSubFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs)
        AddFiles(True)
    End Sub

    Private Sub AddPicturesAndVideosToolStripMenuItem_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub CurrentOnlyToolStripMenuItem3_Click(sender As Object, e As EventArgs)
        AddPicVids(False)
    End Sub

    Private Sub AllSubfoldersToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        AddPicVids(True)
    End Sub

    Private Sub AddPicVids(blnRecurse As Boolean)
        AddFilesToCollection(Showlist, FBCShown, strPicExtensions & strVideoExtensions, blnRecurse)
        FillShowbox(lbxShowList, FilterState.All, Showlist)
    End Sub
    Public Sub UpdateBoxes(strold As String, strnew As String)

        lbxFiles.Items.Remove(strold)
        lbxShowList.Items.Remove(strold)
        lbxShowList.Items.Add(strnew)
    End Sub
    Private Sub tmrLoadLastFolder_Tick(sender As Object, e As EventArgs) Handles tmrLoadLastFolder.Tick
        If strCurrentFilePath = "" Then Exit Sub
        ' strCurrentFilePath="E:\"
        HighlightCurrent(strCurrentFilePath)
        tmrLoadLastFolder.Enabled = False
        'LoadDefaultShowList()
    End Sub


    'Private Sub RandomStartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RandomStartToolStripMenuItem.Click
    '    JumpRandom(False)
    'End Sub


    Delegate Sub UpdateForm_Delegate(ByVal [TS] As StatusStrip, ByVal [text] As String)

    Public Sub UpdateForm_ThreadSafe(ByVal [TSS] As StatusStrip, ByVal [text] As String)
        If [TSS].InvokeRequired Then
            Dim MyDelegate As New UpdateForm_Delegate(AddressOf UpdateForm_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {TSS, [text]})
        Else
            TSS.Text = [text]
        End If
    End Sub


    Private Sub tvMain2_DirectorySelected(sender As Object, e As DirectoryInfoEventArgs) Handles tvMain2.DirectorySelected
        PreferencesSave()
        ChangeFolder(e.Directory.FullName, True)
        ' tvMain2.SelectedFolder = CurrentFolderPath
        FillListbox(lbxFiles, e.Directory, CurrentFilterState, Showlist, blnChooseRandomFile)
        If lbxFiles.Items.Count = 0 Then tbFiles.Text = "0/" & Str(Showlist.Count)
    End Sub

    Private Sub tvMain2_KeyDown(sender As Object, e As KeyEventArgs) Handles tvMain2.KeyDown
        'HandleKeys(sender, e)
        Select Case e.KeyCode
            Case Keys.Down, Keys.Left, Keys.Right, Keys.Up, tvMain2.TraverseKey, tvMain2.TraverseKeyBack, Keys.Tab

            Case Else
                If PFocus = CtrlFocus.Tree Then e.Handled = True
        End Select

    End Sub

    Private Sub tvMain2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tvMain2.KeyPress
        e.Handled = True
    End Sub

    Private Sub btnChooseRandom_Click(sender As Object, e As EventArgs) Handles btnChooseRandom.Click
        blnChooseRandomFile = Not blnChooseRandomFile


    End Sub

    Private Sub frmMain_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        e.Handled = True
    End Sub

    Private Sub ToolStripButton4_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        'Toggle collapse
        CollapseShowlist(Not MasterContainer.Panel2Collapsed)
    End Sub

    Private Sub RefreshTree(sender As Object, e As DriveInfoEventArgs)
        tvMain2.Update()
    End Sub

    Private Sub ToolStripButton6_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        LoadDefaultShowList()
    End Sub

    Private Sub ToolStripProgressBar2_Click(sender As Object, e As EventArgs) Handles TSPB.Click

    End Sub
    Public Sub SetMotion(KeyCode As Integer)

        Dim intSpeed As Integer
        currentWMP.Ctlcontrols.pause()
        tmrMediaSpeed.Enabled = True
        Select Case KeyCode
            Case KeySpeed1
                intSpeed = 1000 / iPlaybackSpeed(0)
            Case KeySpeed2
                intSpeed = 1000 / iPlaybackSpeed(1)
            Case KeySpeed3
                intSpeed = 1000 / iPlaybackSpeed(2)
        End Select

        tmrMediaSpeed.Interval = intSpeed ' 020406

    End Sub
    Private Sub MainWMP_PlayStateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles MainWMP.PlayStateChange
        PlaystateChange(sender, e)
    End Sub

    Private Sub tmrMediaSpeed_Tick(sender As Object, e As EventArgs) Handles tmrMediaSpeed.Tick
        currentWMP.Ctlcontrols.step(1)
    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click

    End Sub

    Private Sub ButtonListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ButtonListToolStripMenuItem.Click
        KeyAssignmentsRestore()
        CollapseShowlist(False)
    End Sub

    Private Sub ButtonListToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ButtonListToolStripMenuItem1.Click
        SaveButtonlist()
    End Sub

    Private Sub tmrUpdateForm_Tick(sender As Object, e As EventArgs) Handles tmrUpdateForm.Tick
        Me.Update()
    End Sub


    Private Sub lbxFiles_DoubleClick(sender As Object, e As EventArgs) Handles lbxFiles.DoubleClick
        Process.Start(lbxFiles.SelectedItem)
    End Sub

    Public Sub UpdateFileInfo(f As FileInfo)
        Dim listcount = lbxFiles.Items.Count
        Dim showcount = lbxShowList.Items.Count
        tbDate.Text = f.LastWriteTime.ToShortDateString & " " & f.LastWriteTime.ToShortTimeString
        tbFiles.Text = listcount & "/" & showcount
        tbFilter.Text = "Filter:" & Filterstates(CurrentFilterState)
        tbLastFile.Text = strCurrentFilePath
        tbRandom.Text = "ORDER:" & UCase(strPlayOrder(ChosenPlayOrder))
        tbShowfile.Text = "SHOWFILE"
        tbSpeed.Text = tbSpeed.Text = "SPEED (" & PlaybackSpeed * 100 & "%)"
        If blnRandomStartPoint Then
            tbStartpoint.Text = "START:RANDOM"
        Else
            tbStartpoint.Text = "START:NORMAL"
        End If

    End Sub

    Private Sub ToolStripButton7_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        SelectSubList()
    End Sub

    Private Sub SelectSubList()
        Dim s As String = InputBox("Enter selection string")
        lbxShowList.Items.Clear()
        'If PFocus=CtrlFocus.Files
        If lbxShowList.Items.Count = 0 Then
            CollapseShowlist(False)
            Showlist = MakeSubList(FileboxContents, s)

        Else
            Oldlist = Showlist
            lbxShowList.Items.Clear()
            Showlist = MakeSubList(Showlist, s)
        End If
        FillShowbox(lbxShowList, CurrentFilterState, Showlist)
    End Sub

    Private Sub ToolStripButton14_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton14.Click
        DeleteShowListFiles()
    End Sub

    Private Sub DeleteShowListFiles()
        If MsgBox("This deletes all files in showlist. Sure?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
            For Each file In lbxShowList.Items
                My.Computer.FileSystem.DeleteFile(file, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)

            Next
            lbxShowList.Items.Clear()
            CollapseShowlist(True)
        End If
    End Sub

    Private Sub tvMain2_GotFocus(sender As Object, e As EventArgs) Handles tvMain2.GotFocus
        PFocus = CtrlFocus.Tree
    End Sub

    Private Sub lbxFiles_GotFocus(sender As Object, e As EventArgs) Handles lbxFiles.GotFocus
        PFocus = CtrlFocus.Files

    End Sub

    Private Sub lbxShowList_GotFocus(sender As Object, e As EventArgs) Handles lbxShowList.GotFocus
        PFocus = CtrlFocus.ShowList
        lCurrentDisplayIndex = lbxShowList.SelectedIndex

    End Sub

    Private Sub LinearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LinearToolStripMenuItem.Click
        'AssignLinear()
    End Sub

    Private Sub HarvestFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HarvestFolderToolStripMenuItem.Click

    End Sub

    Private Sub DeleteEmptyFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteEmptyFoldersToolStripMenuItem.Click
        If Not MsgBox("This deletes all empty directories", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Exit Sub
        Else
            DeleteEmptyFolders(New DirectoryInfo(CurrentFolderPath), True)
        End If

    End Sub

    Private Sub HarvestFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HarvestFoldersToolStripMenuItem.Click
        HarvestCurrent()
    End Sub

    Private Sub HarvestUp()
        Dim di As New DirectoryInfo(CurrentFolderPath)
        HarvestFolder(di, di.Parent, True)
        FillListbox(lbxFiles, di, CurrentFilterState, FileboxContents, False)

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    Private Sub HarvestCurrent()
        Dim di As New DirectoryInfo(CurrentFolderPath)
        HarvestFolder(di, di, True)
        FillListbox(lbxFiles, di, CurrentFilterState, FileboxContents, False)
    End Sub

    Private Sub lbxFiles_KeyDown(sender As Object, e As KeyEventArgs) Handles lbxFiles.KeyDown

    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As CancelEventArgs)

    End Sub

    Private Sub tmrSlowMo_Tick(sender As Object, e As EventArgs) Handles tmrSlowMo.Tick
        MediaAdvance(currentWMP, 1)
    End Sub

    Private Sub tvMain2_Load(sender As Object, e As EventArgs) Handles tvMain2.Load

    End Sub

    Private Sub AddCurrentFileListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddCurrentFileListToolStripMenuItem.Click
        AddCurrentType(False)
    End Sub

    Private Sub AddCurrentAndSubfoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddCurrentAndSubfoldersToolStripMenuItem.Click
        AddCurrentType(True)
    End Sub

    Private Sub DuplicatesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DuplicatesToolStripMenuItem.Click
        FindDuplicates.Show()

    End Sub

    Private Sub lbxShowList_TextChanged(sender As Object, e As EventArgs) Handles lbxShowList.TextChanged
        If Oldlist.Count <> 0 Then
            If MsgBox("Revert list?") Then
                Showlist = Oldlist
                FillShowbox(lbxShowList, CurrentFilterState, Showlist)
            End If
        End If
    End Sub

    Private Sub tmrListbox_Tick(sender As Object, e As EventArgs) Handles tmrListbox.Tick
        FillListbox(lbxFiles, New DirectoryInfo(CurrentFolderPath), CurrentFilterState, Showlist, blnChooseRandomFile)
        tmrListbox.Enabled = False
    End Sub
End Class
