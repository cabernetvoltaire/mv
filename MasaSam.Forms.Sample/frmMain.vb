
Option Explicit On
Imports System.ComponentModel
Imports System.Drawing.Color
Imports System.IO
Imports AxWMPLib
Imports MasaSam.Forms.Controls

Public Class frmMain

    Public Shared ReadOnly Property MyPictures As String
    Public Filterstates As String() = {"All", "Picture only", "Videos only", "Pictures and videos", "Links only", "Not pictures or videos"}
    Public FilterColours As New List(Of Color)
    Public defaultcolour As Color = Color.Aqua
    Public movecolour As Color = Color.Orange
    Public watchfolder As FileSystemWatcher
    Public StartType As Byte
    Public StartPointPercentage As Integer


    Private Sub InitialiseColours()
        FilterColours.Add(Color.MintCream)
        FilterColours.Add(Color.LemonChiffon)
        FilterColours.Add(Color.LightPink)
        FilterColours.Add(Color.LightSeaGreen)
        FilterColours.Add(Color.LightBlue)
        FilterColours.Add(Color.PaleTurquoise)

    End Sub
    Public Sub SetControlColours(blnMove As Boolean)
        Dim d As Color = defaultcolour
        If blnMove Then d = movecolour
        Dim s As Color = FilterColours.Item(CurrentFilterState)
        tvMain2.BackColor = s
        tvMain2.HighlightSelectedNodes()
        lbxFiles.BackColor = s
        lbxShowList.BackColor = s
        If PFocus = CtrlFocus.Files Then lbxFiles.BackColor = d
        If PFocus = CtrlFocus.Tree Then tvMain2.BackColor = d
        If PFocus = CtrlFocus.ShowList Then lbxShowList.BackColor = d

    End Sub
    Private Sub MovietoPic(img As Image)
        currentWMP.Visible = False
        currentWMP.URL = ""
        PreparePic(currentPicBox, pbxBlanker, img)
        currentPicBox.Visible = True
        currentPicBox.BringToFront()
        currentWMP.Visible = False
    End Sub
    Private Sub GetMetadata(spath As String)
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

    Private Sub HandleMovie(blnRandom As Boolean)
        'If it is to jump to a random point, do not show first.

        If tmrJumpVideo.Enabled Then
            tbStartpoint.Text = "START:RANDOM"
            currentWMP.Visible = False
            '   While tmrJumpVideo.Enabled

            'End While
        Else
            tbStartpoint.Text = "START:NORMAL"

            currentWMP.Visible = True
        End If

        currentWMP.URL = strCurrentFilePath
        'Dim f As Int16
        'f = currentWMP.currentMedia.getItemInfo(FrameRate)
        '  MsgBox(f)
        currentWMP.BringToFront()
        ' AxVLCPlugin21.MRL = "file:///" & strCurrentFilePath & ""
        'AxVLCPlugin21.BringToFront()
        'AxVLCPlugin21.playlist.togglePause()
        '  If PlaybackSpeed <> 1 Then currentWMP.settings.rate = PlaybackSpeed
        currentPicBox.Visible = False

        If tmrSlideShow.Enabled Then
            blnRestartSlideShowFlag = True
            tmrSlideShow.Enabled = False 'Slideshow stops if movie. Create separate timer for movie slideshows. 
        End If
    End Sub
    Private Sub SaveShowlist()
        Dim path As String
        With SaveFileDialog1
            .DefaultExt = "msl"
            .Filter = "Metavisua list files|*.msl|All files|*.*"
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                path = .FileName
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
        With OpenFileDialog1

            .DefaultExt = "msl"
            .Filter = "Metavisua list files|*.msl|All files|*.*"

            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                path = .FileName
            End If
        End With
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
        'tbShowfile.Text = "SHOWFILE LOADED:" & path
    End Sub
    Private Sub AxWindowsMediaPlayer1_MediaError(ByVal sender As Object,
    ByVal e As _WMPOCXEvents_MediaErrorEvent) Handles MainWMP.MediaError
        'MsgBox(e.pMediaObject.ToString)
    End Sub
    Public Sub CancelDisplay()
        If currentWMP.Visible Then
            currentWMP.Ctlcontrols.pause()
            currentWMP.URL = ""
        End If
        If currentPicBox.Visible Then
            currentPicBox.Image = Nothing
            GC.Collect()

        End If
        ' currentPicBox.Dispose()
        '        tmrAutoTrail.Enabled = False
        tmrSlideShow.Enabled = False
    End Sub

    Private Sub DeleteFolder(tvw As FileSystemTree, blnConfirm As Boolean)
        With My.Computer.FileSystem
            Dim m As MsgBoxResult = MsgBoxResult.No
            If .DirectoryExists(CurrentFolderPath) Then
                If blnConfirm Then
                    m = MsgBox("Delete folder " & CurrentFolderPath & "?", MsgBoxStyle.YesNoCancel)
                End If
                If Not blnConfirm OrElse m = MsgBoxResult.Yes Then
                    Dim f As New DirectoryInfo(CurrentFolderPath)
                    Try
                        .DeleteDirectory(CurrentFolderPath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    Catch x As System.OperationCanceledException

                        Exit Sub
                    End Try

                    tvw.Traverse(False)
                    tvw.RemoveNode(f.FullName)
                    ' tvw.RefreshTree(f.Parent.FullName)
                End If

            ElseIf m = MsgBoxResult.No Or m = MsgBoxResult.Cancel Then
                ControlSetFocus(lbxFiles)

            End If
        End With
    End Sub


    Public Sub UpdatePlayOrder(blnShowBoxShown As Boolean)
        If blnShowBoxShown Then

            tbRandom.Text = "ORDER:" & UCase(strPlayOrder(ChosenPlayOrder))
            Showlist = SetPlayOrder(ChosenPlayOrder, Showlist)
            lbxShowList.Items.Clear()
            FillShowbox(lbxShowList, CurrentFilterState, Showlist)
        End If
        '        lbxFiles.Items.Clear()
        'FileboxContents = SetPlayOrder(ChosenPlayOrder, FileboxContents)
        Dim e = New DirectoryInfo(CurrentFolderPath)
        FillListbox(lbxFiles, e, FileboxContents, False)
        '       FillShowbox(lbxFiles, CurrentFilterState, FileboxContents)
        ReportFilterandOrder()
    End Sub

    Public Sub ReportFilterandOrder()
        cbxOrder.SelectedIndex = ChosenPlayOrder
        cbxFilter.SelectedIndex = CurrentFilterState
    End Sub


    'Private Function SpeedChange(e As KeyEventArgs, blnTrue As Boolean)
    '   SetMotion(e.KeyCode) 'Alternative speed. Doesn't work at the moment. 
    'End Function
    Private Function SpeedChange(e As KeyEventArgs) As KeyEventArgs
        If e.KeyData = KeySpeed1 + Keys.Control Or e.KeyData = KeySpeed3 + Keys.Control Then

            TweakSpeed(e)
            e.Handled = True
            Exit Function
        End If
        Dim blnPlaying As Boolean = currentWMP.URL <> ""
        Dim Choice As Byte = e.KeyCode - KeySpeed1 'Set slideshow speed if pic showing, and start slideshow
        If Not blnPlaying Then
            'PlaybackSpeed = 30
            tmrSlideShow.Enabled = True
            tmrSlideShow.Interval = iSSpeeds(Choice)
        Else

            PlaybackSpeed = iPlaybackSpeed(Choice) 'Otherwise, set playback speed 'TODO Options
        End If


        If e.KeyCode = KeyToggleSpeed Then
            If blnPlaying Then

                currentWMP.settings.rate = 1
                currentWMP.Ctlcontrols.play()
                tmrSlowMo.Enabled = False
            Else
                tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
            End If

        Else
            ' currentWMP.settings.rate = PlaybackSpeed / 30

            currentWMP.Ctlcontrols.pause()
            'If PlaybackSpeed = 30 Then
            '    currentWMP.settings.rate = 1
            '    currentWMP.Ctlcontrols.play()
            '    tmrSlowMo.Enabled = False
            'Else

            tmrSlowMo.Interval = 1000 / (PlaybackSpeed)
            tmrSlowMo.Enabled = True
        End If

        'End If
        'Report
        blnSpeedRestart = True
        ' tbSpeed.Text = "SPEED (" & PlaybackSpeed * 100 & "%)"
        e.Handled = True
        Return e
    End Function

    Private Shared Sub TweakSpeed(e As KeyEventArgs)
        If e.KeyCode = KeySpeed1 Then 'TODO does this work?
            If e.Control Then 'increase the extremes if Control held TODO Don't know if this works. 
                If e.Shift Then 'decrease if Shift
                    iSSpeeds(0) = iSSpeeds(0) * 0.9
                Else
                    iSSpeeds(0) = iSSpeeds(0) / 0.9
                End If
                e.Handled = True

            End If

        End If
        If e.KeyCode = KeySpeed3 Then
            If e.Control Then 'increase the extremes if Control held
                If e.Shift Then 'decrease if Shift
                    iSSpeeds(2) = iSSpeeds(2) * 0.9
                Else
                    iSSpeeds(2) = iSSpeeds(2) / 0.9
                End If
                e.Handled = True
            End If

        End If

    End Sub

    Private Sub ToggleRandomStartPoint()
        blnRandomStartAlways = Not blnRandomStartAlways
        blnRandomStartPoint = blnRandomStartAlways
        'StartAlways is when the random start has been selected for all files
        'StartPoint is just a flag telling the video to jump to 
    End Sub
    Public Sub GoFullScreen(blnGo As Boolean)
        FullScreen.Changing = True
        If blnGo Then

            SetWMP(FullScreen.FSWMP)
            SetPB(FullScreen.fullScreenPicBox)
            Dim screen As Screen
            If blnSecondScreen Then
                screen = Screen.AllScreens(1)
            Else
                screen = Screen.AllScreens(0)
                SplitterPlace(0.75)
            End If
            FullScreen.StartPosition = FormStartPosition.Manual
            FullScreen.Location = screen.Bounds.Location + New Point(100, 100)
            FullScreen.Show()

        Else
            SplitterPlace(0.25)
            SetWMP(MainWMP)
            SetPB(PictureBox1)
            FullScreen.Close()
        End If
        blnFullScreen = Not blnFullScreen
        tmrPicLoad.Enabled = True
    End Sub
    Private Sub SplitterPlace(i As Decimal)
        ctrMainFrame.SplitterDistance = Me.Width * i
    End Sub
    Private Sub ToggleButtons()
        Buttons_Load()
        blnButtonsLoaded = Not blnButtonsLoaded
        ctrPicAndButtons.Panel2Collapsed = Not blnButtonsLoaded
        UpdateButtonAppearance()
    End Sub


    Public Sub AdvanceFile(blnForward As Boolean, blnTest As Boolean)
        'Advance using whichever control has focus 
        'Unless control pressed, in which case, always advance lbxfiles. 
        Dim diff As Integer

        If blnForward Then
            diff = 1
        Else
            diff = -1
        End If

        Dim lbx As New ListBox
        If PFocus = CtrlFocus.ShowList And Not CtrlDown Then
            lbx = lbxShowList
        Else
            lbx = lbxFiles
        End If
        Dim count As Long
        count = lbx.Items.Count
        ReDim Preserve FBCShown(count)
        lbx.SelectionMode = SelectionMode.One
        If count = 0 Then Exit Sub 'if no filelist, then give up.

        If lbx.SelectedIndex = 0 And Not blnForward Then
            lbx.SelectedIndex = count - 1
        Else
            If count > 0 Then

                If (blnRandomAdvance(CtrlFocus.ShowList) And lbx Is lbxShowList) Or (blnRandomAdvance(CtrlFocus.Files And lbx Is lbxFiles)) Or blnRandomAdvance(PFocus) Then
                    Dim i As Int32
                    i = Int(Rnd() * (count))
                    While FBCShown(i)
                        i = Int(Rnd() * (count))
                    End While
                    lbx.SelectedIndex = i
                Else
                    lbx.SelectedIndex = (lbx.SelectedIndex + diff) Mod count
                End If
            End If


        End If
        FBCShown(lbx.SelectedIndex) = True
        NofShown += 1
        If NofShown >= count Then
            ReDim FBCShown(count)
            NofShown = 0
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
        ' blnRandomStartPoint = False
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
            If blnRandomStartPoint Then
                NewPosition = (Rnd(1) * (currentWMP.currentMedia.duration))
            End If
            tmrJumpVideo.Interval = lngInterval
            tmrJumpVideo.Enabled = True
            tbStartpoint.Text = "START:RANDOM"


        Else
            'MsgBox("Autotrail")
            ToggleAutoTrail()
        End If

    End Sub
    Public Sub ToggleAutoTrail()
        tmrAutoTrail.Enabled = Not tmrAutoTrail.Enabled
        TrailerModeToolStripMenuItem.Checked = tmrAutoTrail.Enabled
    End Sub
    Public Sub HandleKeys(sender As Object, e As KeyEventArgs)
        Me.Cursor = Cursors.WaitCursor
        'MsgBox(e.KeyCode.ToString)
        Select Case e.KeyCode
            Case Keys.Enter + Keys.Control
                If blnLink Then
                    blnLink = False
                    HighlightCurrent(strCurrentFilePath)
                End If
            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                HandleFunctionKeyDown(sender, e)
                e.Handled = True
            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9
                If Not e.Control Then
                    ChangeButtonLetter(e)
                Else
                    Select Case e.KeyCode
                        Case Keys.I
                            CreateNewDirectory(CurrentFolderPath)

                    End Select
                End If
            Case Keys.Left, Keys.Right
                tvMain2_KeyDown(sender, e)
            Case KeyToggleButtons
                ToggleButtons()
            Case KeyEscape
                CancelDisplay()                'currentPicBox.Image.Dispose()
                tmrAutoTrail.Enabled = False
            Case KeyRandomize
                'Cycle through play orders
                ChosenPlayOrder = (ChosenPlayOrder + 1) Mod (PlayOrder.Type + 1) 'Change
                UpdatePlayOrder(Showlist.Count > 0)
            Case KeyRandomize + Keys.Control
                ChosenPlayOrder = (ChosenPlayOrder + 1) Mod (PlayOrder.Type + 1)
                UpdatePlayOrder(Showlist.Count > 0)
            Case KeyAddFile
                AddCurrentFileToShowlistToolStripMenuItem_Click(Nothing, Nothing)

            Case KeyNextFile, KeyPreviousFile, LKeyNextFile, LKeyPreviousFile

                AdvanceFile(e.KeyCode = KeyNextFile, e.Control)
                e.Handled = True
                tmrSlideShow.Enabled = False

            Case KeySmallJumpDown, KeySmallJumpUp, LKeySmallJumpDown, LKeySmallJumpUp
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
            Case KeyMarkPoint, LKeyMarkPoint
                'Addmarker(strCurrentFilePath)
                If MediaMarker = 0 Then
                    MediaMarker = currentWMP.Ctlcontrols.currentPosition
                Else
                    MediaMarker = 0
                End If
                e.Handled = True
            Case KeyLoopToggle
                blnLoopPlay = Not blnLoopPlay
            Case KeyRotateBack
                RotatePic(currentPicBox, False)
            Case KeyRotate
                RotatePic(currentPicBox, True)
            Case KeyJumpAutoT
                If e.Control Then
                    ToggleRandomStartPoint()
                End If
                JumpRandom(e.Shift)

            Case KeyTraverseTree, KeyTraverseTreeBack
                'Handled by Treeview behaviour unless focus is elsewhere. 
                'We want the traverse keys always to work. 

                If PFocus <> CtrlFocus.Tree Then
                    ControlSetFocus(tvMain2)
                    tvMain2_KeyDown(sender, e)
                    'e.Handled = True
                End If
                'TraverseTree(tvMain2.tvFiles, e.KeyCode = KeyTraverseTree)
            Case KeyToggleSpeed
                With currentWMP
                    If .URL <> "" Then
                        If .playState = WMPLib.WMPPlayState.wmppsPaused Then
                            tmrSlowMo.Enabled = False
                            .settings.rate = 1
                            .Ctlcontrols.play()
                        Else
                            .Ctlcontrols.pause()
                        End If
                    Else
                        tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
                    End If
                    blnSpeedRestart = True
                    '  tbSpeed.Text = "SPEED (" & PlaybackSpeed * 100 & "%)"
                End With
            Case KeySpeed1, KeySpeed2, KeySpeed3, KeySpeed3 + Keys.Control
                SpeedChange(e)
            Case KeyFilter 'Cycle through listbox filters
                CurrentFilterState = (CurrentFilterState + 1) Mod (FilterState.NoPicVid + 1) '(GetMaxValue(Of FilterState)() + 1)
                SetFilterState()
                e.Handled = True
            Case KeySelect
                SelectSubList(False)

            Case KeyFullscreen
                If ShiftDown Then
                    blnSecondScreen = True
                Else
                    blnSecondScreen = False
                End If
                GoFullScreen(Not blnFullScreen)
                e.Handled = True

            Case KeyDelete
                'Use Movefiles with current selected list, and option to delete. 
                CancelDisplay()
                Select Case PFocus
                    Case CtrlFocus.Files
                        MoveFiles(ListfromListbox(lbxFiles), "", lbxFiles)
                    Case CtrlFocus.Tree
                        DeleteFolder(tvMain2, Not blnMoveMode)
                    Case CtrlFocus.ShowList
                        '                        If MsgBox("This will DELETE ALL the showlist files! Is this what you want?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes Then Exit Sub
                        MoveFiles(ListfromListbox(lbxShowList), "", lbxShowList)
                End Select
                'DisposePic(currentPicBox)
                'Deletefile(strCurrentFilePath)
                'UpdateBoxes(strCurrentFilePath, "")

            Case KeyMoveToggle
                ToggleMove()

            Case KeyLoopToggle
                currentWMP.settings.setMode("loop", Not currentWMP.settings.getMode("loop"))
                e.Handled = True
            Case KeyTrueSize
                ToggleTrueSize(currentPicBox)

            Case KeyBackUndo
                If LastFolder.Count > 0 Then
                    CurrentFolderPath = LastFolder.Pop
                    tvMain2.SelectedFolder = CurrentFolderPath
                End If
                'strCurrentFilePath = LastPlayed.Pop
                'tmrPicLoad.Enabled = True

            Case KeyReStartSS
                tmrSlideShow.Enabled = Not tmrSlideShow.Enabled

        End Select
        Me.Cursor = Cursors.Default
        ' e.Handled = True
    End Sub
    Private Sub LoopToggle()
        blnLoopPlay = Not blnLoopPlay
        'currentWMP.= blnLoopPlay
    End Sub
    Private Sub SetFilterState()
        cbxFilter.Select(CurrentFilterState, 1)
        'MsgBox(CurrentFilterState.ToString)
        'FillListbox(lbxFiles, New DirectoryInfo(CurrentFolderPath), FileboxContents, blnChooseRandomFile)
        UpdatePlayOrder(Showlist.Count > 0)
        tbFilter.Text = UCase(Filterstates(CurrentFilterState))
        SetControlColours(blnMoveMode)
    End Sub

    Private Sub SetWMP(tWMP As AxWindowsMediaPlayer)
        With currentWMP
            tWMP.URL = .URL
            ' AxVLCPlugin21.MRL = "file:///" & .URL
            tWMP.Ctlcontrols.currentPosition = .Ctlcontrols.currentPosition
            .URL = ""
            currentWMP = tWMP
            .stretchToFit = True
            .Visible = True
            .BringToFront()
            'AxVLCPlugin21.BringToFront()
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
        'If strPath is a link, it highlights the link, not the file
        If strPath = "" Then Exit Sub 'Empty
        If Len(strPath) > 247 Then Exit Sub 'Too long
        Dim finfo As New FileInfo(strPath)
        UpdateFileInfo()
        If InStr(":\", strPath) = Len(strPath) - 2 Then 'TODO root folder
            Dim s As String = Path.GetDirectoryName(strPath)
            If tvMain2.SelectedFolder <> s Then tvMain2.SelectedFolder = s 'Only change tree if it needs changing

        Else
            'If Not blnLink Then

            Dim s As String = Path.GetDirectoryName(strPath)
            If tvMain2.SelectedFolder <> s Then tvMain2.SelectedFolder = s 'Only change tree if it needs changing
            'End If
        End If


        'If Not blnLink Then
        If lbxFiles.SelectedItem <> strPath Then
            'If Not blnChooseOne Then 'TODO Rewrite this so that we can get a random file in the box when blnChooseONe is selected
            'HighlightListboxSelected(strPath, lbxFiles) 'Highlight the current file in lbxFiles
            lbxFiles.SelectedIndex = lbxFiles.FindString(strPath)

            'Else
            '    strPath = lbxFiles.SelectedItem
            'End If
        End If
        'End If
        If Not MasterContainer.Panel2Collapsed Then
            'Highlight the file in lbxShowlist, but only if it has the focus
            If PFocus = CtrlFocus.ShowList AndAlso Not CtrlDown Then
                'HighlightListboxSelected(strPath, lbxShowList)
                lbxShowList.SelectedIndex = lbxShowList.FindString(strPath)
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
        AddFilesToCollection(Showlist, strVideoExtensions, blnRecurse)
        FillShowbox(lbxShowList, FilterState.All, Showlist)
    End Sub
    Private Sub AddFiles(blnRecurse As Boolean)
        ProgressBarOn(1000)

        AddFilesToCollection(Showlist, "", blnRecurse)
        FillShowbox(lbxShowList, CurrentFilterState, Showlist)
        ProgressBarOff()
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


    Private Sub RemoveFilesFromCollection(ByVal list As List(Of String), extensions As String)
        Dim d As New System.IO.DirectoryInfo(CurrentFolderPath)
        FindAllFilesBelow(d, list, extensions, True, "", True, False)
    End Sub

    Private Sub GlobalInitialise()
        '  InitialiseColours()
        CollapseShowlist(True)
        PreferencesGet()
        Randomize()
        AssignExtensionFilters()
        tmrLoadLastFolder.Enabled = True

        tmrPicLoad.Interval = lngInterval
        tmrJumpVideo.Interval = lngInterval
        currentWMP = MainWMP
        currentWMP.stretchToFit = True
        currentWMP.uiMode = "FULL"
        currentWMP.Dock = DockStyle.Fill
        currentPicBox = PictureBox1
        cbxFilter.Items.Clear()

        For i = 0 To Filterstates.Length - 1
            cbxFilter.Items.Add(Filterstates(i))
        Next
        cbxOrder.Items.Clear()

        For i = 0 To strPlayOrder.Length - 1
            cbxOrder.Items.Add(strPlayOrder(i))
        Next
        'Exit Sub
        Try
            KeyAssignmentsRestore(strButtonFile)
            If Not blnButtonsLoaded Then
                ToggleButtons()
            End If

        Catch ex As FileNotFoundException
            ctrPicAndButtons.Panel2Collapsed = True
            Exit Try
        Catch ex As DirectoryNotFoundException
            ctrPicAndButtons.Panel2Collapsed = True

            Exit Try
        End Try
        WatchStart(CurrentFolderPath)

    End Sub
    Public Sub WatchStart(path As String)
        Exit Sub
        ' watchfolder = New System.IO.FileSystemWatcher()
        FileSystemWatcher1.Path = path
        FileSystemWatcher1.NotifyFilter = IO.NotifyFilters.LastWrite

        AddHandler FileSystemWatcher1.Changed, AddressOf logchange
        AddHandler FileSystemWatcher1.Created, AddressOf logchange
        AddHandler FileSystemWatcher1.Deleted, AddressOf logchange
        ' AddHandler FileSystemWatcher1.Renamed, AddressOf logrename
        FileSystemWatcher1.EnableRaisingEvents = True
    End Sub
    Private Sub logchange(ByVal source As Object, ByVal e As System.IO.FileSystemEventArgs) Handles FileSystemWatcher1.Changed
        If e.ChangeType = System.IO.WatcherChangeTypes.Changed Then
            MsgBox("File " & e.FullPath & " has been modified")
        End If

    End Sub
    'Form Controls
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        GlobalInitialise()

    End Sub



    Private Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        'If e.KeyCode = Keys.Tab Then MsgBox("Yes")
        ShiftDown = e.Shift
        CtrlDown = e.Control
        UpdateButtonAppearance()
        'GiveKey(e.KeyCode)
        HandleKeys(sender, e)
        If e.KeyCode = KeyBackUndo Then
            e.Handled = True
        End If
        'e.Handled = True
        '  MsgBox(sender.ToString & " " & e.ToString)
    End Sub
    Private Sub frmMain_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        ShiftDown = e.Shift
        CtrlDown = e.Control
        UpdateButtonAppearance()

        ' MsgBox("Ring")
    End Sub
    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Settings.PreferencesSave()
    End Sub


    'Private Sub lbxFiles_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    'Loads a new media file by triggering the PicLoad Timer

    '    strCurrentFilePath = lbxFiles.SelectedItem

    '    tmrPicLoad.Interval = lngInterval
    '    tmrPicLoad.Enabled = True
    'End Sub



    'Timers

    Private Sub HandleDoc()
        'PrintPreviewControl1.Document = strCurrentFilePath
        'Throw New NotImplementedException()
    End Sub
    'Tool strip buttons
    Private Sub btnToggleMark_Click(sender As Object, e As EventArgs)


    End Sub

    'Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs)
    '    GoFullScreen(True)

    'End Sub

    'Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs)
    '    ClearShowList()
    'End Sub

    Private Sub ClearShowList()
        Showlist.Clear()

        lbxShowList.Items.Clear()
        CollapseShowlist(True)

    End Sub

    'Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs)
    '    AddMovies(True)
    'End Sub
    'Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs)
    '    Addpics(True)
    'End Sub
    'Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs)
    '    ToggleRandomStartPoint()
    'End Sub


    'Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs)

    '    FindDuplicates.Show()
    '    'Duplicates2.Flist = Showlist
    '    'Duplicates2.Show()
    'End Sub
    'Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs)
    '    SaveShowlist()
    '    'currentPicBox.Tag = currentPicBox.ImageLocation
    '    'FullScreen.Show()
    'End Sub
    'Private Sub ToolStripButton12_Click(sender As Object, e As EventArgs)
    '    LoadShowList()
    '    'newThread.Start()
    'End Sub
    'Private Sub ToolStripButton13_Click(sender As Object, e As EventArgs)

    'End Sub
    ''Private Sub ToolStripButton13_SelectedIndexChanged(sender As Object, e As EventArgs)
    ''    Dim i As Integer
    ''    i = ToolStripButton13.SelectedIndex
    ''    Showlist = SetPlayOrder(i, Showlist)
    ''    FillShowbox(lbxShowList, CurrentFilterState, Showlist)
    ''End Sub
    'Private Sub ToolStripButton14_Click(sender As Object, e As EventArgs)
    '    RemoveFilesFromCollection(Showlist, strVideoExtensions)
    '    FillShowbox(lbxShowList, FilterState.All, Showlist)

    'End Sub
    'Private Sub ToolStripButton15_Click(sender As Object, e As EventArgs)
    '    tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
    '    If tmrSlideShow.Enabled = False Then blnRestartSlideShowFlag = False
    'End Sub
    'Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    Dim i As Integer
    '    i = cbxSpeeds.SelectedIndex
    '    ssspeed = CInt(cbxSpeeds.Items(i))
    '    tmrSlideShow.Interval = ssspeed

    'End Sub

    Private Sub Listbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxShowList.SelectedIndexChanged, lbxFiles.SelectedIndexChanged
        With sender
            Dim i As Long = .SelectedIndex
            If .Items.count = 0 Then
                .items.add("If there is nothing showing here, check your filters")
            ElseIf i >= 0 Then
                tmrPicLoad.Enabled = False
                strCurrentFilePath = .Items(i)
                tmrPicLoad.Enabled = True
            End If
        End With
    End Sub





    Private Sub ShowListToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' LoadFiles.RunWorkerAsync()
        LoadShowList()

        'newThread.Start()
    End Sub

    Private Sub tmrSlideShow_Tick(sender As Object, e As EventArgs) Handles tmrSlideShow.Tick
        AdvanceFile(True, False)
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
            currentPicBox.Refresh()
            Dim finfo As New FileInfo(strCurrentFilePath)
            Dim dt As New Date
            'avoid the updating of the write time
            dt = finfo.LastWriteTime
            .Save(strCurrentFilePath)
            finfo.LastWriteTime = dt

        End With
    End Sub
    Private Sub ToolStripButton16_Click(sender As Object, e As EventArgs)
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
    Private Sub SaveListToolStripMenuItem_Click(sender As Object, e As EventArgs)
        SaveShowlist()
    End Sub

    Private Sub IncludingSubfoldersToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Addpics(True)

    End Sub

    Private Sub IncludingSubsetsToolStripMenuItem_Click(sender As Object, e As EventArgs)
        AddFiles(True)
    End Sub

    Private Sub CurrentOnlyToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        AddFiles(False)
    End Sub

    Private Sub CurrentOnlyToolStripMenuItem_Click(sender As Object, e As EventArgs)
        AddFiles(False)
    End Sub




    Private Sub tvMain2_Enter(sender As Object, e As EventArgs) Handles lbxShowList.Enter, lbxFiles.Enter, tvMain2.Enter
        'PFocus = CtrlFocus.Tree
        If sender.Equals(lbxShowList) Then
            PFocus = CtrlFocus.ShowList
        End If
        If sender.Equals(lbxFiles) Then
            PFocus = CtrlFocus.Files
        End If
        If sender.Equals(tvMain2) Then
            PFocus = CtrlFocus.Tree
        End If

        SetControlColours(blnMoveMode)
    End Sub









    'Private Sub tvMain2_NodeSelected(sender As Object, e As TreeViewEventArgs)
    '    ' Exit Sub
    '    If e.Node.ToolTipText = "My Computer" Then Exit Sub

    '    If e.Node.ToolTipText = "" Then Exit Sub
    '    ' MsgBox(e.Node.ToolTipText)
    '    Dim di = New IO.DirectoryInfo(e.Node.ToolTipText)
    '    ChangeFolder(di.FullName, True)
    '    tmrListbox.Interval = 750
    '    tmrListbox.Enabled = True
    'End Sub



    Protected Overrides Sub Finalize()
        MyBase.Finalize()
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
        AddFiles(False)
    End Sub

    Private Sub AllSubfoldersToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        AddFiles(True)
    End Sub

    Private Sub AddPicVids(blnRecurse As Boolean)
        AddFilesToCollection(Showlist, strPicExtensions & strVideoExtensions, blnRecurse)
        FillShowbox(lbxShowList, FilterState.All, Showlist)
    End Sub
    Public Sub UpdateBoxes(strold As String, strnew As String)

        lbxFiles.Items.Remove(strold)
        lbxShowList.Items.Remove(strold)
        If strnew <> "" Then lbxShowList.Items.Add(strnew)
    End Sub
    Private Sub tmrLoadLastFolder_Tick(sender As Object, e As EventArgs) Handles tmrLoadLastFolder.Tick
        tmrLoadLastFolder.Enabled = False
        'MsgBox("LLF")
        If strCurrentFilePath = "" Then Exit Sub
        ' strCurrentFilePath="E:\"
        HighlightCurrent(strCurrentFilePath)
        'LoadDefaultShowList()
    End Sub





    Private Sub tvMain2_DirectorySelected(sender As Object, e As DirectoryInfoEventArgs) Handles tvMain2.DirectorySelected

        ChangeFolder(e.Directory.FullName, True)

        tmrUpdateFolderSelection.Enabled = False
        tmrUpdateFolderSelection.Interval = lngInterval * 8
        tmrUpdateFolderSelection.Enabled = True

    End Sub


    Private Sub tvMain2_KeyDown(sender As Object, e As KeyEventArgs) Handles tvMain2.KeyDown, lbxFiles.KeyDown, lbxShowList.KeyDown
        Select Case e.KeyCode
            Case Keys.Down, Keys.Left, Keys.Right, Keys.Up

            Case tvMain2.TraverseKey, tvMain2.TraverseKeyBack
                tvMain2.Traverse(e.KeyCode = tvMain2.TraverseKeyBack)

                'MsgBox("Pressed")
                e.Handled = True
            Case Else
                If PFocus = CtrlFocus.Tree Then e.Handled = True
        End Select

    End Sub

    Private Sub tvMain2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tvMain2.KeyPress
        ' e.Handled = True
    End Sub

    Private Sub btnChooseRandom_Click(sender As Object, e As EventArgs)
        blnChooseRandomFile = Not blnChooseRandomFile


    End Sub

    Private Sub frmMain_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
    End Sub

    Private Sub ToolStripButton4_Click_1(sender As Object, e As EventArgs)
        'Toggle collapse
        CollapseShowlist(Not MasterContainer.Panel2Collapsed)
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

    Private Sub tmrInitialise_Tick(sender As Object, e As EventArgs) Handles tmrInitialise.Tick
        MsgBox("Initialise")
        ' ToolStripButton3_Click(Me, e)
        tmrInitialise.Enabled = False
    End Sub
    Private Sub tmrPicLoad_Tick(sender As Object, e As EventArgs) Handles tmrPicLoad.Tick


        HighlightCurrent(strCurrentFilePath)
        LastPlayed.Push(strCurrentFilePath)
        tbLastFile.Text = strCurrentFilePath
        fType = FindType(strCurrentFilePath)
        Select Case fType
            Case Filetype.Doc
                HandleDoc()
            Case Filetype.Movie
                HandleMovie(blnRandomStartPoint)

            Case Filetype.Pic
                Dim img As Image
                If Not currentPicBox.Image Is Nothing Then
                    DisposePic(currentPicBox)
                End If
                If Not My.Computer.FileSystem.FileExists(strCurrentFilePath) Then Exit Select

                img = GetImage(strCurrentFilePath)
                Dim e1 As New PaintEventArgs(currentPicBox.CreateGraphics, New Rectangle(New Point(0, 0), currentPicBox.Size))
                If img Is Nothing Then
                    tmrPicLoad.Enabled = False
                    Exit Sub
                End If
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


    Private Sub tmrJumpVideo_Tick(sender As Object, e As EventArgs) Handles tmrJumpVideo.Tick
        tmrJumpVideo.Enabled = False
        'ControlSetFocus(currentWMP)
        'MediaDuration = currentWMP.currentMedia.duration

        currentWMP.Ctlcontrols.currentPosition = NewPosition
        'While currentWMP.playState = WMPLib.WMPPlayState.wmppsTransitioning

        'End While
        currentWMP.Visible = True

        currentWMP.BringToFront()

        If Not blnRandomStartAlways Then blnRandomStartPoint = False
    End Sub


    Private Sub btn1_MouseDown(sender As Object, e As MouseEventArgs) Handles btn1.MouseDown, btn2.MouseDown, btn3.MouseDown, btn4.MouseDown, btn5.MouseDown, btn6.MouseDown, btn7.MouseDown, btn8.MouseDown
        Dim button As Button = sender
        Dim i As Integer = Val(button.Name(3)) - 1

        Dim m As New FolderSelect

        m.Left = button.Left - m.Width / 2
        m.Top = button.Top - m.Height + 50
        m.ButtonNumber = i
        m.Alpha = iCurrentAlpha
        m.Show()
        If strVisibleButtons(i) = "" Then
            m.Folder = CurrentFolderPath
        Else
            m.Folder = strVisibleButtons(i)
        End If
    End Sub

    Private Sub ButtonListToolStripMenuItem_Click(sender As Object, e As EventArgs)
        KeyAssignmentsRestore()
        CollapseShowlist(False)
    End Sub

    Private Sub ButtonListToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        SaveButtonlist()
    End Sub


    Private Sub tmrUpdateForm_Tick(sender As Object, e As EventArgs) Handles tmrUpdateForm.Tick
        MsgBox("UF")
        Me.Update()
    End Sub


    Private Sub lbxFiles_DoubleClick(sender As Object, e As EventArgs) Handles lbxFiles.DoubleClick
        Process.Start(lbxFiles.SelectedItem)
    End Sub

    Public Sub UpdateFileInfo()
        Dim f As New FileInfo(strCurrentFilePath)
        If Not f.Exists Then Exit Sub
        Dim listcount = lbxFiles.Items.Count
        Dim showcount = lbxShowList.Items.Count
        Dim dt As Date
        dt = f.LastAccessTime
        If f.LastWriteTime < dt Then dt = f.LastWriteTime
        If f.CreationTime < dt Then dt = f.CreationTime

        tbDate.Text = dt.ToShortDateString & " " & dt.ToShortTimeString + " (" + Format(f.Length, "#,0.") + " bytes)"
        Dim c As Int16 = lbxFiles.SelectedItems.Count
        If c > 1 Then
            tbFiles.Text = "FOLDER:" & listcount & "(" & c & " selected)" & " SHOW:" & showcount
        Else
            tbFiles.Text = "FOLDER:" & listcount & " SHOW:" & showcount
        End If
        If FilePumpList.Count <> 0 Then tbFiles.Text = tbFiles.Text & " (" & FilePumpList.Count & " files waiting to be moved.)"
        tbFilter.Text = "FILTER:" & Filterstates(CurrentFilterState)
        tbLastFile.Text = strCurrentFilePath
        tbRandom.Text = "ORDER:" & UCase(strPlayOrder(ChosenPlayOrder))
        tbShowfile.Text = "SHOWFILE: " & LastShowList
        tbSpeed.Text = tbSpeed.Text = "SPEED (" & PlaybackSpeed & "fps)"
        tbButton.Text = "BUTTONFILE: " & strButtonFile
        tbZoom.Text = iZoomFactor
        If blnRandomStartPoint Then
            tbStartpoint.Text = "START:RANDOM"
        Else
            tbStartpoint.Text = "START:NORMAL"
        End If

    End Sub

    Private Sub ToolStripButton7_Click_1(sender As Object, e As EventArgs)
        SelectSubList(False)
    End Sub
    Private Function Subdividelist(list As List(Of String), iGroups As Int16)
        'On the current sort parameter
        'Select each chunk
        'Move each chunk to a subfolder of the destination
        'On date
        Dim a As FileInfo
        Dim z As FileInfo
        a = New FileInfo(list(0))
        z = New FileInfo(list(list.Count - 1))
        Dim ldate As Date = a.CreationTime
        Dim udate As Date = z.CreationTime
        Dim span As TimeSpan = udate - ldate
        Dim chunk As New TimeSpan(span.Ticks / iGroups)
        Dim f As FileInfo
        lbxFiles.SelectionMode = SelectionMode.MultiExtended
        MsgBox("Selecting from " & a.CreationTime.ToShortDateString & " to " & (a.CreationTime + chunk).ToShortDateString)
        For Each fileref In list
            f = New FileInfo(fileref)
            If f.CreationTime < a.CreationTime + chunk Then
                For i = 0 To lbxFiles.Items.Count - 1
                    If InStr(UCase(lbxFiles.Items(i)), UCase(fileref)) <> 0 Then
                        lbxFiles.SetSelected(i, True)
                    End If
                Next
            End If
        Next
        'Move to the first subfolder



    End Function

    Private Function SelectSubList(blnRegex As Boolean) As String
        Dim s As String
        If blnRegex Then
            s = InputBox("Enter regex search string")
        Else
            s = InputBox("Enter selection string")
        End If
        If s = "" Then
            Return s
            Exit Function
        End If
        If PFocus = CtrlFocus.Files Or PFocus = CtrlFocus.Tree Then
            SelectFromListbox(lbxFiles, s, blnRegex)
        ElseIf PFocus = CtrlFocus.ShowList Then
            SelectFromListbox(lbxShowList, s, blnRegex)
            'Return s
            'Exit Function
        End If
        If lbxShowList.Items.Count = 0 Then
            '            CollapseShowlist(False)
            '           Showlist = MakeSubList(FileboxContents, s)
        Else
            '          Oldlist = Showlist
            '         lbxShowList.Items.Clear()
            '        Showlist = MakeSubList(Showlist, s)
        End If
        '   FillShowbox(lbxShowList, CurrentFilterState, Showlist)
        lastselection = s

        CancelDisplay()
        Return s
    End Function

    Private Sub ToolStripButton14_Click_1(sender As Object, e As EventArgs)
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




    Private Sub LinearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LinearToolStripMenuItem.Click
        AssignLinear(CurrentFolderPath, iCurrentAlpha, True)
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


    ''' <summary>
    ''' Takes all files in subfolders of di, and places them in di
    ''' </summary>
    Private Sub HarvestCurrent()
        Dim di As New DirectoryInfo(CurrentFolderPath)
        HarvestFolder(di, di, True)
        FillListbox(lbxFiles, di, FileboxContents, False)
        SetControlColours(blnMoveMode)

    End Sub
    Private Sub PromoteFolder()
        Dim di, pa As DirectoryInfo
        di = New DirectoryInfo(CurrentFolderPath)
        pa = di.Parent
        MoveFolder(di.FullName, pa.Parent.FullName, tvMain2, True)
        ChangeFolder(pa.FullName, True)
    End Sub


    Private Sub AddCurrentAndSubfoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddCurrentAndSubfoldersToolStripMenuItem.Click
        ProgressBarOn(1000)
        AddCurrentType(True)
        ProgressBarOff()
    End Sub
    Private Sub AddCurrentFileListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddCurrentFileListToolStripMenuItem.Click
        AddCurrentType(False)
    End Sub




    Private Sub lbxShowList_TextChanged(sender As Object, e As EventArgs) Handles lbxShowList.TextChanged
        If Oldlist.Count <> 0 Then
            If MsgBox("Revert list?") Then
                Showlist = Oldlist
                FillShowbox(lbxShowList, CurrentFilterState, Showlist)
                SetControlColours(blnMoveMode)

            End If
        End If
    End Sub



    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)
        ChosenPlayOrder = cbxOrder.SelectedIndex
        UpdatePlayOrder(Showlist.Count > 0)
    End Sub
    Private Sub LoadListToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles LoadListToolStripMenuItem1.Click
        LoadShowList()

    End Sub

    Private Sub SaveListToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveListToolStripMenuItem1.Click
        SaveShowlist()
    End Sub

    Private Sub LoadListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadButtonFileToolstripMenuItem.Click
        strButtonFile = LoadButtonList()
        KeyAssignmentsRestore(strButtonFile)
    End Sub

    Private Sub SaveListasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveButtonfileasToolStripMenuItem.Click
        SaveButtonlist()

    End Sub

    Private Sub NewListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewButtonFileStripMenuItem.Click
        NewButtonList()
    End Sub



    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        InitialiseColours()

        For i = 0 To 5
            cbxOrder.Items.Add(strPlayOrder(i))
        Next
        ConstructMenuShortcuts()
        ConstructMenutooltips()
    End Sub




    Private Sub ToggleMove()
        blnMoveMode = Not blnMoveMode
        UpdateButtonAppearance()
        SetControlColours(blnMoveMode)
    End Sub

    Private Sub BurstFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BurstFolderToolStripMenuItem.Click
        BurstFolder(New DirectoryInfo(CurrentFolderPath))
    End Sub

    Private Sub LoadFiles_DoWork(sender As Object, e As DoWorkEventArgs) Handles LoadFiles.DoWork
        LoadShowList()
        SetControlColours(blnMoveMode)

    End Sub

    Private Sub tmrUpdateFolderSelection_Tick(sender As Object, e As EventArgs) Handles tmrUpdateFolderSelection.Tick
        'PreferencesSave()

        FillListbox(lbxFiles, New DirectoryInfo(CurrentFolderPath), FileboxContents, blnChooseRandomFile)
        'If lbxFiles.Items.Count = 0 Then tbFiles.Text = "0/" & Str(Showlist.Count)
        'tvMain2.SelectedFolder = CurrentFolderPath 'TODO Dodgy?
        tmrUpdateFolderSelection.Enabled = False
    End Sub


    Private Sub tmrSlowMo_Tick(sender As Object, e As EventArgs) Handles tmrSlowMo.Tick
        MediaAdvance(currentWMP, 1)
        'Throw New Exception

    End Sub

    Private Sub ClearCurrentListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearCurrentListToolStripMenuItem.Click
        ClearShowList()
    End Sub
    Private Sub PicFullScreen() Handles PictureBox1.DoubleClick
        If ShiftDown Then
            blnSecondScreen = True
        Else
            blnSecondScreen = False
        End If
        GoFullScreen(True)

    End Sub
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs)
        Subdividelist(FileboxContents, 10)
    End Sub

    Private Sub TnButton_Click(sender As Object, e As EventArgs)
        ThumbnailsStart()

    End Sub

    Private Shared Sub ThumbnailsStart()
        Dim t As New Thumbnails
        t.ThumbnailHeight = 125
        If PFocus <> CtrlFocus.ShowList Then
            t.FileList = Duplicatelist(FileboxContents)
        Else
            t.FileList = Duplicatelist(Showlist)

        End If
        t.Text = CurrentFolderPath
        t.SetBounds(0, 0, 750, 900)
        t.Show()
    End Sub

    Private Sub toggleMove_Click(sender As Object, e As EventArgs)
        ToggleMove()
    End Sub

    Private Sub RandomStartToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ToggleRandomStartPoint()
    End Sub


    Private Sub tsbOnlyOne_Click(sender As Object, e As EventArgs)
        blnChooseOne = Not blnChooseOne
    End Sub


    Private Sub AlphabeticToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AlphabeticToolStripMenuItem.Click
        AssignAlphabetic(True)
        SaveButtonlist()
    End Sub

    Private Sub ToggleMoveModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleMoveModeToolStripMenuItem.Click
        ToggleMove()

    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
        CurrentFilterState = cbxFilter.SelectedIndex
        SetFilterState()

    End Sub

    Private Sub SingleFilePerFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SingleFilePerFolderToolStripMenuItem.Click
        blnChooseOne = True
        AddFilesToCollection(Showlist, strPicExtensions & strVideoExtensions, True)
        FillShowbox(lbxShowList, FilterState.All, Showlist)
    End Sub

    Private Sub BundleToolStripMenuItem_Click(sender As Object, e As EventArgs)
        BundleFiles(lbxFiles, CurrentFolderPath)
    End Sub

    Private Sub BundleFiles(lbx1 As ListBox, strFolder As String)
        If lbx1.SelectedItems.Count <= 1 Then
            SelectSubList(False)
        End If
        MoveFiles(ListfromListbox(lbx1), CurrentFolderPath, lbx1)
    End Sub

    Private Sub ToolStripButton1_Click_1(sender As Object, e As EventArgs) Handles TreeToolStripMenuItem.Click
        AssignTree(CurrentFolderPath)
        SaveButtonlist()

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs)
        DeadLinksSelect()
    End Sub

    Private Sub DeadLinksSelect()
        If PFocus = CtrlFocus.Files Then
            SelectDeadLinks(lbxFiles)
        ElseIf PFocus = CtrlFocus.ShowList Then
            SelectDeadLinks(lbxShowList)

        End If
    End Sub

    Private Sub btn1_MouseEnter(sender As Object, e As EventArgs) Handles btn1.MouseEnter


    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup

    End Sub

    Private Sub lbxFiles_MouseHover(sender As Object, e As EventArgs) Handles lbxFiles.MouseHover
        'MouseHoverInfo(lbxFiles, ToolTip1)
    End Sub

    Private Sub tmrAutoTrail_Tick(sender As Object, e As EventArgs) Handles tmrAutoTrail.Tick
        'Change the case statements to weight the speeds differently.
        'Should probably be made programmable.
        Dim x As Single = Rnd()
        If x < 0.1 Then
            'Very Slow
            SpeedChange(New KeyEventArgs(KeySpeed1))
            tmrAutoTrail.Interval = Int(Rnd() * 4 * 1000) + 1000
        ElseIf x < 0.3 Then
            'Slow
            SpeedChange(New KeyEventArgs(KeySpeed2))
            tmrAutoTrail.Interval = Int(Rnd() * 6 * 1000) + 1000
        ElseIf x < 0.8 Then
            'slightly slow
            SpeedChange(New KeyEventArgs(KeySpeed3))
            tmrAutoTrail.Interval = Int(Rnd() * 8 * 1000) + 1000
        Else
            'Normal

            HandleKeys(sender, New KeyEventArgs(KeyToggleSpeed))
            tmrAutoTrail.Interval = Int(Rnd() * 2 * 1000) + 750
        End If

        If Int(Rnd() * 12) + 1 = 1 Then
            '    To change the file. 1604
            HandleKeys(sender, New KeyEventArgs(KeyNextFile))

        End If
        ' tmrAutoTrail.Interval = tmrAutoTrail.Interval / MediaDuration * 100

        HandleKeys(sender, New KeyEventArgs(KeyJumpAutoT)) 'This actually does the main job

    End Sub

    Private Sub tmrPumpFiles_Tick(sender As Object, e As EventArgs) Handles tmrPumpFiles.Tick
        If FilePumpList.Count <> 0 Then
            tmrPumpFiles.Enabled = False
            MoveFiles(FilePumpList, lbxFiles)
            FilePumpList.Clear()
        End If
    End Sub

    Private Sub tsbToggleRandomAdvance_Click(sender As Object, e As EventArgs)
        blnRandomAdvance(PFocus) = Not blnRandomAdvance(PFocus)
    End Sub


    Private Sub ConstructMenuShortcuts()
        Dim prefixkeys As Keys
        'CTRL+
        prefixkeys = Keys.Control
        'AddCurrentFileListToolStripMenuItem.ShortcutKeys = KeyAddFile
        LoadListToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.L
        SaveListToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.S
        BundleToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.B
        AddCurrentFileToShowlistToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.F
        'Alt+
        prefixkeys = Keys.Alt
        ToggleRandomAdvanceToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.A
        ToggleMoveModeToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.M
        ToggleJumpToMarkToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.J
        ToggleRandomSelectToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.R
        SlowToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.S
        NormalToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.N
        FastToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.F
        TrailerModeToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.T
        RandomiseNormalToggleToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.Z

        'CTRL+SHIFT
        prefixkeys = Keys.Control + Keys.Shift
        DeleteEmptyFoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.E
        ClearCurrentListToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.C
        NewButtonFileStripMenuItem.ShortcutKeys = prefixkeys + Keys.N
        LoadButtonFileToolstripMenuItem.ShortcutKeys = prefixkeys + Keys.L
        SaveButtonfileasToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.S
        DuplicatesToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.D
        ThumbnailsToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.T
        DeleteEmptyFoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.E
        HarvestFoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.H
        BurstFolderToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.B
        AddCurrentAndSubfoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.A
        'CTRL + ALT
        prefixkeys = Keys.Control + Keys.Alt

        SingleFilePerFolderToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.S
        SlowToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.S
        NormalToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.N
        FastToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.F
        LinearToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.L
        AlphabeticToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.A
        TreeToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.T


        'ToggleRandomStartToolStripMenuItem.ShortcutKeys = Keys.Shift + Keys.D
    End Sub
    Private Sub ConstructMenutooltips()
        AddCurrentFileListToolStripMenuItem.ToolTipText = "Add the current file list to the show list"
        DeleteEmptyFoldersToolStripMenuItem.ToolTipText = "Deletes all empty folders below the currently selected one"
        LoadListToolStripMenuItem1.ToolTipText = "Load a previously saved show list"
        SaveListToolStripMenuItem1.ToolTipText = "Save the current show list"
        BundleToolStripMenuItem.ToolTipText = "Moves all the selected files to a subfolder of their current location"



        ToggleRandomAdvanceToolStripMenuItem.ToolTipText = "Toggles advancing the file randomly or sequentially"
        ToggleMoveModeToolStripMenuItem.ToolTipText = "Toggles move mode, which changes what the f keys do."
        ToggleJumpToMarkToolStripMenuItem.ToolTipText = "Makes movies start at the same fixed point"
        ToggleRandomSelectToolStripMenuItem.ToolTipText = "Toggles whether the first file, or a random file, is selected when the folder is changed."
        ToggleRandomStartToolStripMenuItem.ToolTipText = "Toggles whether movies begin at a random point"
        SlowToolStripMenuItem.ToolTipText = "Sets slowest slideshow speed"
        NormalToolStripMenuItem.ToolTipText = "Sets middle slideshow speed"
        FastToolStripMenuItem.ToolTipText = "Sets fastest slideshow speed"
        TrailerModeToolStripMenuItem.ToolTipText = "Toggles auto-trail mode for movies"
        RandomiseNormalToggleToolStripMenuItem.ToolTipText = "Sets either all, or none of the random functions in one go"

        ClearCurrentListToolStripMenuItem.ToolTipText = "Clears the current show list"
        NewButtonFileStripMenuItem.ToolTipText = "Creates a new button file"
        LoadButtonFileToolstripMenuItem.ToolTipText = "Loads a previously-saved button file"
        SaveButtonfileasToolStripMenuItem.ToolTipText = "Saves the current button file"
        DuplicatesToolStripMenuItem1.ToolTipText = "Opens the duplicates analyser"
        ThumbnailsToolStripMenuItem.ToolTipText = "Creates an interactive page of thumbnails"
        HarvestFoldersToolStripMenuItem.ToolTipText = "Takes all files from subfolders having fewer than a given number of files, and places them in the selected folder"
        BurstFolderToolStripMenuItem.ToolTipText = "Takes all files from the current folder, places them in the parent, and deletes the folder"
        AddCurrentAndSubfoldersToolStripMenuItem.ToolTipText = "Adds all files in the current folder, and all subfolder, and adds to the show list"
        'CTRL + ALT
        SelectDeadLinksToolStripMenuItem.ToolTipText = "Selects any .lnk files which are orphans"
        SingleFilePerFolderToolStripMenuItem.ToolTipText = "Adds a single file from all subfolders to the show list"
        SlowToolStripMenuItem1.ToolTipText = "Sets slowest movie speed"
        NormalToolStripMenuItem1.ToolTipText = "Sets middle movie speed"
        FastToolStripMenuItem1.ToolTipText = "Sets fastest movie speed"
        LinearToolStripMenuItem.ToolTipText = "Assigns next 8 folders to the f buttons"
        AlphabeticToolStripMenuItem.ToolTipText = "Assigns subfolders to the f buttons, alphabetically"
        TreeToolStripMenuItem.ToolTipText = "Assigns subfolders to the f buttons, hierarchically (preference given to higher up tree)"


    End Sub

    Private Sub ToggleJumptoMarkToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleJumpToMarkToolStripMenuItem.Click
        ToggleJumpToMark()
    End Sub

    Private Sub ToggleJumpToMark()
        blnJumpToMark = Not blnJumpToMark
        ToggleJumpToMarkToolStripMenuItem.Checked = blnJumpToMark
    End Sub



    Private Sub ToggleRandomSelectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleRandomSelectToolStripMenuItem.Click
        ToggleRandomSelect()
        ToggleRandomSelectToolStripMenuItem.Checked = blnChooseRandomFile
    End Sub
    Private Sub ToggleRandomSelect()
        blnChooseRandomFile = Not blnChooseRandomFile
        ToggleRandomSelectToolStripMenuItem.Checked = blnChooseRandomFile
    End Sub

    Private Sub ToggleRandomAdvanceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleRandomAdvanceToolStripMenuItem.Click
        ToggleRandomAdvance()
    End Sub

    Private Sub ToggleRandomAdvance()
        blnRandomAdvance(PFocus) = Not blnRandomAdvance(PFocus)
        ToggleRandomAdvanceToolStripMenuItem.Checked = blnRandomAdvance(PFocus)
    End Sub

    Private Sub ToggleRandomStartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleRandomStartToolStripMenuItem.Click
        blnRandomStartPoint = Not blnRandomStartPoint
        ToggleRandomStartToolStripMenuItem.Checked = blnRandomAdvance(PFocus)
    End Sub

    Private Sub BundleToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles BundleToolStripMenuItem.Click
        BundleFiles(lbxFiles, CurrentFolderPath)
    End Sub

    Private Sub ThumbnailsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThumbnailsToolStripMenuItem.Click
        ThumbnailsStart()
    End Sub

    Private Sub TrailerModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TrailerModeToolStripMenuItem.Click
        tmrAutoTrail.Enabled = True
    End Sub



    Private Sub RandomiseNormalToggleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RandomiseNormalToggleToolStripMenuItem.Click
        Static Randomised As Boolean
        Randomised = Not Randomised
        RandomFunctionsToggle(Randomised)


    End Sub

    Public Sub RandomFunctionsToggle(blnRandomise As Boolean)
        blnRandomStartPoint = blnRandomise
        blnRandomStartAlways = blnRandomise
        blnRandomAdvance(PFocus) = blnRandomise
        blnChooseRandomFile = blnRandomise


    End Sub

    Private Sub DuplicatesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DuplicatesToolStripMenuItem1.Click
        FindDuplicates.Show()
    End Sub



    Private Sub SelectDeadLinksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectDeadLinksToolStripMenuItem.Click
        DeadLinksSelect()

    End Sub

    Private Sub PictureBox1_MouseWheel(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseWheel
        PictureFunctions.Mousewheel(PictureBox1, sender, e)
    End Sub


    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        picBlanker = pbxBlanker
        PictureFunctions.MouseMove(PictureBox1, sender, e)
    End Sub
    Public Sub HandleFunctionKeyDown(sender As Object, e As KeyEventArgs)
        Dim i As Byte = e.KeyCode - Keys.F5
        'Move files
        If e.Control Xor blnMoveMode Then 'Move files if CTRL held
            CancelDisplay()
            ChangeWatcherPath(CurrentFolderPath)

            Select Case PFocus
                Case CtrlFocus.Files
                    MoveFiles(ListfromListbox(lbxFiles), strVisibleButtons(i), lbxFiles)
                Case CtrlFocus.Tree
                    MoveFolder(CurrentFolderPath, strVisibleButtons(i), tvMain2, blnMoveMode)
                Case CtrlFocus.ShowList
                    ' If MsgBox("This will move all the showlist files to the folder " & strVisibleButtons(i) & ". Is this what you want?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes Then Exit Sub
                    MoveFiles(ListfromListbox(lbxShowList), strVisibleButtons(i), lbxShowList)
            End Select
        Else
            'Assign buttons
            If e.Shift Or strVisibleButtons(i) = "" Then
                AssignButton(i, iCurrentAlpha, 1, CurrentFolderPath, True) 'Just assign
                If My.Computer.FileSystem.FileExists(strButtonFile) Then
                    KeyAssignmentsStore(strButtonFile)
                Else
                    SaveButtonlist()
                End If
            Else
                'SWITCH folder
                If strVisibleButtons(i) <> CurrentFolderPath Then
                    ChangeFolder(strVisibleButtons(i), True)
                    'CancelDisplay()
                    tvMain2.SelectedFolder = CurrentFolderPath
                ElseIf blnChooseRandomFile Then
                    AdvanceFile(True, True)

                End If
            End If

        End If

        SetControlColours(blnMoveMode)
    End Sub

    Private Sub frmMain_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        Select Case e.Button
            Case MouseButtons.XButton1, MouseButtons.XButton2
                AdvanceFile(e.Button = MouseButtons.XButton2, True)
                e = Nothing
            Case Else
                PicClick(PictureBox1)
        End Select

    End Sub

    Private Sub cbxFilter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxFilter.SelectedIndexChanged
        CurrentFilterState = cbxFilter.SelectedIndex
        If cbxFilter.Focused Then

            SetFilterState()
        End If
    End Sub

    Private Sub cbxOrder_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxOrder.SelectedIndexChanged
        ChosenPlayOrder = cbxOrder.SelectedIndex
        If cbxOrder.Focused Then
            UpdatePlayOrder(Showlist.Count > 0)
        End If
    End Sub

    Private Sub PicFullScreen(sender As Object, e As EventArgs) Handles PictureBox1.DoubleClick

    End Sub

    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        SelectSubList(False)
    End Sub

    Private Sub RegexSearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegexSearchToolStripMenuItem.Click
        SelectSubList(True)
    End Sub

    Private Sub lbxFiles_MouseMove(sender As Object, e As MouseEventArgs) Handles lbxFiles.MouseMove
        'MouseHoverInfo(lbxFiles, ToolTip1)
    End Sub

    Private Sub AddCurrentFileToShowlistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddCurrentFileToShowlistToolStripMenuItem.Click
        AddSingleFileToList(Showlist, strCurrentFilePath)
        FillShowbox(lbxShowList, CurrentFilterState, Showlist)
        CollapseShowlist(False)
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        '   ExtractMetaData(PictureBox1.Image)
    End Sub


    Private Sub StartPointComboBox_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles StartPointComboBox.SelectedIndexChanged
        StartType = StartPointComboBox.SelectedIndex
        Select Case StartType

            Case StartTypes.Particular
                StartPointTrackBar.Enabled = True
                StartPointPercentage = StartPointTrackBar.Value
            Case StartTypes.Beginning
                blnRandomStartAlways = False
            Case StartTypes.NearBeginning


            Case Else
                StartPointTrackBar.Enabled = False

        End Select
    End Sub

    Private Sub StartPointTrackBar_ValueChanged(sender As Object, e As EventArgs) Handles StartPointTrackBar.ValueChanged
        StartPointPercentage = StartPointTrackBar.Value
    End Sub
End Class
