
Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports AxWMPLib
Imports System.Threading
Imports MasaSam.Forms.Controls


Public Class MainForm

    Public Shared ReadOnly Property MyPictures As String

    Public defaultcolour As Color = Color.Aqua
    Public movecolour As Color = Color.Orange
    Public sound As New AxWindowsMediaPlayer
    Public WithEvents FNG As New FileNamesGrouper
    Public WithEvents Random As New RandomHandler
    Public WithEvents NavigateMoveState As New StateHandler()
    Public WithEvents CurrentFilterState As New FilterHandler
    Public WithEvents PlayOrder As New SortHandler
    Public WithEvents StartPoint As New StartPointHandler
    Private WithEvents FM As New FilterMove()
    Private WithEvents DM As New DateMove
    Private WithEvents Att As New Attributes
    Private WithEvents Op As New OrphanFinder
    Public WithEvents SP As New SpeedHandler
    Public DraggedFolder As String

    Public T As Thread


    Public Sub OnFolderMoved(path As String)
        tvMain2.RemoveNode(path)
    End Sub


    'Shared Function Main(ByVal cmdArgs() As String) As Integer
    '    MsgBox("The Main procedure is starting the application.")
    '    Dim returnValue As Integer = 0
    '    ' See if there are any arguments.  
    '    If cmdArgs.Length > 0 Then
    '        For argNum As Integer = 0 To UBound(cmdArgs, 1)
    '            MsgBox(cmdArgs(argNum))
    '            ' Insert code to examine cmdArgs(argNum) and take  
    '            ' appropriate action based on its value.  
    '        Next argNum
    '    End If
    '    ' Insert call to appropriate starting place in your code.  
    '    ' On return, assign appropriate value to returnValue.  
    '    ' 0 usually means successful completion.  
    '    MsgBox("The application is terminating with error level " &
    '         CStr(returnValue) & ".")
    '    Return returnValue
    'End Function
    Public Sub ChangeFavourites()
        'With a maintained list of files for the favourites folder
        'check that a file to be moved is not on it. 
        'if it is, change the link to point to the new position.
    End Sub
    Public Sub OnSpeedChange() Handles SP.SpeedChanged
        If SP.Slideshow Then
            tbSpeed.Text = "Slide Interval=" & SP.Interval
        Else
            tbSpeed.Text = "Speed:" & SP.FrameRate & "fps"
            SwitchSound(Not SP.Fullspeed)

        End If
    End Sub
    Public Sub SwitchSound(slow As Boolean)
        If slow Then
            currentWMP.settings.mute = True
            SoundWMP.settings.mute = False
            SoundWMP.URL = currentWMP.URL
            SoundWMP.Ctlcontrols.currentPosition = currentWMP.Ctlcontrols.currentPosition
            SoundWMP.settings.rate = SP.FrameRate / 30
            If tmrAutoTrail.Enabled Then
            Else

            End If

        Else
            SoundWMP.URL = ""
            currentWMP.settings.mute = False

        End If
    End Sub


    Public Sub OnRenameFolderStart(sender As Object, e As KeyEventArgs) Handles tvMain2.KeyDown
        If e.KeyCode = Keys.F2 Then
            CancelDisplay()
        End If
    End Sub


    Public Sub OnFileListChanged() Handles FM.FilesMoved
        UpdatePlayOrder(False)

    End Sub
    Public Sub OnRandomChanged() Handles Random.RandomChanged
        If Random.All Then
            PlayOrder.State = SortHandler.Order.Random
            StartPoint.State = StartPointHandler.StartTypes.Random
        Else
            PlayOrder.State = SortHandler.Order.Original
            '       StartPoint.State = StartPointHandler.StartTypes.NearBeginning
        End If

        If Random.StartPoint Then
            StartPoint.State = StartPointHandler.StartTypes.Random
        End If
        ToggleRandomAdvanceToolStripMenuItem.Checked = Random.NextSelect
        ToggleRandomSelectToolStripMenuItem.Checked = Random.OnDirChange
        ToggleRandomStartToolStripMenuItem.Checked = Random.StartPoint
        chbNextFile.Checked = Random.NextSelect
        chbInDir.Checked = Random.OnDirChange


    End Sub

    Public Sub OnStartChanged() Handles StartPoint.StateChanged, StartPoint.StartPointChanged
        'blnJumpToMark = True
        MediaMarker = 0
        tbxAbsolute.Text = New TimeSpan(0, 0, StartPoint.StartPoint).ToString("hh\:mm\:ss")
        tbxPercentage.Text = Int(100 * StartPoint.StartPoint / StartPoint.Duration).ToString & "%"
        tbPercentage.Value = StartPoint.Percentage
        tbAbsolute.Maximum = StartPoint.Duration
        '    tbAbsolute.Value = StartPoint.StartPoint
        Select Case StartPoint.State
            Case StartPointHandler.StartTypes.ParticularAbsolute
                tbxPercentage.Enabled = False
                tbxAbsolute.Enabled = True
            Case StartPointHandler.StartTypes.ParticularPercentage
                tbxAbsolute.Enabled = False
                tbxPercentage.Enabled = True
            Case Else
                tbxAbsolute.Enabled = False
                tbxPercentage.Enabled = False
        End Select

        FullScreen.Changing = False
        MediaJumpToMarker()

        cbxStartPoint.SelectedIndex = StartPoint.State
        tbStartpoint.Text = "START:" & StartPoint.Description

    End Sub
    Public Sub OnStateChanged(sender As Object, e As EventArgs) Handles NavigateMoveState.StateChanged, CurrentFilterState.StateChanged, PlayOrder.StateChanged
        'If StartingUpFlag Then Exit Sub
        If sender IsNot NavigateMoveState Then
            UpdatePlayOrder(Showlist.Count > 0)
        Else
        End If
        ReportAction(NavigateMoveState.Instructions)
        cbxOrder.SelectedIndex = PlayOrder.State
        cbxFilter.SelectedIndex = CurrentFilterState.State
        tbRandom.Text = "ORDER:" & UCase(PlayOrder.Description)
        tbFilter.Text = "FILTER:" & UCase(CurrentFilterState.Description)
        tbState.Text = UCase(NavigateMoveState.Description)
        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)
        If lbxFiles.Items.Count = 0 And CurrentFilterState.State <> FilterHandler.FilterState.All Then lbxFiles.Items.Add("If there is nothing showing here, check the filters")
    End Sub
    Private Sub SetControlColours(MainColor As Color, FilterColor As Color)
        tvMain2.BackColor = FilterColor
        tvMain2.HighlightSelectedNodes()
        lbxFiles.BackColor = FilterColor
        lbxShowList.BackColor = FilterColor

        If PFocus = CtrlFocus.Files Then lbxFiles.BackColor = MainColor
        If PFocus = CtrlFocus.Tree Then tvMain2.BackColor = MainColor
        If PFocus = CtrlFocus.ShowList Then lbxShowList.BackColor = MainColor

    End Sub
    ''' <summary>
    ''' Switches between picbox and movie
    ''' </summary>
    ''' <param name="img"></param>
    Private Sub MovietoPic(img As Image)
        PreparePic(currentPicBox, pbxBlanker, img)
        currentPicBox.Visible = True
        currentPicBox.BringToFront()
        currentWMP.Visible = False
        currentWMP.URL = ""
        currentWMP.Visible = False
        SwitchSound(False)
        tbState.Text = ""
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
        ' If blnRandom Then currentWMP.Visible = False
        Try
            currentWMP.URL = Media.MediaPath
        Catch

        End Try

        currentWMP.BringToFront()

        currentPicBox.Visible = False

        If tmrSlideShow.Enabled Then
            blnRestartSlideShowFlag = True
            tmrSlideShow.Enabled = False 'Slideshow stops if movie. Create separate timer for movie slideshows. 
            SP.Slideshow = False
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
        'MsgBox(e.ToString)
    End Sub
    Public Sub CancelDisplay()
        If currentWMP.Visible Then
            'currentWMP.Ctlcontrols.pause()
            currentWMP.URL = ""
        End If
        If currentPicBox.Visible Then
            currentPicBox.Image = Nothing
            GC.Collect()

        End If

        tmrSlideShow.Enabled = False
        SP.Slideshow = False
    End Sub

    Private Sub DeleteFolder(tvw As FileSystemTree, blnConfirm As Boolean)
        With My.Computer.FileSystem
            Dim m As MsgBoxResult = MsgBoxResult.No
            If .DirectoryExists(Media.MediaDirectory) Then
                If NavigateMoveState.State = StateHandler.StateOptions.Navigate Then
                    m = MsgBox("Delete folder " & Media.MediaDirectory & "?", MsgBoxStyle.YesNoCancel)
                End If
                If Not blnConfirm OrElse m = MsgBoxResult.Yes Then
                    Dim f As New DirectoryInfo(Media.MediaDirectory)
                    Try
                        .DeleteDirectory(Media.MediaDirectory, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
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
        If PFocus <> CtrlFocus.ShowList Then
            Dim e = New DirectoryInfo(Media.MediaDirectory)
            FillListbox(lbxFiles, e, False)
        End If
        If blnShowBoxShown And PFocus = CtrlFocus.ShowList Then
            Dim s = lbxShowList.SelectedItem

            Showlist = SetPlayOrder(PlayOrder.State, Showlist)
            lbxShowList.Items.Clear()
            FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)
            Try

                lbxShowList.SelectedIndex = lbxShowList.FindString(s)
            Catch ex As Exception
                Throw New FileNotFoundException
            End Try
        End If

    End Sub



    'Private Function SpeedChange(e As KeyEventArgs, blnTrue As Boolean)
    '   SetMotion(e.KeyCode) 'Alternative speed. Doesn't work at the moment. 
    'End Function
    Private Function SpeedChange(e As KeyEventArgs) As KeyEventArgs

        Dim blnPlaying As Boolean = currentWMP.URL <> ""
        SP.Speed = e.KeyCode - KeySpeed1 'Set slideshow speed if pic showing, and start slideshow
        If Not blnPlaying Then
            'PlaybackSpeed = 30
            tmrSlideShow.Enabled = True
            tmrSlideShow.Interval = SP.Interval
        Else

            PlaybackSpeed = 1000 / SP.FrameRate 'Otherwise, set playback speed 'TODO Options
        End If
        SP.Fullspeed = False

        If e.KeyCode = KeyToggleSpeed Then
            If blnPlaying Then
                If currentWMP.playState = WMPLib.WMPPlayState.wmppsPaused And tmrSlowMo.Enabled = False Then
                    NewPosition = currentWMP.Ctlcontrols.currentPosition
                    currentWMP.settings.rate = 1
                    currentWMP.Ctlcontrols.play()
                    tmrSlowMo.Enabled = False
                    SP.Fullspeed = True
                Else
                    currentWMP.Ctlcontrols.pause()
                    SP.Fullspeed = False
                End If
            Else
                tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
            End If

        Else
            tmrSlowMo.Interval = 1000 / SP.FrameRate
            tmrSlowMo.Enabled = True
            currentWMP.Ctlcontrols.pause()
        End If

        'End If
        'Report
        'blnSpeedRestart = True
        '  tbSpeed.Text = "SPEED (" & SP.FrameRate & " fps)"
        e.SuppressKeyPress = True
        Return e
    End Function

    'Private Shared Sub TweakSpeed(e As KeyEventArgs)
    '    If e.KeyCode = KeySpeed1 Then 'TODO does this work?
    '        If e.Control Then 'increase the extremes if Control held TODO Don't know if this works. 
    '            If e.Shift Then 'decrease if Shift
    '                iSSpeeds(0) = iSSpeeds(0) * 0.9
    '            Else
    '                iSSpeeds(0) = iSSpeeds(0) / 0.9
    '            End If
    '            e.SuppressKeyPress = True

    '        End If

    '    End If
    '    If e.KeyCode = KeySpeed3 Then
    '        If e.Control Then 'increase the extremes if Control held
    '            If e.Shift Then 'decrease if Shift
    '                iSSpeeds(2) = iSSpeeds(2) * 0.9
    '            Else
    '                iSSpeeds(2) = iSSpeeds(2) / 0.9
    '            End If
    '            e.SuppressKeyPress = True
    '        End If

    '    End If

    'End Sub

    Private Sub ToggleRandomStartPoint()
        Random.StartPoint = Not Random.StartPoint
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
            FullScreen.StartPosition = FormStartPosition.CenterScreen

            FullScreen.Location = screen.Bounds.Location + New Point(100, 100)
            currentWMP.Size = screen.Bounds.Size

            FullScreen.Show()

        Else
            SplitterPlace(0.25)
            SetWMP(MainWMP)
            SetPB(PictureBox1)
            FullScreen.Close()
            FullScreen.Changing = False
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
        ReDim Preserve FBCShown(count) 'Boolean list of files shown so far. But why now? Shouldn't this be when the folder is changed?
        lbx.SelectionMode = SelectionMode.One
        If count = 0 Then Exit Sub 'if no filelist, then give up.
        If lbx.SelectedIndex = -1 Then lbx.SelectedIndex = 0

        If lbx.SelectedIndex = 0 And Not blnForward Then
            lbx.SelectedIndex = count - 1
        Else
            If count > 0 Then

                If Random.NextSelect Then

                    Dim i As Int32
                    i = Int(Rnd() * (count))
                    While FBCShown(i) 'To avoid showing one already shown
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
        If NofShown >= count Then 'Re-sets when all shown. Quite nice. 
            ReDim FBCShown(count)
            NofShown = 0
        End If

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
        ElseIf Shiftdown Then
            'Changes the jumpsize while Shift held
            SP.ChangeJump(False, blnBack)
        Else
            'Ordinary jump
            iJumpFactor = 1
        End If
        If blnBack Then
            NewPosition = currentWMP.Ctlcontrols.currentPosition - SP.AbsoluteJump / iJumpFactor
        Else
            NewPosition = currentWMP.Ctlcontrols.currentPosition + SP.AbsoluteJump / iJumpFactor
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
            NewPosition = Math.Min(.currentMedia.duration, .Ctlcontrols.currentPosition + .currentMedia.duration * Math.Sign(e.KeyCode - (KeyBigJumpOn + KeyBigJumpBack) / 2) / (iJumpFactor * SP.FractionalJump))
        End With
        tmrJumpVideo.Enabled = True

    End Sub
    Public Sub JumpRandom(blnAutoTrail As Boolean)
        If Media.MediaType = Filetype.Movie Then

            If Not blnAutoTrail Then
                'Random.StartPoint = True
                NewPosition = (Rnd(1) * (currentWMP.currentMedia.duration))
                tmrJumpVideo.Enabled = True
                tbStartpoint.Text = "START:RANDOM"


            Else
                'MsgBox("Autotrail")
                ToggleAutoTrail()
            End If
        Else
        End If

    End Sub
    Public Sub ToggleAutoTrail()
        Static m As New StartPointHandler
        tmrAutoTrail.Enabled = Not tmrAutoTrail.Enabled
        TrailerModeToolStripMenuItem.Checked = tmrAutoTrail.Enabled
        If tmrAutoTrail.Enabled Then
            m = StartPoint
            StartPoint.State = StartPointHandler.StartTypes.Random
            SwitchSound(True)
        Else
            StartPoint = m
            SwitchSound(False)
            Debug.Print("Normal")
            tmrSlowMo.Enabled = False
            currentWMP.settings.rate = 1
            currentWMP.Ctlcontrols.play()
        End If
    End Sub
    Public Sub HandleKeys(sender As Object, e As KeyEventArgs)
        Me.Cursor = Cursors.WaitCursor
        'MsgBox(e.KeyCode.ToString)
        Select Case e.KeyCode
#Region "Function Keys"
            Case Keys.F2
                CancelDisplay()
            Case Keys.F4 And e.Alt
                Me.Close()
            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                HandleFunctionKeyDown(sender, e)
                e.SuppressKeyPress = True

            Case KeyDelete
                'Use Movefiles with current selected list, and option to delete. 
                CancelDisplay()
                If e.Shift Then
                    DeleteFolder(tvMain2, NavigateMoveState.State = StateHandler.StateOptions.Navigate)
                Else
                    Dim m As List(Of String) = ListfromListbox(lbxFiles)
                    MoveFiles(m, "", lbxFiles)
                End If
#End Region

#Region "Alpha and Numeric"

            Case Keys.Enter And e.Control
                If Media.IsLink Then
                    HighlightCurrent(Media.MediaPath)
                    CurrentFilterState.State = FilterHandler.FilterState.All
                End If
            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9
                'If e.Alt Then Me.frmMain_KeyDown(Me, e)
                If Not e.Control Then
                    ChangeButtonLetter(e)
                Else
                    Select Case e.KeyCode
                        Case Keys.I
                            CreateNewDirectory(Media.MediaDirectory)

                    End Select
                End If
#End Region

#Region "Control Keys"
            Case KeyTraverseTree, KeyTraverseTreeBack
                ''e.suppresskeypress by Treeview behaviour unless focus is elsewhere. 
                ''We want the traverse keys always to work. 
                ''  ControlSetFocus(tvMain2)
                If PFocus = CtrlFocus.Tree Then
                Else
                    tvMain2.tvFiles_KeyDown(sender, e)
                End If
            Case Keys.Left, Keys.Right, Keys.Up, Keys.Down
                If PFocus <> CtrlFocus.ShowList Then
                    ControlSetFocus(tvMain2)
                    tvMain2.tvFiles_KeyDown(sender, e)
                End If

            Case KeyEscape
                CancelDisplay()                'currentPicBox.Image.Dispose()
                tmrAutoTrail.Enabled = False
                'PlayOrder.Toggle()
            Case KeyToggleButtons
                ToggleButtons()
            Case KeyNextFile, KeyPreviousFile, LKeyNextFile, LKeyPreviousFile
                If PFocus <> CtrlFocus.ShowList Then
                    ControlSetFocus(lbxFiles)
                End If
                AdvanceFile(e.KeyCode = KeyNextFile, e.Control)
                e.SuppressKeyPress = True
                tmrSlideShow.Enabled = False

#End Region


#Region "Video Navigation"
            Case KeySmallJumpDown, KeySmallJumpUp, LKeySmallJumpDown, LKeySmallJumpUp
                If e.Alt Then
                    SpeedIncrease(e)
                Else
                    MediaSmallJump(e)
                End If
                e.SuppressKeyPress = True

            Case KeyBigJumpOn, KeyBigJumpBack
                MediaLargeJump(e)
                e.SuppressKeyPress = True

            Case KeyMarkFavourite
                If e.Shift Then
                    NavigateToFavourites()
                Else
                    CreateFavourite(Media.MediaPath)

                End If
            Case KeyJumpToPoint
                Dim m As New StartPointHandler
                m.Duration = MediaDuration
                m.State = StartPointHandler.StartTypes.NearEnd
                NewPosition = m.StartPoint
                tmrJumpVideo.Enabled = True
                e.SuppressKeyPress = True
            Case KeyMarkPoint, LKeyMarkPoint
                'Addmarker(Media.MediaPath)
                If MediaMarker = 0 Then

                    MediaMarker = currentWMP.Ctlcontrols.currentPosition
                Else
                    MediaMarker = 0
                End If
                e.SuppressKeyPress = True

            Case KeyMuteToggle
                currentWMP.settings.mute = Not currentWMP.settings.mute
                e.SuppressKeyPress = True

            Case KeyLoopToggle
                blnLoopPlay = Not blnLoopPlay

            Case KeyJumpAutoT
                If e.Shift Then
                    StartPoint.IncrementState()
                    'blnJumpToMark = True
                Else
                    JumpRandom(e.Control And e.Shift)

                End If

            Case KeyToggleSpeed
                With currentWMP
                    If .URL <> "" Then
                        If .playState = WMPLib.WMPPlayState.wmppsPaused Then
                            SP.Unpause = True
                            tmrSlowMo.Enabled = False
                            .Ctlcontrols.currentPosition = Media.Position
                            '.settings.rate = 1
                            .Ctlcontrols.play()
                            SP.Fullspeed = True
                        Else
                            SP.Unpause = False
                            .Ctlcontrols.pause()
                            SP.Fullspeed = False
                        End If
                    Else
                        tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
                        SP.Slideshow = tmrSlideShow.Enabled
                    End If
                    'blnSpeedRestart = True
                    '  tbSpeed.Text = "SPEED (" & PlaybackSpeed * 100 & "%)"
                End With
            Case KeySpeed1, KeySpeed2, KeySpeed3, KeySpeed3 + Keys.Control
                SpeedChange(e)
#End Region

#Region "Picture Functions"
            Case KeyRotateBack
                RotatePic(currentPicBox, False)
            Case KeyRotate
                RotatePic(currentPicBox, True)
#End Region
#Region "States"


            Case KeyRandomize
                If e.Shift Then
                    PlayOrder.ReverseOrder = Not PlayOrder.ReverseOrder
                Else
                    PlayOrder.IncrementState()
                End If

            Case KeyFilter 'Cycle through listbox filters
                CurrentFilterState.IncrementState()

                e.SuppressKeyPress = True
            Case KeySelect
                SelectSubList(False)

            Case KeyFullscreen
                If ShiftDown Then
                    blnSecondScreen = True
                Else
                    blnSecondScreen = False
                End If
                GoFullScreen(Not blnFullScreen)
                e.SuppressKeyPress = True



            Case KeyMoveToggle
                If e.Shift Then
                    NavigateMoveState.IncrementState()
                Else
                    ToggleMove()
                End If

            Case KeyLoopToggle
                currentWMP.settings.setMode("loop", Not currentWMP.settings.getMode("loop"))
                e.SuppressKeyPress = True
            Case KeyTrueSize
                PicClick(currentPicBox)

#End Region
            Case KeyBackUndo
                If LastFolder.Count > 0 Then

                    Media.MediaDirectory = LastFolder.Pop

                    tvMain2.SelectedFolder = Media.MediaDirectory
                End If

            Case KeyReStartSS
                tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
                SP.Slideshow = tmrSlideShow.Enabled

        End Select
        Me.Cursor = Cursors.Default
        ' e.suppresskeypress = True
    End Sub

    Private Sub SpeedIncrease(e As KeyEventArgs)
        If e.KeyCode = KeySmallJumpUp Then
            SP.IncreaseSpeed()

        Else
            SP.DecreaseSpeed()
        End If
        tmrSlideShow.Interval = SP.Interval
        tmrSlowMo.Interval = 1000 / SP.FrameRate

    End Sub

    Private Sub NavigateToFavourites()
        CurrentFilterState.State = FilterHandler.FilterState.LinkOnly
        ChangeFolder(FavesFolderPath)

        tvMain2.SelectedFolder = Media.MediaDirectory
    End Sub

    Private Sub LoopToggle()
        blnLoopPlay = Not blnLoopPlay
        'currentWMP.= blnLoopPlay
    End Sub
    Private Sub SetFilterState()

    End Sub

    Private Sub SetWMP(tWMP As AxWindowsMediaPlayer)
        With currentWMP
            tWMP.URL = .URL
            ' AxVLCPlugin21.MRL = "file:///" & .URL
            'Media.Position = .Ctlcontrols.currentPosition

            tWMP.Ctlcontrols.currentPosition = .Ctlcontrols.currentPosition
            Debug.Print(tWMP.Ctlcontrols.currentPosition)
            .URL = ""
            currentWMP = tWMP
            Debug.Print(currentWMP.Ctlcontrols.currentPosition)

            '            .Dock = DockStyle.Fill
            .stretchToFit = True
            .Visible = True
            .BringToFront()
            'AxVLCPlugin21.BringToFront()
        End With
        Media.Player = currentWMP

    End Sub
    Private Sub SetPB(tPB As PictureBox)
        tPB.Image = currentPicBox.Image
        'currentPicBox.Image.Dispose()
        currentPicBox.Visible = False
        currentWMP.Visible = False
        currentWMP.URL = ""
        tPB.Visible = True
        currentPicBox = tPB
        Media.Picture = currentPicBox
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
        strPath = Media.MediaPath
        'If strPath is a link, it highlights the link, not the file
        If strPath = "" Then Exit Sub 'Empty
        If Len(strPath) > 247 Then Exit Sub 'Too long

        Dim finfo As New FileInfo(strPath)
        'Change the tree
        Dim s As String = Path.GetDirectoryName(strPath)
        Media.MediaDirectory = s
        If tvMain2.SelectedFolder <> s Then tvMain2.SelectedFolder = s 'Only change tree if it needs changing
        'Select file in filelist
        If lbxFiles.SelectedItem <> strPath Then
            Dim m As Int16 = lbxFiles.FindString(strPath)
            If m <> -1 Then lbxFiles.SelectedIndex = lbxFiles.FindString(strPath)
        End If

        Att.DestinationLabel = lblAttributes
        If Not tmrSlideShow.Enabled And CheckBox1.Checked Then
            Att.UpdateLabel(strPath)
        Else
            Att.Text = ""
        End If

        If Not MasterContainer.Panel2Collapsed Then 'Showlist is visible
            'Select in the showlist unless CTRL held
            If PFocus = CtrlFocus.ShowList AndAlso Not CtrlDown Then
                If lbxShowList.FindString(strPath) <> -1 Then lbxShowList.SelectedIndex = lbxShowList.FindString(strPath)
            End If
        End If

    End Sub

    'Private Sub HighlightListboxSelected(strPath As String, ctrl As ListBox)
    '    With ctrl
    '        For i = 0 To .Items.Count - 1
    '            If .Items(i) = strPath Then
    '                .SetSelected(i, True)
    '                Exit For
    '            End If

    '        Next
    '    End With

    'End Sub


    'Private Sub AddMovies(blnRecurse As Boolean)
    '    AddFilesToCollection(Showlist, VIDEOEXTENSIONS, blnRecurse)
    '    FillShowbox(lbxShowList, FilterHandler.FilterState.All, Showlist)
    'End Sub
    Private Sub AddFiles(blnRecurse As Boolean)
        ProgressBarOn(1000)

        AddFilesToCollection(Showlist, "", blnRecurse)
        FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)
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
        Dim d As New System.IO.DirectoryInfo(Media.MediaDirectory)
        FindAllFilesBelow(d, list, extensions, True, "", True, False)
    End Sub

    Private Sub GlobalInitialise()
        '  InitialiseColours()
        CollapseShowlist(True)
        cbxFilter.Items.Clear()
        cbxOrder.Items.Clear()

        For Each s In CurrentFilterState.Descriptions
            cbxFilter.Items.Add(s)
        Next
        For Each s In PlayOrder.Descriptions
            cbxOrder.Items.Add(s)
        Next
        For Each s In StartPoint.Descriptions
            cbxStartPoint.Items.Add(s)
        Next
        My.Application.SaveMySettingsOnExit = True

        PreferencesGet()
        Randomize()
        AssignExtensionFilters()
        tmrLoadLastFolder.Enabled = True

        tmrPicLoad.Interval = lngInterval * 10
        tmrJumpVideo.Interval = lngInterval
        NavigateMoveState.State = StateHandler.StateOptions.Navigate
        currentWMP = MainWMP
        currentWMP.stretchToFit = True
        currentWMP.uiMode = "FULL"
        currentWMP.Dock = DockStyle.Fill
        Media.Player = currentWMP

        currentPicBox = PictureBox1
        Media.Picture = currentPicBox
        tbPercentage.Enabled = True

        AddHandler FileHandling.FolderMoved, AddressOf OnFolderMoved

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
        If Media.MediaDirectory <> "" Then
            '   WatchStart(Media.MediaDirectory)

        End If
        OnRandomChanged()

        'StartingUpFlag = False
    End Sub
    Public Sub WatchStart(path As String)
        'Exit Sub
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
        Exit Sub 'TODO This gets repeatedly called. 
        If e.ChangeType = System.IO.WatcherChangeTypes.Changed Then
            MsgBox("File " & e.FullPath & " has been modified")
        End If

    End Sub
    'Form Controls
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        GlobalInitialise()



    End Sub



    Public Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control
        UpdateButtonAppearance()
        HandleKeys(sender, e)
        If e.KeyCode = KeyBackUndo Then
            e.SuppressKeyPress = True
        End If
    End Sub
    Private Sub frmMain_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        ShiftDown = e.Shift
        CtrlDown = e.Control
        UpdateButtonAppearance()
        'If e.KeyCode = KeyTraverseTree Or e.KeyCode = KeyTraverseTreeBack Then
        '    ControlSetFocus(tvMain2)
        'End If
        ' MsgBox("Ring")
    End Sub
    Private Sub frmMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        Mysettings.PreferencesSave()
    End Sub


    'Private Sub lbxFiles_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    'Loads a new media file by triggering the PicLoad Timer

    '    Media.MediaPath = lbxFiles.SelectedItem

    '    tmrPicLoad.Interval = lngInterval
    '    tmrPicLoad.Enabled = True
    'End Sub



    'Timers



    Private Sub ClearShowList()
        Showlist.Clear()

        lbxShowList.Items.Clear()
        CollapseShowlist(True)

    End Sub


    Private Sub RemoveFiles(sender As Object, e As EventArgs)
        ' RemoveFilesFromCollection(Showlist, strVideoExtensions)
        FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)

    End Sub

    Private Sub Listbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxFiles.SelectedIndexChanged, lbxShowList.SelectedIndexChanged
        With sender
            Dim i As Long = .SelectedIndex
            tmrPicLoad.Enabled = False
            If i >= 0 Then
                If i = 0 And InStr(.items(i), "filter") Then Exit Sub
                Media.MediaPath = .items(i)
                tmrPicLoad.Enabled = True
            End If
        End With
    End Sub


    Private Sub ShowListToolStripMenuItem_Click(sender As Object, e As EventArgs)
        LoadShowList()

    End Sub

    Private Sub tmrSlideShow_Tick(sender As Object, e As EventArgs) Handles tmrSlideShow.Tick
        SP.Slideshow = True
        AdvanceFile(True, False)
    End Sub

    ' Orientations.



    Private Sub RotatePic(currentPicBox As PictureBox, blnLeft As Boolean)
        If currentPicBox.Image Is Nothing Then Exit Sub
        If Media.MediaType <> Filetype.Pic Then Exit Sub
        With currentPicBox.Image
            If blnLeft Then
                .RotateFlip(RotateFlipType.Rotate90FlipNone)
            Else
                .RotateFlip(RotateFlipType.Rotate270FlipNone)

            End If
            currentPicBox.Refresh()
            Dim finfo As New FileInfo(Media.MediaPath)
            Dim dt As New Date
            'avoid the updating of the write time
            dt = finfo.LastWriteTime
            .Save(Media.MediaPath)
            finfo.LastWriteTime = dt

        End With
    End Sub
    Private Sub ToolStripButton16_Click(sender As Object, e As EventArgs)
        RotatePic(currentPicBox, True)
    End Sub
    Private Function StringList(List As List(Of String), strSearch As String) As List(Of String)
        Application.DoEvents()
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
        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)

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


    'Public Sub UpdateBoxes(strold As String, strnew As String)

    '    lbxFiles.Items.Remove(strold)
    '    lbxShowList.Items.Remove(strold)
    '    If strnew <> "" Then lbxShowList.Items.Add(strnew)
    'End Sub
    Private Sub tmrLoadLastFolder_Tick(sender As Object, e As EventArgs) Handles tmrLoadLastFolder.Tick

        tmrLoadLastFolder.Enabled = False
        'MsgBox("LLF")
        If Media.MediaPath = "" Then Exit Sub
        ' Media.MediaPath="E:\"
        HighlightCurrent(Media.MediaPath)
        'LoadDefaultShowList()
    End Sub


    Private Sub tvMain2_DirectorySelected(sender As Object, e As DirectoryInfoEventArgs) Handles tvMain2.DirectorySelected

        If e.Directory.Exists Then ChangeFolder(e.Directory.FullName)

        tmrUpdateFolderSelection.Enabled = False
        tmrUpdateFolderSelection.Interval = lngInterval * 5
        tmrUpdateFolderSelection.Enabled = True

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

        ' SoundWMP.settings.mute = True
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
        HighlightCurrent(Media.MediaPath)
        fType = Media.MediaType

        Select Case fType
            Case Filetype.Doc

            Case Filetype.Movie, Filetype.Link
                HandleMovie(Random.StartPoint Or StartPoint.State <> StartPointHandler.StartTypes.Beginning)

            Case Filetype.Pic
                Dim img As Image
                If Not currentPicBox.Image Is Nothing Then
                    DisposePic(currentPicBox)
                End If
                img = GetImage(Media.MediaPath)
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
                tbLastFile.Text = "Unhandled file:" & Media.MediaPath

                tmrPicLoad.Enabled = False
                Exit Sub
        End Select
        'MainWMP.fullScreen = blnFullScreen

        tmrPicLoad.Enabled = False
        'If FullScreen.Changing Then FullScreen.Changing = False
    End Sub
    ''' <summary>
    ''' Jumps video to NewPosition
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>

    Private Sub tmrJumpVideo_Tick(sender As Object, e As EventArgs) Handles tmrJumpVideo.Tick
        tmrJumpVideo.Enabled = False

        currentWMP.Ctlcontrols.currentPosition = NewPosition
        SoundWMP.Ctlcontrols.currentPosition = NewPosition

        Console.WriteLine("Newposition " & NewPosition)
        currentWMP.Visible = True
        currentWMP.BringToFront()

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
        ' If strVisibleButtons(i) = "" Then
        m.Folder = Media.MediaDirectory
        'Else
        'm.Folder = strVisibleButtons(i)
        'End If
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
        If Media.MediaPath = "" Then Exit Sub
        Dim f As New FileInfo(Media.MediaPath)
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
        tbFilter.Text = "FILTER:" & UCase(CurrentFilterState.Description)
        tbLastFile.Text = Media.MediaPath
        tbRandom.Text = "ORDER:" & UCase(PlayOrder.Description)
        tbShowfile.Text = "SHOWFILE: " & LastShowList
        '   tbSpeed.Text = tbSpeed.Text = "SPEED (" & PlaybackSpeed & "fps)"
        tbButton.Text = "BUTTONFILE: " & strButtonFile
        tbZoom.Text = iZoomFactor
        If Random.StartPoint Then
            tbStartpoint.Text = "START:RANDOM"
        Else
            tbStartpoint.Text = "START:NORMAL"
        End If

    End Sub


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
        AssignLinear(Media.MediaDirectory, iCurrentAlpha, True)
    End Sub



    Private Sub DeleteEmptyFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteEmptyFoldersToolStripMenuItem.Click
        If Not MsgBox("This deletes all empty directories", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            Exit Sub
        Else
            DeleteEmptyFolders(New DirectoryInfo(Media.MediaDirectory), True)
        End If

    End Sub

    Private Sub HarvestFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HarvestFoldersToolStripMenuItem.Click
        HarvestCurrent()
        tvMain2.RefreshTree(Media.MediaDirectory)

    End Sub


    ''' <summary>
    ''' Takes all files in subfolders of di, and places them in di
    ''' </summary>
    Private Sub HarvestCurrent()
        Dim di As New DirectoryInfo(Media.MediaDirectory)
        HarvestFolder(di, True, False)
        FillListbox(lbxFiles, di, False)
        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)
        'SetControlColours(blnMoveMode)

    End Sub



    Private Sub AddCurrentAndSubfoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddCurrentAndSubfoldersToolStripMenuItem.Click
        ProgressBarOn(1000)
        AddCurrentType(True)
        ProgressBarOff()
    End Sub
    Private Sub AddCurrentFileListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddCurrentFileListToolStripMenuItem.Click
        AddCurrentType(False)
    End Sub








    Private Sub LoadListToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        LoadShowList()

    End Sub

    Private Sub SaveListToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveButtonfileasToolStripMenuItem.Click
        SaveButtonlist()
    End Sub

    Private Sub LoadListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadButtonFileToolstripMenuItem.Click

        strButtonFile = LoadButtonList()
        KeyAssignmentsRestore(strButtonFile)
    End Sub

    Private Sub SaveListasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        SaveButtonlist()

    End Sub

    Private Sub NewListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewButtonFileStripMenuItem.Click
        NewButtonList()
    End Sub



    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load


        ConstructMenuShortcuts()
        ConstructMenutooltips()
    End Sub




    Private Sub ToggleMove()
        NavigateMoveState.ToggleState()

        UpdateButtonAppearance()
    End Sub

    Private Sub BurstFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BurstFolderToolStripMenuItem.Click
        BurstFolder(New DirectoryInfo(Media.MediaDirectory))
        tvMain2.RefreshTree(Media.MediaDirectory)
    End Sub

    Private Sub tmrUpdateFolderSelection_Tick(sender As Object, e As EventArgs) Handles tmrUpdateFolderSelection.Tick
        'PreferencesSave()
        ListBox1.Items.Clear()

        FillListbox(lbxFiles, New DirectoryInfo(Media.MediaDirectory), Random.OnDirChange)

        'If lbxFiles.Items.Count = 0 Then tbFiles.Text = "0/" & Str(Showlist.Count)
        'tvMain2.SelectedFolder = Media.MediaDirectory 'TODO Dodgy?
        tmrUpdateFolderSelection.Enabled = False
    End Sub

    Private Sub OnFilenamesParsed() Handles FNG.WordsParsed

        For Each s In FNG.WordList
            ListBox1.Items.Add(s)

        Next
        Dim i As Integer = 0

        For Each g In FNG.Groups
            Console.WriteLine(FNG.GroupNames(i))

            i += 1
            For Each m In g
                Console.WriteLine(m)
            Next
            Console.WriteLine()
        Next

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


    Private Shared Sub ThumbnailsStart()
        Dim t As New Thumbnails
        t.ThumbnailHeight = 70
        If PFocus <> CtrlFocus.ShowList Then
            t.List = Duplicatelist(FileboxContents)
        Else
            t.List = Duplicatelist(Showlist)

        End If
        't.LayoutPanel = Thumbnails.FlowLayoutPanel1
        t.Text = Media.MediaDirectory
        t.SetBounds(0, 0, 750, 900)
        t.Show()

    End Sub

    Private Sub toggleMove_Click(sender As Object, e As EventArgs)
        ToggleMove()
    End Sub

    Private Sub RandomStartToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ToggleRandomStartPoint()
    End Sub


    'Private Sub tsbOnlyOne_Click(sender As Object, e As EventArgs)
    '    blnChooseOne = Not blnChooseOne
    'End Sub


    Private Sub AlphabeticToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AlphabeticToolStripMenuItem.Click
        AssignAlphabetic(True)
        SaveButtonlist()
    End Sub

    Private Sub ToggleMoveModeToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ToggleMove()

    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
        CurrentFilterState.State = cbxFilter.SelectedIndex

    End Sub

    Private Sub SingleFilePerFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SingleFilePerFolderToolStripMenuItem.Click
        blnChooseOne = True
        AddFilesToCollection(Showlist, PICEXTENSIONS & VIDEOEXTENSIONS, True)
        FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)
    End Sub

    'Private Sub BundleToolStripMenuItem_Click(sender As Object, e As EventArgs)
    '    BundleFiles(lbxFiles, Media.MediaDirectory)
    'End Sub

    Public Sub BundleFiles(lbx1 As ListBox, strFolder As String)
        If lbx1.SelectedItems.Count <= 1 Then
            SelectSubList(False)
        End If
        MoveFiles(ListfromListbox(lbx1), strFolder, lbx1)
        tvMain2.RefreshTree(Media.MediaDirectory)
        tmrUpdateFileList.Enabled = True
    End Sub
    Private Sub Groupfiles(ByVal m As FileNamesGrouper)
        For i As Integer = 0 To m.Groups.Count - 1
            MoveFiles(m.Groups.Item(i), Media.MediaDirectory & "\" & m.GroupNames.Item(i))
        Next
        tvMain2.RefreshTree(Media.MediaDirectory)
        tmrUpdateFileList.Enabled = True
    End Sub

    Private Sub ToolStripButton1_Click_1(sender As Object, e As EventArgs) Handles TreeToolStripMenuItem.Click
        AssignTree(Media.MediaDirectory)
        SaveButtonlist()

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs)
        DeadLinksSelect()
    End Sub

    Private Sub DeadLinksSelect()
        Dim s As New List(Of String)
        If PFocus <> CtrlFocus.ShowList Then
            SelectDeadLinks(lbxFiles)
            For Each m In lbxFiles.SelectedItems
                s.Add(m.ToString)
            Next

        Else
            SelectDeadLinks(lbxShowList)

            For Each m In lbxFiles.SelectedItems
                s.Add(m.ToString)
            Next

        End If





    End Sub



    Private Sub lbxFiles_MouseHover(sender As Object, e As EventArgs) Handles lbxFiles.MouseHover
        MouseHoverInfo(lbxFiles, ToolTip1)
    End Sub

    Private Sub tmrAutoTrail_Tick(sender As Object, e As EventArgs) Handles tmrAutoTrail.Tick
        'Change the case statements to weight the speeds differently.
        'Should probably be made programmable.
        Dim trail As New TrailMode
        Dim vs, s, ss, n As Byte
        Dim x As Single = Rnd()
        Dim ch As Integer = 15

        vs = TrackBar1.Value
        s = TrackBar2.Value
        ss = TrackBar3.Value
        n = TrackBar4.Value

        With trail
            .AdvanceChance = ch
            .RandomTimes = {vs, s, ss, n}

        End With
        ' trail.EqualiseSpeeds(14, 6)
        Dim speedkeys = {KeySpeed1, KeySpeed2, KeySpeed3}
        Dim i As Byte = trail.ChosenSpeed(x)
        Select Case i
            Case 0, 1, 2
                SP.FrameRate = trail.Speeds(i)
                SpeedChange(New KeyEventArgs(speedkeys(i)))
                tmrAutoTrail.Interval = Int(Rnd() * trail.RandomTimes(i) * 500) + 500
            Case 3
                'Normal
                Debug.Print("Normal")
                tmrSlowMo.Enabled = False
                currentWMP.settings.rate = 1
                currentWMP.Ctlcontrols.play()
                tmrAutoTrail.Interval = Int(Rnd() * trail.RandomTimes(i) * 500) + 500

        End Select
        'If x < 0.2 Then
        '    Debug.Print("Very Slow")
        '    SpeedChange(New KeyEventArgs(KeySpeed1))
        '    tmrAutoTrail.Interval = Int(Rnd() * vs * 500) + 500
        'ElseIf x < 0.4 Then
        '    Debug.Print("Slow")
        '    'Slow
        '    SpeedChange(New KeyEventArgs(KeySpeed2))
        '    tmrAutoTrail.Interval = Int(Rnd() * s * 500) + 500
        'ElseIf x < 0.75 Then
        '    Debug.Print("Slightly Slow")

        '    'slightly slow
        '    SpeedChange(New KeyEventArgs(KeySpeed3))
        '    tmrAutoTrail.Interval = Int(Rnd() * ss * 500) + 500
        'Else
        '    'Normal
        '    Debug.Print("Normal")
        '    tmrSlowMo.Enabled = False
        '    currentWMP.settings.rate = 1
        '    currentWMP.Ctlcontrols.play()
        '    tmrAutoTrail.Interval = Int(Rnd() * n * 500) + 750
        'End If

        If Int(Rnd() * trail.AdvanceChance) + 1 = 1 Then
            '    To change the file.
            Debug.Print("Change file")

            HandleKeys(sender, New KeyEventArgs(KeyNextFile))

        End If

        HandleKeys(sender, New KeyEventArgs(KeyJumpAutoT)) 'This actually does the main job

    End Sub

    Private Sub tmrPumpFiles_Tick(sender As Object, e As EventArgs) Handles tmrPumpFiles.Tick
        If FilePumpList.Count <> 0 Then
            tmrPumpFiles.Enabled = False
            MoveFiles(FilePumpList, lbxFiles)
            FilePumpList.Clear()
        End If
    End Sub

    'Private Sub tsbToggleRandomAdvance_Click(sender As Object, e As EventArgs)
    '    blnRandomAdvance(PFocus) = Not blnRandomAdvance(PFocus)
    'End Sub


    Private Sub ConstructMenuShortcuts()
        'Dim prefixkeys As Keys
        ''CTRL+
        'prefixkeys = Keys.Control
        ''AddCurrentFileListToolStripMenuItem.ShortcutKeys = KeyAddFile
        'LoadListToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.L
        'SaveListToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.S
        'BundleToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.B
        'AddCurrentFileToShowlistToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.F
        ''Alt+
        'prefixkeys = Keys.Alt
        'ToggleRandomAdvanceToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.A
        'ToggleMoveModeToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.M
        'ToggleJumpToMarkToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.J
        'ToggleRandomSelectToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.R
        'SlowToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.S
        'NormalToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.N
        'FastToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.F
        'TrailerModeToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.T
        'RandomiseNormalToggleToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.Z

        ''CTRL+SHIFT
        'prefixkeys = Keys.Control + Keys.Shift
        'DeleteEmptyFoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.E
        'ClearCurrentListToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.C
        'NewButtonFileStripMenuItem.ShortcutKeys = prefixkeys + Keys.N
        'LoadButtonFileToolstripMenuItem.ShortcutKeys = prefixkeys + Keys.L
        'SaveButtonfileasToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.S
        'DuplicatesToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.D
        'ThumbnailsToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.T
        'DeleteEmptyFoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.E
        'HarvestFoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.H
        'BurstFolderToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.B
        'AddCurrentAndSubfoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.A
        ''CTRL + ALT
        'prefixkeys = Keys.Control + Keys.Alt

        'SingleFilePerFolderToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.S
        'SlowToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.S
        'NormalToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.N
        'FastToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.F
        'LinearToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.L
        'AlphabeticToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.A
        'TreeToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.T


        'ToggleRandomStartToolStripMenuItem.ShortcutKeys = Keys.Shift + Keys.D
    End Sub
    Private Sub ConstructMenutooltips()
        AddCurrentFileListToolStripMenuItem.ToolTipText = "Add the current file list to the show list"
        DeleteEmptyFoldersToolStripMenuItem.ToolTipText = "Deletes all empty folders below the currently selected one"
        LoadListToolStripMenuItem1.ToolTipText = "Load a previously saved show list"
        SaveListToolStripMenuItem1.ToolTipText = "Save the current show list"
        BundleToolStripMenuItem.ToolTipText = "Moves all the selected files to a subfolder of their current location"



        ToggleRandomAdvanceToolStripMenuItem.ToolTipText = "Toggles advancing the file randomly or sequentially"
        'ToggleMoveModeToolStripMenuItem.ToolTipText = "Toggles move mode, which changes what the f keys do."
        ToggleJumpToMarkToolStripMenuItem.ToolTipText = "Makes movies start at the same fixed point"
        ToggleRandomSelectToolStripMenuItem.ToolTipText = "Toggles whether the first file, or a random file, is selected when the folder is changed."
        ToggleRandomStartToolStripMenuItem.ToolTipText = "Toggles whether movies begin at a random point"
        'SlowToolStripMenuItem.ToolTipText = "Sets slowest slideshow speed"
        'NormalToolStripMenuItem.ToolTipText = "Sets middle slideshow speed"
        'FastToolStripMenuItem.ToolTipText = "Sets fastest slideshow speed"
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
        'SlowToolStripMenuItem1.ToolTipText = "Sets slowest movie speed"
        'NormalToolStripMenuItem1.ToolTipText = "Sets middle movie speed"
        'FastToolStripMenuItem1.ToolTipText = "Sets fastest movie speed"
        LinearToolStripMenuItem.ToolTipText = "Assigns next 8 folders to the f buttons"
        AlphabeticToolStripMenuItem.ToolTipText = "Assigns subfolders to the f buttons, alphabetically"
        TreeToolStripMenuItem.ToolTipText = "Assigns subfolders to the f buttons, hierarchically (preference given to higher up tree)"


    End Sub



    Private Sub ToggleRandomSelectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleRandomSelectToolStripMenuItem.Click
        ToggleRandomSelect()
    End Sub
    Private Sub ToggleRandomSelect()
        Random.OnDirChange = Not Random.OnDirChange
    End Sub

    Private Sub ToggleRandomAdvanceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleRandomAdvanceToolStripMenuItem.Click
        ToggleRandomAdvance()
    End Sub

    Private Sub ToggleRandomAdvance()
        Random.NextSelect = Not Random.NextSelect
        '  blnRandomAdvance(PFocus) = Random.NextSelect
    End Sub

    Private Sub ToggleRandomStartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleRandomStartToolStripMenuItem.Click
        Random.StartPoint = Not Random.StartPoint
    End Sub

    Private Sub BundleToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles BundleToolStripMenuItem.Click
        BundleFiles(lbxFiles, Media.MediaDirectory)
    End Sub

    Private Sub ThumbnailsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThumbnailsToolStripMenuItem.Click
        ThumbnailsStart()
    End Sub

    Private Sub TrailerModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TrailerModeToolStripMenuItem.Click
        tmrAutoTrail.Enabled = True
    End Sub



    Private Sub RandomiseNormalToggleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RandomiseNormalToggleToolStripMenuItem.Click
        ' Static Randomised As Boolean
        'Randomised = Not Randomised
        RandomFunctionsToggle()


    End Sub

    Public Sub RandomFunctionsToggle()
        Random.All = Not Random.All
        'blnRandomAdvance(PFocus) = Random.NextSelect


    End Sub

    Private Sub DuplicatesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DuplicatesToolStripMenuItem1.Click
        'Dim s As List(Of List(Of String))
        'Duplicates.InputList = FileboxContents
        's = Duplicates.Duplicates
        'For Each m In s
        '    Debug.Print("")
        '    For Each rs In m
        '        Debug.Print(rs)
        '    Next
        'Next
        With FindDuplicates

            If PFocus = CtrlFocus.ShowList Then
                .List = Showlist
            Else

                .List = FileboxContents


            End If
            If .DuplicatesCount > 0 Then
                .ThumbnailHeight = 100
                .Show()
            Else
                MsgBox("No duplicates found.")
            End If
        End With
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
        Dim s As StateHandler.StateOptions = NavigateMoveState.State
        'Move files
        CancelDisplay()
        If e.Shift And e.Control And e.Alt Or strVisibleButtons(i) = "" Then
            AssignButton(i, iCurrentAlpha, 1, Media.MediaDirectory, True) 'Just assign
            If My.Computer.FileSystem.FileExists(strButtonFile) Then
                KeyAssignmentsStore(strButtonFile)
            Else
                SaveButtonlist()
            End If
            Exit Sub
        End If
        Select Case s
            Case StateHandler.StateOptions.Move, StateHandler.StateOptions.Copy
                'ChangeWatcherPath(CurrentFolderPath)
                If e.Control And e.Shift Then
                    If strVisibleButtons(i) <> Media.MediaDirectory Then
                        ChangeFolder(strVisibleButtons(i))
                        'CancelDisplay()
                        tvMain2.SelectedFolder = Media.MediaDirectory
                    ElseIf Random.OnDirChange Then
                        AdvanceFile(True, True)

                    End If
                ElseIf e.Shift Then
                    Dim path As String = Media.MediaDirectory
                    OnFolderMoved(Media.MediaDirectory)

                    T = New Thread(New ThreadStart(Sub() MoveFolder(path, strVisibleButtons(i), True)))
                    T.IsBackground = True
                    T.SetApartmentState(ApartmentState.STA)

                    T.Start()


                Else
                    MoveFiles(ListfromListbox(lbxFiles), strVisibleButtons(i), lbxFiles)
                    If lbxShowList.Visible Then
                        MoveFiles(ListfromListbox(lbxShowList), strVisibleButtons(i), lbxShowList)
                    End If
                    ' If MsgBox("This will move all the showlist files to the folder " & strVisibleButtons(i) & ". Is this what you want?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes Then Exit Sub
                End If
            Case StateHandler.StateOptions.Navigate

                If e.Shift And e.Control And strVisibleButtons(i) <> "" Then
                    OnFolderMoved(Media.MediaDirectory)

                    T = New Thread(New ThreadStart(Sub() MoveFolder(Media.MediaDirectory, strVisibleButtons(i), True)))
                    T.IsBackground = True
                    T.SetApartmentState(ApartmentState.STA)

                    T.Start()

                ElseIf e.Shift Then
                    MoveFiles(ListfromListbox(lbxFiles), strVisibleButtons(i), lbxFiles)

                Else
                    'SWITCH folder
                    If strVisibleButtons(i) <> Media.MediaDirectory Then
                        ChangeFolder(strVisibleButtons(i))
                        'CancelDisplay()
                        tvMain2.SelectedFolder = Media.MediaDirectory
                    ElseIf Random.OnDirChange Then
                        AdvanceFile(True, True)

                    End If
                End If
        End Select


        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)

        'SetControlColours(blnMoveMode)
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
        If CurrentFilterState.State <> cbxFilter.SelectedIndex Then
            CurrentFilterState.State = cbxFilter.SelectedIndex
        End If
    End Sub

    Private Sub cbxOrder_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxOrder.SelectedIndexChanged
        If PlayOrder.State <> cbxOrder.SelectedIndex Then
            PlayOrder.State = cbxOrder.SelectedIndex
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
        AddCurrentFileToShowList()
    End Sub

    Private Sub AddCurrentFileToShowList()
        AddSingleFileToList(Showlist, Media.MediaPath)
        FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)
        CollapseShowlist(False)
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        '   ExtractMetaData(PictureBox1.Image)
    End Sub


    Private Sub StartPointComboBox_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cbxStartPoint.SelectedIndexChanged
        If StartPoint.State <> cbxStartPoint.SelectedIndex Then
            StartPoint.State = cbxStartPoint.SelectedIndex


        End If


    End Sub

    Private Sub tbPercentage_ValueChanged(sender As Object, e As EventArgs) Handles tbPercentage.ValueChanged
        'StartPoint.State = StartPointHandler.StartTypes.ParticularPercentage
        StartPoint.Percentage = tbPercentage.Value

    End Sub








    Private Sub AbsoluteTrackBar_ValueChanged(sender As Object, e As EventArgs) Handles tbAbsolute.ValueChanged
        tbAbsolute.Maximum = MediaDuration
        tbAbsolute.TickFrequency = tbAbsolute.Maximum / 25
        StartPoint.Absolute = tbAbsolute.Value


    End Sub


    Private Sub chbNextFile_CheckedChanged(sender As Object, e As EventArgs) Handles chbNextFile.CheckedChanged
        Random.NextSelect = chbNextFile.Checked
    End Sub

    Private Sub chbInDir_CheckedChanged(sender As Object, e As EventArgs) Handles chbInDir.CheckedChanged
        Random.OnDirChange = chbInDir.Checked
    End Sub


    Private Sub FilterMoveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FilterMoveToolStripMenuItem.Click
        '   FM.Recursive = False
        FM.FilterMoveFiles(Media.MediaDirectory, False)
    End Sub

    Private Sub FilterMoveRecursiveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FilterMoveRecursiveToolStripMenuItem.Click
        FM.FilterMoveFiles(Media.MediaDirectory, True)
        ReportAction(Format("Filtering {0}", Media.MediaDirectory))

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        DM.FilterByDate(New FileInfo(Media.MediaPath).DirectoryName, False, DateMove.DMY.Year)
        ReportAction("Filtering by Year")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim GRP As New Grouper(Media.MediaDirectory)
        GRP.FilesPerList = 25
        Dim lists As List(Of List(Of String)) = GRP.Sublists
        blnSuppressCreate = True
        Dim i = 0
        For Each list In lists
            ReportAction("Moving " & list.ToString & " to " & Str(i))

            MoveFiles(list, Media.MediaDirectory & "\" & Str(i), lbxFiles)
            i += 1
        Next
        tvMain2.RefreshTree(Media.MediaDirectory)
    End Sub

    Private Sub ReUniteFavesLinks()

        Dim f As New IO.DirectoryInfo(FavesFolderPath)
        Dim r As New List(Of String)
        For Each m In f.GetFiles
            r.Add(m.FullName)
        Next
        Op.OrphanList = r
    End Sub
    Public Sub ReportAction(Msg As String)
        Console.WriteLine(Msg)
        Label4.Text = Msg
    End Sub

    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs)

    End Sub
    Private Sub Filter(index As DateMove.DMY)
        CancelDisplay()
        DM.FilterByDate(Media.MediaDirectory, False, index)
        tvMain2.RefreshTree(Media.MediaDirectory)
    End Sub




    Private Sub tbAbsolute_MouseUp(sender As Object, e As MouseEventArgs) Handles tbAbsolute.MouseUp
        StartPoint.State = StartPointHandler.StartTypes.ParticularAbsolute
        StartPoint.StartPoint = tbAbsolute.Value
        tbxAbsolute.Text = New TimeSpan(0, 0, tbAbsolute.Value).ToString("hh\:mm\:ss")
        tbxPercentage.Text = Str(StartPoint.Percentage) & "%"
        ' MediaJumpToMarker()

    End Sub

    Private Sub tbPercentage_MouseUp(sender As Object, e As MouseEventArgs) Handles tbPercentage.MouseUp
        StartPoint.State = StartPointHandler.StartTypes.ParticularPercentage
        StartPoint.StartPoint = tbPercentage.Value
        tbxAbsolute.Text = New TimeSpan(0, 0, tbAbsolute.Value).ToString("hh\:mm\:ss")
        tbxPercentage.Text = Str(StartPoint.Percentage) & "%"

    End Sub

    Private Sub ReclaimDeadLinksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReclaimDeadLinksToolStripMenuItem.Click
        ReUniteFavesLinks()

    End Sub

    Private Sub tvMain2_MouseDown(sender As Object, e As MouseEventArgs) Handles tvMain2.MouseDown
        tvMain2.DoDragDrop(DraggedFolder, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub

    Private Sub tvMain2_DragEnter(sender As Object, e As DragEventArgs) Handles tvMain2.DragEnter
        If (e.Data.GetDataPresent(DataFormats.Text)) Then
            If (e.KeyState And 8) = 8 Then
                e.Effect = DragDropEffects.Copy
            Else
                e.Effect = DragDropEffects.Move
            End If
        End If
    End Sub

    Private Sub RecursiveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RecursiveToolStripMenuItem.Click
        HarvestBelow(New DirectoryInfo(Media.MediaDirectory))
    End Sub

    Private Sub BySizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        DM.FilterBySize(Media.MediaDirectory, False)
        tvMain2.RefreshTree(Media.MediaDirectory)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PromoteFolder(New DirectoryInfo(Media.MediaDirectory))
    End Sub


    Private Sub ShowlistToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ShowlistToolStripMenuItem.Click
        Dim x As New Listform
        x.Show
    End Sub

    Private Sub GroupingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click
        '    FNG.Filenames = New IO.DirectoryInfo(Media.MediaDirectory).GetFiles
        Groupfiles(FNG)
        Exit Sub

    End Sub

    Private Sub ButtonFormToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ButtonFormToolStripMenuItem.Click
        ButtonForm.Show()
    End Sub


    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Filter(DateMove.DMY.Month)
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Filter(DateMove.DMY.Day)
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Filter(DateMove.DMY.Year)
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Filter(DateMove.DMY.Minute)

    End Sub

    Private Sub tmrUpdateFileList_Tick(sender As Object, e As EventArgs) Handles tmrUpdateFileList.Tick
        FillListbox(lbxFiles, New DirectoryInfo(Media.MediaDirectory), Random.OnDirChange)
        tmrUpdateFileList.Enabled = False
    End Sub

    Private Sub lbxFiles_MouseDown(sender As Object, e As MouseEventArgs) Handles lbxFiles.MouseDown

    End Sub

    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem8.Click
        DM.FilterByCalendar(New DirectoryInfo(Media.MediaDirectory))
    End Sub

    Private Sub ToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem9.Click
        DM.FilterByType(Media.MediaDirectory)
    End Sub

End Class
