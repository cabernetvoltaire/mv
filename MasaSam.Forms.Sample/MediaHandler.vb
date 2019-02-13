﻿Imports AxWMPLib
Public Class MediaHandler

    Public Event MediaFinished(ByVal sender As Object, ByVal e As EventArgs)
    Public Event StartChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public WithEvents PositionUpdater As New Timer
    Private DefaultFile As String = "C:\exiftools.exe"
    Private mPaused As Boolean
    Public WithEvents StartPoint As New StartPointHandler
    Public Property Speed As New SpeedHandler

    Private mType As Filetype
    Public Property MediaType() As Filetype
        Get
            If mMediaPath = "" Then
            Else

                mType = FindType(mMediaPath)
            End If
            If mType = Filetype.Link Then
                mIsLink = True
                mLinkPath = LinkTarget(mMediaPath)
                mType = FindType(mLinkPath)

            Else
                IsLink = False
            End If
            Return mType
        End Get
        Set(ByVal value As Filetype)
            mType = value
        End Set
    End Property

    Private mPicBox As New PictureBox
    Public Property Picture() As PictureBox
        Get
            Return mPicBox
        End Get
        Set(ByVal value As PictureBox)
            mPicBox = value
        End Set
    End Property

    Private WithEvents mPlayer As New AxWMPLib.AxWindowsMediaPlayer
    Public Property Player() As AxWMPLib.AxWindowsMediaPlayer
        Get
            Return mPlayer
        End Get
        Set(ByVal value As AxWMPLib.AxWindowsMediaPlayer)
            mPlayer = value

            mPlayer.settings.enableErrorDialogs = False
            mPlayer.settings.autoStart = True
        End Set
    End Property
    Private mPlayPosition As Long
    Public Property Position() As Long
        Get
            '   mPlayPosition = mPlayer.Ctlcontrols.currentPosition
            Return mPlayPosition
        End Get
        Set(ByVal value As Long)
            mPlayPosition = value
            mPlayer.Ctlcontrols.currentPosition = mPlayPosition
        End Set
    End Property

    Private mDuration As Long
    Public Property Duration() As Long
        Get
            ' mDuration = mPlayer.currentMedia.duration
            Return mDuration
        End Get
        Set(value As Long)
            mDuration = value
            StartPoint.Duration = value
        End Set
    End Property
    Private mFrameRate As Int32
    Public Property FrameRate() As Int32
        Get
            Return mFrameRate
        End Get
        Set(ByVal value As Int32)
            mFrameRate = mPlayer.network.frameRate
        End Set
    End Property
    Private mMarkers As List(Of String)
    Public Property Markers() As List(Of String)
        Get
            Return mMarkers
        End Get
        Set(ByVal value As List(Of String))
            mMarkers = value
        End Set
    End Property
    Private mMediaPath As String
    Public Property MediaPath() As String
        Get
            Return mMediaPath
        End Get
        Set(ByVal value As String)
            'If path changes, we need to check it exists, and if so, change stored directory as well, 
            'And raise a media changed event. 
            Dim b As String = mMediaPath
            If value = b Then
            ElseIf value = "" Then
                mMediaPath = DefaultFile
                mMediaDirectory = New IO.FileInfo(mMediaPath).Directory.FullName

                RaiseEvent MediaChanged(Me, New EventArgs)

            Else

                mMediaPath = value
                mType = FindType(value)
                Dim f As New IO.FileInfo(value)
                If f.Exists Then
                    If mType = Filetype.Link Then
                        mIsLink = True
                        mLinkPath = LinkTarget(f.FullName)
                    Else
                        mIsLink = False
                        mLinkPath = ""
                    End If
                    Me.LoadMedia()

                    mMediaDirectory = f.Directory.FullName
                Else
                    mMediaPath = DefaultFile
                    mMediaDirectory = New IO.FileInfo(mMediaPath).Directory.FullName
                End If

                RaiseEvent MediaChanged(Me, New EventArgs)
            End If


        End Set
    End Property
    Public Sub New()
        PositionUpdater.Interval = 500
        PositionUpdater.Enabled = False
        '     StartPoint = Media.StartPoint
    End Sub
    Private mMediaDirectory As String
    Public ReadOnly Property MediaDirectory() As String
        Get
            'Dim f As New IO.FileInfo(mMediaPath)
            ' mMediaDirectory = f.Directory.FullName
            Return mMediaDirectory
        End Get
        'Set(ByVal value As String)
        '    If value <> mMediaDirectory Then
        '        mMediaDirectory = value
        '        RaiseEvent MediaChanged(Me, New EventArgs)
        '    End If
        'End Set
    End Property

    Private mIsLink As Boolean = False
    Public Property IsLink() As Boolean
        Get
            Return mIsLink
        End Get
        Set(ByVal value As Boolean)
            mIsLink = value
            If mIsLink Then
                GetBookmark()
            Else
                mLinkPath = ""
            End If
        End Set
    End Property
    Private mBookmark As Long = -1
    Public Property Bookmark() As Long
        Get
            Return mBookmark
        End Get
        Set(ByVal value As Long)
            mBookmark = value
        End Set
    End Property
    Public Sub GetBookmark()
        If InStr(mMediaPath, "%") <> 0 Then

            Dim s As String()
            s = mMediaPath.Split("%")
            mBookmark = Val(s(1))
        Else
            mBookmark = -1
        End If

    End Sub
    Public Function UpdateBookmark(path As String, time As String) As String
        If Right(path, 4) <> ".lnk" Then
            Return path
            Exit Function
        End If
        If InStr(path, "%") <> 0 Then
            Dim m() As String = path.Split("%")
            path = m(0) & "%" & time & "%" & m(m.Length - 1)
        Else
            path = path.Replace(".lnk", "%" & time & "%.lnk")
        End If
        mMediaPath = path
        Return path

    End Function
    Private mLinkPath As String

    Public ReadOnly Property LinkPath() As String

        Get
            Return mLinkPath
        End Get

    End Property

    Private Function FindType(file As String) As Filetype
        Try
            Dim info As New IO.FileInfo(file)
            Select Case LCase(info.Extension)
                Case ""
                    Return Filetype.Unknown
                Case ".lnk"
                    mIsLink = True
                    IsLink = True
                    Return Filetype.Link
            End Select

            Dim strExt = LCase(info.Extension)
            If InStr(VIDEOEXTENSIONS, strExt) <> 0 Then
                Return Filetype.Movie
            ElseIf InStr(PICEXTENSIONS, strExt) <> 0 Then
                Return Filetype.Pic
            ElseIf InStr(".txt.prn.sty.doc", strExt) <> 0 Then
                Return Filetype.Doc
            Else
                Return Filetype.Unknown


            End If

        Catch ex As Exception
            Return Filetype.Unknown
        End Try

    End Function
    Public Sub MediaJumpToMarker()
        If mBookmark <> -1 Then
            If StartPoint.State = StartPointHandler.StartTypes.ParticularAbsolute Then
                mPlayPosition = mBookmark
            Else
                Dim m As New StartPointHandler With {
                    .State = StartPointHandler.StartTypes.NearEnd
                }
                mPlayPosition = m.StartPoint
            End If

        Else
            mPlayPosition = StartPoint.StartPoint
        End If
        mPlayer.Ctlcontrols.currentPosition = mPlayPosition
        Debug.Print("MediaJumpMarker set position to " & mPlayPosition)
        '     mPlayer.Ctlcontrols.play()
        '  mPlayer.Refresh()
    End Sub
    Private Sub LoadMedia()

        Select Case mType
            Case Filetype.Doc

            Case Filetype.Link
                Select Case FindType(mLinkPath)
                    Case Filetype.Movie
                        If mBookmark <> -1 Then
                            If StartPoint.State = StartPointHandler.StartTypes.ParticularAbsolute Then StartPoint.Absolute = mBookmark
                        End If
                        HandleMovie(mLinkPath)
                    Case Filetype.Pic
                        HandlePic(mLinkPath)
                End Select
            Case Filetype.Movie
                HandleMovie(mMediaPath)
            Case Filetype.Pic
                HandlePic(mMediaPath)
            Case Filetype.Unknown

                Exit Sub
        End Select
        If mMediaPath <> "" Then My.Computer.Registry.CurrentUser.SetValue("File", Media.MediaPath)
    End Sub
    Public Sub HandlePic(path As String)

        Dim img As Image
        If Not Picture.Image Is Nothing Then
            DisposePic(Picture)
        End If
        img = GetImage(path)
        If img Is Nothing Then
            Exit Sub
        End If
        MainForm.OrientPic(img)
        'Resume if in middle of slideshow
        'If blnRestartSlideShowFlag Then
        '    tmrSlideShow.Enabled = True
        '    blnRestartSlideShowFlag = False
        'End If
        MainForm.MovietoPic(img)
    End Sub

    Private Sub HandleMovie(URL As String)
        Static LastURL As String
        If URL <> LastURL Then
            If mPlayer Is Nothing Then
            Else
                mPlayer.URL = URL
                LastURL = URL
            End If
        End If
        MediaJumpToMarker()
    End Sub
#Region "Event Handlers"
    Private Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles mPlayer.PlayStateChange
        Select Case e.newState

            Case WMPLib.WMPPlayState.wmppsStopped
                'MediaJumpToMarker()

            Case WMPLib.WMPPlayState.wmppsMediaEnded
                'Debug.Print("Ended:" & StartPoint.StartPoint & " " & StartPoint.Duration)
                If Not MainForm.tmrAutoTrail.Enabled And mPlayer.Visible Then
                    MainForm.AdvanceFile(True, False)
                Else
                    ' MediaJumpToMarker()
                End If
            Case WMPLib.WMPPlayState.wmppsPlaying
                'ReportTime("Playing")
                MainForm.SwitchSound(False)
                Duration = mPlayer.currentMedia.duration
                If mPaused Then
                    mPaused = False
                    Exit Sub
                End If
                If FullScreen.Changing Or Speed.Unpause Then 'Hold current position if switching to FS or back. 
                End If
                mPaused = False
            Case WMPLib.WMPPlayState.wmppsPaused ', WMPLib.WMPPlayState.wmppsTransitioning
                If Not Speed.Fullspeed Then
                    mPaused = False
                    MainForm.SwitchSound(True)
                Else
                    mPaused = True

                End If
            Case Else

        End Select
    End Sub
    Private Sub OnStartChange(sender As Object, e As EventArgs) Handles StartPoint.StartPointChanged, StartPoint.StateChanged
        MediaJumpToMarker()
        RaiseEvent StartChanged(sender, e)

    End Sub
    Private Sub UpdatePosition() Handles PositionUpdater.Tick
        mPlayPosition = mPlayer.Ctlcontrols.currentPosition
        'Duration = mPlayer.currentMedia.duration

    End Sub

#End Region
End Class