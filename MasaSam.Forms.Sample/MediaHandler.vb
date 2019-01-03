Public Class MediaHandler

    Public Event MediaFinished(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaClosed(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaChanged(ByVal sender As Object, ByVal e As EventArgs)
    Private DefaultFile As String = "C:\exiftools.exe"

    Private mType As Filetype
    Public Property MediaType() As Filetype
        Get
            If mMediaPath = "" Then
            Else

                mType = FindType(mMediaPath)
            End If
            If mType = Filetype.Link Then
                IsLink = True
            Else
                IsLink = False
            End If
            Return mType
        End Get
        Set(ByVal value As Filetype)
            mType = value
        End Set
    End Property

    Private mPicBox As PictureBox
    Public Property Picture() As PictureBox
        Get
            Return mPicBox
        End Get
        Set(ByVal value As PictureBox)
            mPicBox = value
        End Set
    End Property

    Private mPlayer As New AxWMPLib.AxWindowsMediaPlayer
    Public Property Player() As AxWMPLib.AxWindowsMediaPlayer
        Get
            Return mPlayer
        End Get
        Set(ByVal value As AxWMPLib.AxWindowsMediaPlayer)
            mPlayer = value
        End Set
    End Property
    Private mPlayPosition As Long
    Public Property Position() As Long
        Get
            '  mPlayPosition = mPlayer.Ctlcontrols.currentPosition
            Return mPlayPosition
        End Get
        Set(ByVal value As Long)
            mPlayPosition = value
            '    mPlayer.Ctlcontrols.currentPosition = mPlayPosition
        End Set
    End Property

    Private mDuration As Long
    Public ReadOnly Property Duration() As Long
        Get
            mDuration = mPlayer.currentMedia.duration
            Return mDuration
        End Get
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
                    mMediaDirectory = f.Directory.FullName
                Else
                    mMediaPath = DefaultFile
                    mMediaDirectory = New IO.FileInfo(mMediaPath).Directory.FullName
                End If

                RaiseEvent MediaChanged(Me, New EventArgs)
            End If


        End Set
    End Property
    Private Function Findfile(f As IO.FileInfo, dir As IO.DirectoryInfo) As IO.FileInfo

        'If f doesn't exist
        'Split the path up into bits
        'Progressively check the existence of each path
        'When a part doesn't exist, choose the first file in that directory
        While Not f.Exists
            Dim directoryNames As New List(Of String)(f.FullName.Split(System.IO.Path.DirectorySeparatorChar))
            Dim path As String = ""
            For i = 0 To directoryNames.Count - 1
                If i = 0 Then
                    path = directoryNames(i) & IO.Path.DirectorySeparatorChar
                Else
                    path = path & directoryNames(i) & IO.Path.DirectorySeparatorChar
                End If
                Dim subdir As New IO.DirectoryInfo(path)
                If subdir.Exists Then
                    Findfile(f, subdir)
                Else
                    f = Nothing
                End If
            Next
            Try
                If Not f.Exists Then
                    Try
                        f = dir.GetFiles.First
                    Catch ex As IO.DirectoryNotFoundException
                        f = Nothing
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As System.NullReferenceException
                f = Nothing
            End Try
        End While
        Return f
    End Function

    Public Sub New()

    End Sub
    Private mMediaDirectory As String
    Public Property MediaDirectory() As String
        Get
            'Dim f As New IO.FileInfo(mMediaPath)
            ' mMediaDirectory = f.Directory.FullName
            Return mMediaDirectory
        End Get
        Set(ByVal value As String)
            If value <> mMediaDirectory Then
                mMediaDirectory = value
                RaiseEvent MediaChanged(Me, New EventArgs)
            End If
        End Set
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

                    'mMediaPath = LinkTarget(info.FullName) ' CreateObject("WScript.Shell").CreateShortcut(info.FullName).TargetPath
                    'mLinkPath = info.FullName
                    'MediaDirectory = info.Directory.FullName
                    'MainForm.Text = "Metavisua - " & Media.MediaPath

                    'Try
                    'If My.Computer.FileSystem.FileExists(mMediaPath) Then
                    'info = New IO.FileInfo(mMediaPath)
                    'Else
                    'Return Filetype.Unknown
                    'Exit Function
                    'End If

                    'Catch ex As Exception
                    'End Try
                    Return Filetype.Link
                    '           Exit Function
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

        Catch ex As IO.PathTooLongException
            Return Filetype.Unknown
        End Try

    End Function

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
    Private mBookmark As Long
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
            mBookmark = 0
        End If

    End Sub
    Private mLinkPath As String
    Public ReadOnly Property LinkPath() As String
        Get
            Return mLinkPath
        End Get

    End Property
End Class
