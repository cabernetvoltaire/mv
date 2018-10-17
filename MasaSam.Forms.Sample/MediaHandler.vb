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
            mPlayPosition = mPlayer.Ctlcontrols.currentPosition
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
            If value <> "" And value <> mMediaPath Then
                Dim f As New IO.FileInfo(value)
                If f.Exists Then
                    mMediaPath = value
                Else
                    mMediaPath = DefaultFile
                End If
                If b <> mMediaPath Then
                    mMediaDirectory = f.Directory.FullName
                    RaiseEvent MediaChanged(Me, New EventArgs)
                End If
            Else
                mMediaPath = DefaultFile

            End If
        End Set
    End Property

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
                    mMediaPath = LinkTarget(info.FullName) ' CreateObject("WScript.Shell").CreateShortcut(info.FullName).TargetPath
                    mLinkPath = info.FullName
                    Try
                        If My.Computer.FileSystem.FileExists(mMediaPath) Then
                            info = New IO.FileInfo(mMediaPath)
                        Else
                            Return Filetype.Unknown
                            Exit Function
                        End If

                    Catch ex As Exception
                    End Try
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
        Catch ex As IO.PathTooLongException
            Return Filetype.Unknown
        End Try
    End Function

    Private mIsLink As Boolean
    Public Property IsLink() As Boolean
        Get
            Return mIsLink
        End Get
        Set(ByVal value As Boolean)
            mIsLink = value
            If mIsLink Then
            Else
                mLinkPath = ""
            End If
        End Set
    End Property

    Private mLinkPath As String
    Public Property LinkPath() As String
        Get
            Return mLinkPath
        End Get
        Set(ByVal value As String)
            mLinkPath = value
            If mLinkPath <> "" Then mIsLink = True
        End Set
    End Property
End Class
