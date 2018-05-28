Public Class MediaHandler

    Public Event MediaFinished(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaClosed(ByVal sender As Object, ByVal e As EventArgs)
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
            mPlayer.Ctlcontrols.currentPosition = mPlayPosition
        End Set
    End Property

    Private mDuration As Long
    Public ReadOnly Property Duration() As Long
        Get
            mDuration = mPlayer.currentMedia.duration
            Return mDuration
        End Get
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
            mMediaPath = value
        End Set
    End Property
    Public Sub New()

    End Sub
    Public Sub SetPlaybackSpeed()

    End Sub

    Private Function FindType(file As String) As Filetype
        Dim IsLink As Boolean = False
        Try
            Dim info As New IO.FileInfo(file)
            Select Case LCase(info.Extension)
                Case ""
                    Return Filetype.Unknown
                Case ".lnk"
                    IsLink = True
                    mMediaPath = LinkTarget(info.FullName) ' CreateObject("WScript.Shell").CreateShortcut(info.FullName).TargetPath
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

            strExt = LCase(info.Extension)
            If InStr(strVideoExtensions, strExt) <> 0 Then
                Return Filetype.Movie
            ElseIf InStr(strPicExtensions, strExt) <> 0 Then
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


End Class
