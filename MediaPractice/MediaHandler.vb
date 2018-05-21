Imports AxWMPLib

Public Class MediaHandler
    Inherits AxWindowsMediaPlayer
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

    Sub OnMediaReady()
        mDuration = MyBase.currentMedia.duration
    End Sub
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



    Public Sub SetPlaybackSpeed()

    End Sub

    Private Sub MediaHandler_PlayStateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles Me.PlayStateChange

    End Sub
End Class
