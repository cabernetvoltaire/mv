﻿Imports AxWMPLib
Public Class MovieSwapper
    Private mNext As New NextFile
    Private mMedia1 As New MediaHandler
    Private mMedia2 As New MediaHandler
    Public SH As New StartPointHandler
    Public WithEvents Media1 As New AxWindowsMediaPlayer
    Public WithEvents Media2 As New AxWindowsMediaPlayer
    Private mFileList As New List(Of String)
    Private mListIndex As Integer
    Private mListbox As New ListBox
    Public Event LoadedMedia(Media As AxWindowsMediaPlayer, MH As MediaHandler)

    Public Event MediaShown(Media As AxWindowsMediaPlayer, MH As MediaHandler)

    Public Property Listbox() As ListBox
        Get
            Return mListbox
        End Get
        Set(ByVal value As ListBox)
            mListbox = value
            mNext.Listbox = value
            For Each m In mListbox.Items
                mFileList.Add(m)
            Next
        End Set
    End Property

    Public Property ListIndex() As Integer
        Get
            Return mListIndex
        End Get
        Set(ByVal value As Integer)
            mListIndex = value
            mNext.CurrentIndex = value
            SetIndex(value)
        End Set
    End Property
    Public Sub New(MP1 As AxWindowsMediaPlayer, MP2 As AxWindowsMediaPlayer)
        '  Listbox = Lbox
        ' mNext.Listbox = Lbox
        Media1 = MP1
        Media2 = MP2

        mMedia1.Player = MP1
        mMedia2.Player = MP2
    End Sub

    'Private Sub PSChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) ' Handles Media1.PlayStateChange, Media2.PlayStateChange 'Swapper

    '    If sender Is Media1 Then
    '        MovieHandler.PlaystateChangeNew(sender, e, SH, mMedia1)
    '        If e.newState = WMPLib.WMPPlayState.wmppsPlaying Then
    '            RaiseEvent LoadedMedia(Media1, mMedia1)
    '            RaiseEvent MediaShown(Media2, mMedia2)

    '        End If

    '    Else

    '        MovieHandler.PlaystateChangeNew(sender, e, SH, mMedia2)
    '        If e.newState = WMPLib.WMPPlayState.wmppsPlaying Then
    '            RaiseEvent LoadedMedia(Media2, mMedia2)
    '            RaiseEvent MediaShown(Media1, mMedia1)

    '        End If

    '    End If
    'End Sub
    Private Sub SetIndex(index As Integer)
        Dim Current As String
        Dim Nxt As String
        Static oldindex As Integer
        If ListIndex < oldindex Then
            mNext.Forwards = False
        Else
            mNext.Forwards = True
        End If
        mNext.CurrentIndex = index
        Current = mNext.CurrentItem
        Nxt = mNext.NextItem
        Select Case Current
            Case mMedia1.MediaPath
                SwitchPlayers(Media2, Media1)
                mMedia2.MediaPath = Nxt
                MainForm.Media = mMedia1
                'Media2.URL = Nxt
           '     RaiseEvent MediaShown(Media1, mMedia1)
            Case mMedia2.MediaPath
                SwitchPlayers(Media1, Media2)
                mMedia1.MediaPath = Nxt
                MainForm.Media = mMedia2
                'Media1.URL = Nxt

            Case Else
                'Media2.URL = Nxt
                'Media1.URL = Current
                SwitchPlayers(Media2, Media1)
                mMedia1.MediaPath = Current
                mMedia2.MediaPath = Nxt
                MainForm.Media = mMedia1

        End Select



        oldindex = index
    End Sub
    Private Sub SwitchPlayers(OldWMP As AxWindowsMediaPlayer, NewWMP As AxWindowsMediaPlayer)
        NewWMP.Visible = True
        OldWMP.Visible = False
        NewWMP.BringToFront()
        NewWMP.settings.mute = False
        MainForm.currentWMP = NewWMP
        OldWMP.settings.mute = True
    End Sub

End Class