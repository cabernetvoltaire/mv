Imports AxWMPLib
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
    Public Event LoadedMedia(Media As AxWindowsMediaPlayer)
    Public Event MediaShown(Media As AxWindowsMediaPlayer)

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

    Private Sub PSChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles Media1.PlayStateChange, Media2.PlayStateChange 'Swapper
        'Static WMP As AxWindowsMediaPlayer
        'WMP = currentWMP
        MainForm.currentWMP = sender
        If sender Is Media1 Then
            ' Media = mMedia1
            MovieHandler.PlaystateChangeNew(sender, e, SH, mMedia1)
            RaiseEvent LoadedMedia(Media1)
        Else
            'Media = mMedia2

            MovieHandler.PlaystateChangeNew(sender, e, SH, mMedia2)
            RaiseEvent LoadedMedia(Media2)

        End If
        'currentWMP = WMP
    End Sub
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
        If Media1.URL = "" Then
            Media1.URL = Current
            Media1.BringToFront()
            Media2.URL = Nxt

        ElseIf Media1.URL = Current Then
            Media2.URL = Nxt
            mMedia2.MediaPath = Nxt
            Mysettings.Media = mMedia1
            SwitchPlayers(Media2, Media1)

        Else
            Media1.URL = Nxt
            mMedia1.MediaPath = Nxt
            Mysettings.Media = mMedia1
            '            MainForm.LoadMedia(Me, Nothing)
            SwitchPlayers(Media1, Media2)


        End If
        oldindex = index
    End Sub
    Private Sub SwitchPlayers(OldWMP As AxWindowsMediaPlayer, NewWMP As AxWindowsMediaPlayer)
        NewWMP.Visible = True
        NewWMP.BringToFront()
        'NewWMP.Ctlcontrols.play()
        NewWMP.settings.mute = False
        OldWMP.Visible = False
        OldWMP.settings.mute = True
        '        MainForm.currentWMP = NewWMP
        If NewWMP Is Media1 Then
            Mysettings.Media = mMedia1
        Else
            Mysettings.Media = mMedia2

        End If
        ' MainForm.tmrPicLoad.Enabled = True
        RaiseEvent MediaShown(NewWMP)
    End Sub

End Class
