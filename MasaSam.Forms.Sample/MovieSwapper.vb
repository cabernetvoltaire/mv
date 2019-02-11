Imports AxWMPLib
Public Class MovieSwapper
    Public NextF As New NextFile
    Private mMedia1 As New MediaHandler
    Private mMedia2 As New MediaHandler
    Private mMedia3 As New MediaHandler
    'Public WithEvents Medias As New AxWindowsMediaPlayer()
    '  Public SH As New StartPointHandler
    Public WithEvents Media1 As New AxWindowsMediaPlayer
    Public WithEvents Media2 As New AxWindowsMediaPlayer
    Public WithEvents Media3 As New AxWindowsMediaPlayer

    Private mFileList As New List(Of String)
    Private mListIndex As Integer
    Private mListbox As New ListBox
    Public Event LoadedMedia(MH As MediaHandler)

    Public Event MediaShown(MH As MediaHandler)

    Public Property Listbox() As ListBox
        Get
            Return mListbox
        End Get
        Set(ByVal value As ListBox)
            mListbox = value
            NextF.Listbox = value
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
            NextF.CurrentIndex = value
            SetIndex(value)
        End Set
    End Property
    Public Sub New(MP1 As AxWindowsMediaPlayer, MP2 As AxWindowsMediaPlayer, MP3 As AxWindowsMediaPlayer)
        '  Listbox = Lbox
        ' mNext.Listbox = Lbox
        Media1 = MP1
        Media2 = MP2
        Media3 = MP3
        '  Medias = (MP1, MP2, MP3)

        mMedia1.Player = MP1
        mMedia2.Player = MP2
        mMedia3.Player = MP3

    End Sub

    Private Sub SetIndex(index As Integer)
        Dim Current As String
        Dim Nxt As String
        Dim Prev As String
        Static oldindex As Integer = -1
        NextF.Forwards = (mListIndex > oldindex)
        NextF.CurrentIndex = index
        Current = NextF.CurrentItem
        Nxt = NextF.NextItem
        Prev = NextF.PreviousItem
        Console.WriteLine("Previous:" & Prev)
        Console.WriteLine("Current:" & Current)
        Console.WriteLine("Next:" & Nxt)

        Select Case Current
            Case mMedia1.MediaPath
                mMedia2.MediaPath = Nxt
                mMedia3.MediaPath = Prev
                RaiseEvent LoadedMedia(mMedia2)
                RaiseEvent LoadedMedia(mMedia3)
                ShowPlayer(mMedia1)

            Case mMedia2.MediaPath
                mMedia3.MediaPath = Nxt
                mMedia1.MediaPath = Prev
                RaiseEvent LoadedMedia(mMedia1)
                RaiseEvent LoadedMedia(mMedia3)
                ShowPlayer(mMedia2)
            Case mMedia3.MediaPath
                mMedia1.MediaPath = Nxt
                mMedia2.MediaPath = Prev
                RaiseEvent LoadedMedia(mMedia1)
                RaiseEvent LoadedMedia(mMedia2)
                ShowPlayer(mMedia3)


            Case Else
                mMedia1.MediaPath = Current
                mMedia2.MediaPath = Nxt
                mMedia3.MediaPath = Prev
                RaiseEvent LoadedMedia(mMedia2)
                RaiseEvent LoadedMedia(mMedia1)
                RaiseEvent LoadedMedia(mMedia3)
                ShowPlayer(mMedia2)
        End Select

        oldindex = index
    End Sub


    Private Sub ShowPlayer(ByRef NewMediaHandler As MediaHandler)
        Media1.settings.mute = True
        Media2.settings.mute = True
        Media3.settings.mute = True
        With NewMediaHandler.Player
            .Visible = True
            .BringToFront()
            .settings.mute = False
        End With
        RaiseEvent MediaShown(NewMediaHandler)

    End Sub

End Class
