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
    Public Event LoadedMedia(ByRef MH As MediaHandler)

    Public Event MediaShown(ByRef MH As MediaHandler)

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
        Static oldindex As Integer
        NextF.Forwards = (mListIndex > oldindex) Or (mListIndex = 0)
        NextF.CurrentIndex = index
        Current = NextF.CurrentItem
        Nxt = NextF.NextItem
        Prev = NextF.PreviousItem

        Console.WriteLine()
        Console.WriteLine("Previous:" & Prev)
        Console.WriteLine("Current:" & Current)
        Console.WriteLine("Next:" & Nxt)
        Select Case Current
            Case mMedia1.MediaPath
                RotateMedia(mMedia1, mMedia2, mMedia3, Nxt, Prev)
            Case mMedia2.MediaPath
                RotateMedia(mMedia2, mMedia3, mMedia1, Nxt, Prev)
            Case mMedia3.MediaPath
                RotateMedia(mMedia3, mMedia1, mMedia2, Nxt, Prev)
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
    Private Sub RotateMedia(ByRef ThisMH As MediaHandler, ByRef NextMH As MediaHandler, ByRef PrevMH As MediaHandler, nxt As String, prev As String)

        NextMH.MediaPath = nxt
        NextMH.Player.Visible = False
        PrevMH.MediaPath = prev
        PrevMH.Player.Visible = False
        RaiseEvent LoadedMedia(NextMH)
        RaiseEvent LoadedMedia(PrevMH)
        If ThisMH.MediaType = Filetype.Movie Then
            ShowPlayer(ThisMH)
        ElseIf ThisMH.MediaType = Filetype.Pic Then
            ShowPicture(ThisMH)
        End If

    End Sub

    Private Sub ShowPlayer(ByRef MHX As MediaHandler)
        Media1.settings.mute = True
        Media2.settings.mute = True
        Media3.settings.mute = True
        With MHX.Player
            .Visible = True
            .BringToFront()
            .settings.mute = False
        End With
        RaiseEvent MediaShown(MHX)

    End Sub

    Private Sub ShowPicture(ByRef MHX As MediaHandler)
        MHX.Picture.BringToFront()
        RaiseEvent MediaShown(MHX)

    End Sub
End Class
