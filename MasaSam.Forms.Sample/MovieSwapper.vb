Imports AxWMPLib
Public Class MediaSwapper
    Public NextF As New NextFile
    Private mMedia1 As New MediaHandler("mMedia1")
    Private mMedia2 As New MediaHandler("mMedia2")
    Private mMedia3 As New MediaHandler("mMedia3")

    Private mFileList As New List(Of String) '
    Private mListIndex As Integer
    Private mListbox As New ListBox
    Private mListcount As Integer
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
            If value < 0 Then Exit Property
            mListIndex = value

            NextF.CurrentIndex = value
            SetIndex(value)
        End Set
    End Property
    Public Sub New(ByRef MP1 As AxWindowsMediaPlayer, ByRef MP2 As AxWindowsMediaPlayer, ByRef MP3 As AxWindowsMediaPlayer)
        AssignPlayers(MP1, MP2, MP3)

    End Sub
    Public Sub AssignPlayers(ByRef MP1 As AxWindowsMediaPlayer, ByRef MP2 As AxWindowsMediaPlayer, ByRef MP3 As AxWindowsMediaPlayer)
        mMedia1.Player = MP1
        mMedia2.Player = MP2
        mMedia3.Player = MP3

    End Sub
    Private Function FindMH(path As String) As MediaHandler
        If mMedia1.MediaPath = path Then
            Return mMedia1
        ElseIf mMedia2.MediaPath = path Then
            Return mMedia2
        ElseIf mMedia3.MediaPath = path Then
            Return mMedia3
        Else
            Return Nothing
        End If
    End Function
    Private Function FreeMH(c As String, p As String, n As String) As MediaHandler
        If mMedia1.MediaPath <> c And mMedia1.MediaPath <> p And mMedia1.MediaPath <> n Then
            Return mMedia1
        ElseIf mMedia2.MediaPath <> c And mMedia2.MediaPath <> p And mMedia2.MediaPath <> n Then
            Return mMedia2
        ElseIf mMedia3.MediaPath <> c And mMedia3.MediaPath <> p And mMedia3.MediaPath <> n Then
            Return mMedia3
        Else
            Return Nothing
        End If
    End Function
    Private Sub SetIndex(index As Integer)
        Dim Current As String
        Dim Nxt As String
        Dim Prev As String
        mListcount = Listbox.Items.Count
        Static oldindex As Integer

        NextF.Forwards = (mListIndex > oldindex) Or (mListIndex = 0)
        NextF.CurrentIndex = index
        Current = NextF.CurrentItem
        Nxt = NextF.NextItem
        Prev = NextF.PreviousItem

        Select Case Current
            Case mMedia1.MediaPath
                RotateMedia(mMedia1, mMedia2, mMedia3)
            Case mMedia2.MediaPath
                RotateMedia(mMedia2, mMedia3, mMedia1)
            Case mMedia3.MediaPath
                RotateMedia(mMedia3, mMedia1, mMedia2)
            Case Else
                mMedia1.MediaPath = Current
                mMedia2.MediaPath = Nxt
                mMedia3.MediaPath = Prev
                RotateMedia(mMedia1, mMedia2, mMedia3)
        End Select
        oldindex = index
    End Sub

    Private Sub Prepare(ByRef MH As MediaHandler, path As String)
        Debug.Print("PREPARE: " & MH.Player.Name)
        MH.MediaPath = path
        MH.Player.Visible = True
        '        MH.StartPoint.State = Media.StartPoint.State
        '  MH.Pause(True)
    End Sub
    Private Sub RotateMedia(ByRef ThisMH As MediaHandler, ByRef NextMH As MediaHandler, ByRef PrevMH As MediaHandler)

        Prepare(PrevMH, NextF.PreviousItem)
        Prepare(NextMH, NextF.NextItem)
        Prepare(ThisMH, NextF.CurrentItem)

        If ThisMH.MediaType = Filetype.Movie Then
            ShowPlayer(ThisMH)

        ElseIf ThisMH.MediaType = Filetype.Pic Then
            ShowPicture(ThisMH)
        End If

    End Sub
    Public Sub SetStartStates(ByRef SH As StartPointHandler)
        mMedia1.StartPoint.State = SH.State
        mMedia2.StartPoint.State = SH.State
        mMedia3.StartPoint.State = SH.State


    End Sub
    Public Sub SetStartpoints(ByRef SH As StartPointHandler)
        '  Dim dur As Long

        'dur = mMedia1.StartPoint.Duration
        'dur = mMedia3.StartPoint.Duration
        'dur = mMedia2.StartPoint.Duration
        mMedia1.StartPoint = SH

        mMedia2.StartPoint = SH

        mMedia3.StartPoint = SH
        '  SetIndex(ListIndex)
        'mMedia2.StartPoint.Duration = dur
        ' mMedia1.StartPoint.Duration = dur
        'mMedia3.StartPoint.Duration = dur

    End Sub



    Private Sub MuteAll()
        mMedia1.Player.settings.mute = True
        mMedia2.Player.settings.mute = True
        mMedia3.Player.settings.mute = True

    End Sub
    Private Sub ShowPlayer(ByRef MHX As MediaHandler)
        '  MHX.MediaJumpToMarker()
        Debug.Print("SHOWPLAYER" & MHX.Player.Name)
        MuteAll()
        'MHX.Pause(False)
        Debug.Print(MHX.Player.URL & " unpaused")
        With MHX.Player

            '.Visible = True
            .BringToFront()
            .settings.mute = False
        End With
        RaiseEvent MediaShown(MHX)

    End Sub

    Private Sub ShowPicture(ByRef MHX As MediaHandler)
        MuteAll()
        MHX.Picture.BringToFront()
        RaiseEvent MediaShown(MHX)

    End Sub
End Class
