Imports AxWMPLib

Public Class FullScreen
    Public Shared Property Changing As Boolean
    Public Event FullScreenClosing()

    Private Sub FullScreen_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        InitialisePlayer(FSWMP)
        InitialisePlayer(FSWMP3)
        InitialisePlayer(FSWMP2)
        MSFiles.AssignPlayers(FSWMP, FSWMP2, FSWMP3)

    End Sub
    Private Sub InitialisePlayer(WMP As AxWindowsMediaPlayer)
        WMP.uiMode = "None"
        WMP.Dock = DockStyle.Fill

    End Sub

    Private Sub FullScreen_Keydown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control

        If e.KeyCode = Keys.Escape Then
            MainForm.GoFullScreen(False)
            ' RaiseEvent FullScreenClosing()
        End If

        MainForm.HandleKeys(Me, e)
        e.SuppressKeyPress = True
    End Sub

    Public Sub FSWMP_PlayStateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles FSWMP.PlayStateChange
        'PlaystateChange(sender, e)

    End Sub


    Private Sub FSWMP_KeyDownEvent(sender As Object, e As _WMPOCXEvents_KeyDownEvent) Handles FSWMP.KeyDownEvent
        If e.nKeyCode = Keys.Escape Then
            MainForm.GoFullScreen(False)
        End If
    End Sub



    Private Sub fullScreenPicBox_MouseWheel(sender As Object, e As MouseEventArgs) Handles fullScreenPicBox.MouseWheel, Me.MouseWheel
        PictureFunctions.Mousewheel(fullScreenPicBox, sender, e)
    End Sub

    Private Sub fullScreenPicBox_MouseMove(sender As Object, e As MouseEventArgs) Handles fullScreenPicBox.MouseMove, Me.MouseMove
        picBlanker = FSBlanker
        PictureFunctions.MouseMove(fullScreenPicBox, sender, e)
    End Sub

    Private Sub fullScreenPicBox_KeyDown(sender As Object, e As KeyEventArgs) Handles fullScreenPicBox.KeyDown, Me.KeyUp
        ShiftDown = e.Shift
        CtrlDown = e.Control
    End Sub


    Private Sub FullScreen_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub fullScreenPicBox_MouseDown(sender As Object, e As MouseEventArgs) Handles fullScreenPicBox.MouseDown
        Select Case e.Button
            Case MouseButtons.XButton1, MouseButtons.XButton2
                MainForm.AdvanceFile(e.Button = MouseButtons.XButton2)
                e = Nothing

            Case Else
                PicClick(fullScreenPicBox)
        End Select
    End Sub


End Class