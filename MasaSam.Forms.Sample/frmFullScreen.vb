Imports AxWMPLib

Public Class FullScreen


    Private Sub FullScreen_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'FSWMP.stretchToFit = True 'TODo improve this
        FSWMP.uiMode = "None"
        FSWMP.Location = Me.Location
        Timer1.Interval = 25
        Timer1.Enabled = True



    End Sub

    Private Sub FullScreen_Keydown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control
        If e.KeyCode = Keys.Escape Then
            frmMain.GoFullScreen(False)
        End If

        frmMain.HandleKeys(Me, e)
        e.Handled = True
    End Sub

    Public Sub FSWMP_PlayStateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles FSWMP.PlayStateChange
        PlaystateChange(sender, e)
    End Sub



    'Private Sub FSWMP_KeyDownevent(sender As Object, e As KeyEventArgs) Handles FSWMP.KeyDown
    '    If e.KeyCode = Keys.Escape Then
    '        frmMain.GoFullScreen(False)
    '    End If

    '    frmMain.HandleKeys(Me, e)
    '    e.Handled = True
    'End Sub

    Private Sub FSWMP_KeyDownEvent(sender As Object, e As _WMPOCXEvents_KeyDownEvent) Handles FSWMP.KeyDownEvent
        If e.nKeyCode = Keys.Escape Then
            frmMain.GoFullScreen(False)
        End If
    End Sub

    Private Sub fullScreenPicBox_Click(sender As Object, e As EventArgs) Handles fullScreenPicBox.Click
        picBlanker = FSBlanker
        PicClick(fullScreenPicBox)
    End Sub

    Private Sub fullScreenPicBox_MouseWheel(sender As Object, e As MouseEventArgs) Handles fullScreenPicBox.MouseWheel, Me.MouseWheel
        PictureFunctions.Mousewheel(fullScreenPicBox, sender, e)
    End Sub

    Private Sub fullScreenPicBox_MouseMove(sender As Object, e As MouseEventArgs) Handles fullScreenPicBox.MouseMove, Me.MouseMove
        picBlanker = FSBlanker
        PictureFunctions.MouseMove(fullScreenPicBox, sender, e)
    End Sub

    Private Sub fullScreenPicBox_KeyDown(sender As Object, e As KeyEventArgs) Handles fullScreenPicBox.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control
    End Sub

    Private Sub FullScreen_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        ShiftDown = e.Shift
        CtrlDown = e.Control
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If iScreenstate = Screenstate.TrueSize Then
            ZoomPicture(fullScreenPicBox, True, Autozoomrate)
        End If
    End Sub

    Private Sub FSBlanker_Click(sender As Object, e As EventArgs) Handles FSBlanker.Click

    End Sub
End Class