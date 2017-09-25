Imports AxWMPLib
Module MovieHandler
    Public Property MediaMarker As Long = 0
    Public Property MediaDuration As Long = lngMediaDuration
    Public Property NewPosition As Long

    Public Sub MediaJumpToMarker()
        'Jumps to lMediaMarker unless not set, in which case, jumps to 65s before end.
        If MediaMarker <> 0 Then
            NewPosition = MediaMarker
        Else
            NewPosition = Math.Max(0, MediaDuration - 65)
        End If
        frmMain.tmrJumpVideo.Enabled = True
    End Sub
    Public Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent)
        Select Case e.newState
            Case WMPLib.WMPPlayState.wmppsMediaEnded
                Dim KeyEvent As New KeyEventArgs(KeyNextFile)
                frmMain.HandleKeys(sender, KeyEvent)
            Case WMPLib.WMPPlayState.wmppsPlaying
                If blnRandomStartPoint Then frmMain.JumpRandom(False)
        End Select
    End Sub
End Module
