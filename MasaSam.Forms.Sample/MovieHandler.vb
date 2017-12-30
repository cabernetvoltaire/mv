Imports AxWMPLib
Module MovieHandler
    Public Property MediaMarker As Long = 0
    Public Property MediaDuration As Long = lngMediaDuration
    Public Property NewPosition As Long
    Public Property FromFinish As Long = 65
    Public Sub MediaJumpToMarker()
        'Jumps to lMediaMarker unless not set, in which case, jumps to 65s before end.
        If MediaMarker <> 0 Then
            NewPosition = MediaMarker
        Else
            NewPosition = Math.Max(0, MediaDuration - FromFinish)
        End If
        frmMain.tmrJumpVideo.Enabled = True
    End Sub
    Public Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent)
        'MsgBox(e.newState)
        Select Case e.newState
            Case WMPLib.WMPPlayState.wmppsMediaEnded
                If Not frmMain.tmrAutoTrail.Enabled Then
                    Dim KeyEvent As New KeyEventArgs(KeyNextFile)
                    frmMain.HandleKeys(sender, KeyEvent)
                End If

            Case WMPLib.WMPPlayState.wmppsPlaying
                MediaDuration = currentWMP.currentMedia.duration
                If blnJumpToMark Then
                    MediaJumpToMarker()

                ElseIf blnRandomStartPoint Then
                If FullScreen.Changing Then
                Else
                    frmMain.JumpRandom(False)

                    End If
                End If
        End Select
    End Sub
End Module
