Imports AxWMPLib
Module MovieHandler
    Public Event MediaEnded()

    Public Property MediaMarker As Long = 0
    Public Property MediaDuration As Long ' = lngMediaDuration
    Public Property NewPosition As Long
    Public Property FromFinish As Long = 65
    Public Sub MediaJumpToMarker()
        Static LastTriggered As DateTime
        '  If (Now() - LastTriggered).TotalMilliseconds > 200 Then

        'Jumps to lMediaMarker unless not set, in which case, jumps to current startpoint
        If MediaMarker <> 0 Then
                NewPosition = MediaMarker
            Else
            NewPosition = frmMain.StartPoint.StartPoint


            Console.WriteLine(":New position is " & NewPosition & " of " & MediaDuration)
            End If

            frmMain.tmrJumpVideo.Enabled = True
        'Else
        'End If
        LastTriggered = Now

    End Sub
    Public Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent)
        'MsgBox(e.newState)
        Static Justpaused As Boolean
        Select Case e.newState
            Case WMPLib.WMPPlayState.wmppsMediaEnded
                If Not frmMain.tmrAutoTrail.Enabled Then
                    frmMain.AdvanceFile(True, True)
                End If

            Case WMPLib.WMPPlayState.wmppsPlaying
                If Justpaused Then
                    Justpaused = False
                    Exit Sub
                End If
                MediaDuration = currentWMP.currentMedia.duration
                frmMain.StartPoint.Duration = MediaDuration
                If FullScreen.Changing Then
                    NewPosition = currentWMP.Ctlcontrols.currentPosition
                    frmMain.tmrJumpVideo.Enabled = True
                Else
                    MediaJumpToMarker()

                End If
                Justpaused = False
            Case WMPLib.WMPPlayState.wmppsPaused
                If frmMain.tmrSlowMo.Enabled Then
                    Justpaused = False
                Else

                    Justpaused = True
                End If
        End Select
    End Sub

    Private Sub GetAttributes(sender As Object)
        Dim AttributeName As String = ""
        'This routine returns each of the attributes within the Media File
        For i As Integer = 0 To sender.currentMedia.attributeCount - 1
            AttributeName += sender.currentMedia.getAttributeName(i) & vbCrLf
        Next
        MsgBox(AttributeName)
        'The Attributes returned are:
        ' Duration
        ' FileType
        ' Is_Trusted
        ' MediaType
        ' SourceURL
        ' Streams
        ' Title
        ' WMServerVersion

        AttributeName = ""
    End Sub
End Module
