Imports AxWMPLib
Module MovieHandler
    Public Event MediaEnded()

    Public Property MediaMarker As Long = 0
    Public Property MediaDuration As Long = lngMediaDuration
    Public Property NewPosition As Long
    Public Property FromFinish As Long = 65
    Public Property FrameRate As Int16
    Public Sub MediaJumpToMarker()
        'Jumps to lMediaMarker unless not set, in which case, jumps to 65s before end.
        If MediaMarker <> 0 Then
            NewPosition = MediaMarker
        Else
            Select Case frmMain.StartType
                Case StartTypes.NearBeginning
                    NewPosition = FromFinish
                Case StartTypes.NearEnd
                    NewPosition = Math.Max(0, MediaDuration - FromFinish)

                Case StartTypes.Particular
                    NewPosition = MediaDuration * frmMain.StartPointPercentage / 100
            End Select
        End If

        frmMain.tmrJumpVideo.Enabled = True
    End Sub
    Public Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent)
        'MsgBox(e.newState)
        Select Case e.newState
            Case WMPLib.WMPPlayState.wmppsMediaEnded
                If Not frmMain.tmrAutoTrail.Enabled Then
                    frmMain.AdvanceFile(True, True)
                End If

            Case WMPLib.WMPPlayState.wmppsPlaying

                MediaDuration = currentWMP.currentMedia.duration
                If blnJumpToMark Then
                    MediaJumpToMarker(False)

                ElseIf blnRandomStartPoint Then
                    If FullScreen.Changing Then
                    Else
                        frmMain.JumpRandom(False)

                    End If
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
