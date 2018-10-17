Imports AxWMPLib
Module MovieHandler
    Public Event MediaEnded()

    Public Property MediaMarker As Long = 0
    Public Property MediaDuration As Long ' = lngMediaDuration
    Public Property NewPosition As Long
    Public Property FromFinish As Long = 65
    Public Sub MediaJumpToMarker()

        If MediaMarker <> 0 Then
            NewPosition = MediaMarker
        Else
            NewPosition = MainForm.StartPoint.StartPoint


            Console.WriteLine(":New position is " & NewPosition & " of " & MediaDuration)
            End If

            MainForm.tmrJumpVideo.Enabled = True
        'Else
        'End If

    End Sub
    Public Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent)
        'MsgBox(e.newState)
        Static Justpaused As Boolean
        Select Case e.newState
            Case WMPLib.WMPPlayState.wmppsMediaEnded
                If Not MainForm.tmrAutoTrail.Enabled Then
                    MainForm.AdvanceFile(True, True)
                End If

            Case WMPLib.WMPPlayState.wmppsPlaying
                If Justpaused Then
                    Justpaused = False
                    Exit Sub
                End If
                MediaDuration = currentWMP.currentMedia.duration
                MainForm.StartPoint.Duration = MediaDuration
                MainForm.SwitchSound(False)

                If FullScreen.Changing Then 'Hold current position if switching to FS or back. 
                    NewPosition = currentWMP.Ctlcontrols.currentPosition
                    MainForm.tmrJumpVideo.Enabled = True
                Else

                    MediaJumpToMarker()

                End If
                '  GetAttributes(sender)
                Justpaused = False
            Case WMPLib.WMPPlayState.wmppsPaused
                '                MediaJumpToMarker()

                If MainForm.tmrSlowMo.Enabled Then
                    Justpaused = False
                    MainForm.SwitchSound(True)

                Else

                    Justpaused = True
                End If
        End Select
    End Sub

    Private Sub GetAttributes(sender As Object)
        Dim AttributeName As String = ""
        'This routine returns each of the attributes within the Media File
        For i As Integer = 0 To sender.currentMedia.attributeCount - 1

            AttributeName += sender.currentMedia.getAttributeName(i) & vbTab & sender.currentMedia.getItemInfo(sender.currentMedia.getAttributeName(i)) & vbCrLf
        Next
        MainForm.lblAttributes.Text = AttributeName
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
