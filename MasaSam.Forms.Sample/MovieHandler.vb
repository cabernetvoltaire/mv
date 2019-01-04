Imports AxWMPLib
Module MovieHandler
    Public Event MediaEnded()

    Public Property MediaMarker As Long = 0
    Public Property MediaDuration As Long ' = lngMediaDuration
    Public Property NewPosition As Long
    Public Property FromFinish As Long = 65
    Public Sub MediaJumpToMarker()
        'ReportTime("Jumpto Marker")
        If Media.MediaPath = "" Then Exit Sub
        If MediaMarker <> 0 Then
            NewPosition = MediaMarker
            ' Media.Bookmark = MediaMarker
        Else
            NewPosition = MainForm.StartPoint.StartPoint
            'ReportTime("Jump to Marker " & NewPosition / Media.Duration)


            Console.WriteLine("New position is " & NewPosition & " of ")
        End If

        MainForm.tmrJumpVideo.Enabled = True
        'Else
        'End If
    End Sub
    Public Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent)
        Dim wmp As AxWindowsMediaPlayer = CType(sender, AxWindowsMediaPlayer)


        'ReportTime("Playstate " & e.newState)

        Static Justpaused As Boolean
        'MsgBox(e.newState.ToString)
        Select Case e.newState
            Case WMPLib.WMPPlayState.wmppsMediaEnded
                If Not MainForm.tmrAutoTrail.Enabled Then
                    MainForm.AdvanceFile(True, False)
                End If

            Case WMPLib.WMPPlayState.wmppsPlaying
                'ReportTime("Playing")
                Media.Duration = currentWMP.currentMedia.duration
                MainForm.StartPoint.Duration = Media.Duration
                MainForm.SwitchSound(False)
                wmp.Visible = True
                If Justpaused Then
                    Justpaused = False
                    Exit Sub
                End If
                '                MediaDuration = currentWMP.currentMedia.duration
                If FullScreen.Changing Or MainForm.SP.Unpause Then 'Hold current position if switching to FS or back. 
                    NewPosition = currentWMP.Ctlcontrols.currentPosition
                    MainForm.tmrJumpVideo.Enabled = True
                Else
                    ' ReportFault("MHPSCHange", "Just before MJ2M", True)
                    MediaJumpToMarker()
                End If
                '  GetAttributes(sender)
                Justpaused = False
            Case WMPLib.WMPPlayState.wmppsPaused ', WMPLib.WMPPlayState.wmppsTransitioning
                '                MediaJumpToMarker()
                If MainForm.tmrSlowMo.Enabled Then
                    Justpaused = False
                    MainForm.SwitchSound(True)
                Else
                    Justpaused = True
                End If

                wmp.Visible = True
            Case Else
                wmp.Visible = False
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
    Public Sub MediaAdvance(wmp As AxWMPLib.AxWindowsMediaPlayer, stp As Long)
        wmp.Ctlcontrols.step(stp)
        wmp.Refresh()
        Media.Position = wmp.Ctlcontrols.currentPosition

    End Sub
End Module
