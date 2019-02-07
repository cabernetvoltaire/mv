Imports AxWMPLib
Module MovieHandler
    ' Public Event MediaEnded()

    Public Property MediaMarker As Long = -1
    Public Property NewPosition As Long
    Public Sub MediaJumpToMarker(SP As StartPointHandler)
        'TODO Move to MediaHandler
        'ReportTime("Jumpto Marker")
        If Media.MediaPath = "" Then Exit Sub
        If MediaMarker <> -1 Then
            NewPosition = MediaMarker
            Debug.Print("MediaMarker is " & MediaMarker)
        Else
            NewPosition = SP.StartPoint
            Console.WriteLine("New position is " & NewPosition & " of " & Media.Duration)
        End If
        MainForm.tmrJumpVideo.Enabled = True

    End Sub

    Public Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent)
        Dim wmp As AxWindowsMediaPlayer = CType(sender, AxWindowsMediaPlayer)


        'ReportTime("Playstate " & e.newState)

        Static Justpaused As Boolean
        'MsgBox(e.newState.ToString)
        Select Case e.newState
            Case WMPLib.WMPPlayState.wmppsMediaEnded
                'wmp.Visible = False
                If Not MainForm.tmrAutoTrail.Enabled Then
                    MainForm.AdvanceFile(True, False)
                End If
            Case WMPLib.WMPPlayState.wmppsPlaying
                'ReportTime("Playing")
                Media.Duration = MainForm.currentWMP.currentMedia.duration
                MainForm.StartPoint.Duration = Media.Duration
                MainForm.SwitchSound(False)
                If Justpaused Then
                    Justpaused = False
                    Exit Sub
                End If
                '                MediaDuration = currentWMP.currentMedia.duration
                If FullScreen.Changing Or MainForm.SP.Unpause Then 'Hold current position if switching to FS or back. 
                    NewPosition = MainForm.currentWMP.Ctlcontrols.currentPosition
                    MainForm.tmrJumpVideo.Enabled = True
                Else
                    ' ReportFault("MHPSCHange", "Just before MJ2M", True)
                    ' MediaJumpToMarker()
                    MainForm.OnStartChanged()
                End If
                wmp.Visible = True
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
    Public Sub PlaystateChangeNew(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent, ByRef SP As StartPointHandler, ByRef MH As MediaHandler)
        'TODO Move to MediaHandler
        Dim wmp As AxWindowsMediaPlayer = CType(sender, AxWindowsMediaPlayer)
        'ReportTime("Playstate " & e.newState)

        Static Justpaused As Boolean
        'MsgBox(e.newState.ToString)
        Select Case e.newState
            Case WMPLib.WMPPlayState.wmppsMediaEnded
                'wmp.Visible = False
                If Not MainForm.tmrAutoTrail.Enabled Then
                    MainForm.AdvanceFile(True, False)
                End If
            Case WMPLib.WMPPlayState.wmppsPlaying
                'ReportTime("Playing")
                MH.Duration = wmp.currentMedia.duration
                SP.Duration = MH.Duration
                MainForm.SwitchSound(False)
                If Justpaused Then
                    Justpaused = False
                    Exit Sub
                End If
                '                MediaDuration = currentWMP.currentMedia.duration
                If FullScreen.Changing Or MainForm.SP.Unpause Then 'Hold current position if switching to FS or back. 
                    NewPosition = wmp.Ctlcontrols.currentPosition
                    MainForm.tmrJumpVideo.Enabled = True

                Else
                    ' ReportFault("MHPSCHange", "Just before MJ2M", True)
                    ' MediaJumpToMarker()
                    wmp.Ctlcontrols.currentPosition = SP.StartPoint

                    MainForm.OnStartChanged()
                End If
                'wmp.Visible = True
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

                'wmp.Visible = True
            Case Else
                '         wmp.Visible = False
        End Select
    End Sub
    'TODO Move to MediaHandler

End Module
