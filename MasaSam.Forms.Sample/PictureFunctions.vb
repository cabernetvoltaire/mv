Module PictureFunctions
    Public LboxFiles As New List(Of String)

    Public imgList As New List(Of Image)
    'Properties - to put and get from in registry
    Public Property iScreenstate As Byte = Screenstate.Fitted
    Public Property lSpeedfactor As Long = 50 'Larger is slower
    Public Property iZoomFactor As Integer = 100
    Public mediaLoopSize As Integer = 50
    Public picBlanker As PictureBox



    'Enums and structures
    Public Enum Screenstate As Byte
        Fitted
        TrueSize
        Zoomed
    End Enum


    'Misc. globals
    Public ePicMousePoint As Point
    Public Property bImageDimensionState As Byte
    Public Property ShiftDown As Boolean
    Public Property CtrlDown As Boolean
    Public Property ZoneSize As Decimal = 0.7

    Public Function ClassifyImage(lvpw As Long, lvph As Long, liw As Long, lih As Long) As Byte
        Dim s As String
        If lvpw > liw Then 'Viewport wider than pic
            If lvph > lih Then
                ClassifyImage = PicState.Underscan
                s = "underscan"
            Else
                ClassifyImage = PicState.TooTall
                s = "tootall"
            End If
        Else 'Viewport narrower than pic
            If lvph > lih Then
                ClassifyImage = PicState.TooWide
                s = "toowide"
            Else 'Viewport shorter than pic
                ClassifyImage = PicState.Overscan
                s = "overscan"
            End If
        End If

    End Function
    Public Enum PicState As Byte
        Overscan
        Underscan
        TooWide
        TooTall

    End Enum
    Public Sub PicClick(pic As PictureBox)
        iScreenstate = (iScreenstate + 1) Mod 2
        DisposePic(pic)
        Dim img As Image = GetImage(strCurrentFilePath)
        PreparePic(pic, picBlanker, img)
    End Sub
    Public Sub DisposePic(box As PictureBox)

        box.Image.Dispose()

        GC.SuppressFinalize(box)
        box.Image = Nothing
    End Sub
    Public Sub Mousewheel(pbx1 As PictureBox, sender As Object, e As MouseEventArgs)
        ePicMousePoint.X = e.X
        ePicMousePoint.Y = e.Y
        If sender.Equals(pbx1) Then
            ePicMousePoint.X = ePicMousePoint.X + pbx1.Left
            ePicMousePoint.Y = ePicMousePoint.Y + pbx1.Top

        End If
        If ShiftDown Then 'And iScreenstate <> Screenstate.Fitted Then
            'If iScreenstate = Screenstate.Fitted Then
            '    'Dim zoomfactor As Decimal = pbx1.Width / pbx1.Image.Width
            '    'pbx1.Width = pbx1.Width / zoomfactor
            '    'pbx1.Height = pbx1.Height / zoomfactor

            '    'iScreenstate = Screenstate.Zoomed
            '    Exit Sub

            'End If

            If CtrlDown Then
                ZoomPicture(pbx1, e.Delta > 0, 2) 'TODO Options
            Else
                ZoomPicture(pbx1, e.Delta > 0, 15)

            End If

        Else
            'Wheel advances file if nothing else held
            frmMain.AdvanceFile(e.Delta < 0, False)
            frmMain.tmrSlideShow.Enabled = False 'Break slideshow if scrolled
            Dim img As Image = GetImage(strCurrentFilePath)
            PreparePic(pbx1, picBlanker, img)
        End If

    End Sub

    Public Function GetImage(strPath As String) As Image
        If strPath = "" Then Return Nothing
        Try
            Dim img As Image = Image.FromFile(strPath)
            Return img
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing

        End Try
    End Function

    Public Sub PreparePic(pbx As PictureBox, pbxBlanker As PictureBox, img As Image)
        If img Is Nothing Then Exit Sub
        If Not pbx.Image Is Nothing Then
            pbx.Image.Dispose()
        End If
        pbx.Image = img
        CentralCtrl(pbxBlanker, ZoneSize) 'insert central zone

        iZoomFactor = 100 * pbx.Width / pbx.Image.Width 'How much is image zoomed currently?

        'Make pic box same size as image (still within container)
        pbx.Width = pbx.Image.Width
        pbx.Height = pbx.Image.Height
        'How does it exceeed the container, if at all?
        If Not iScreenstate = Screenstate.Fitted Then
            ' bImageDimensionState = ClassifyImage(pnlFS.Width, pnlFS.Height, pbx1.Width, pbx1.Height)
        End If

        SetState(pbx, iScreenstate)

        PlacePic(pbx)

    End Sub
    Private Sub CentralCtrl(pbx As Control, proportion As Decimal)
        Dim ctr As Control = pbx.GetContainerControl
        With pbx
            .Width = ctr.Width * proportion
            .Height = ctr.Height * proportion
            .Left = (ctr.Width - .Width) / 2
            .Top = (ctr.Height - .Height) / 2

        End With
    End Sub
    Public Sub MouseMove(pbx1 As PictureBox, sender As Object, e As MouseEventArgs)
        Dim mouse As Point
        mouse.X = e.X
        mouse.Y = e.Y
        If sender.Equals(pbx1) Then
            mouse.X = mouse.X + pbx1.Left
            mouse.Y = mouse.Y + pbx1.Top
        End If
        If Not ShiftDown Then MovePic(mouse, picBlanker, pbx1, pbx1.GetContainerControl)
        '        MovePic(mouse, pbx1, pnlFS) 'Todo OPTION zoom prevents move.
    End Sub
    Public Sub PlacePic(ByRef pbx As PictureBox)
        Dim outside As Control = pbx.GetContainerControl
        'If iScreenstate <> Screenstate.Fitted Then
        '    pbx.Width = pbx.Width * iZoomFactor
        '    pbx.Height = pbx.Height * iZoomFactor
        'End If
        Select Case iScreenstate
            Case Screenstate.Zoomed
                pbx.Left = outside.Width / 2 - pbx.Width / 2
                pbx.Top = outside.Height / 2 - pbx.Height / 2
            Case Screenstate.TrueSize, Screenstate.Fitted
                pbx.Left = outside.Width / 2 - pbx.Width / 2
                pbx.Top = outside.Height / 2 - pbx.Height / 2

        End Select

    End Sub

    Public Sub SetState(pbx As PictureBox, Sstate As Byte)
        'Sets the screenstate, docking style. Changes the sizemode of pbx
        'If iScreenstate = Sstate Then Exit Sub
        Select Case Sstate
            Case Screenstate.Fitted
                pbx.Dock = DockStyle.Fill
                pbx.SizeMode = PictureBoxSizeMode.Zoom
            Case Screenstate.TrueSize
                pbx.Dock = DockStyle.None
                pbx.SizeMode = PictureBoxSizeMode.Normal
                'Expand or shrink the picture box accordingly
                pbx.Width = pbx.Image.Width
                pbx.Height = pbx.Image.Height

            Case Screenstate.Zoomed

                pbx.Dock = DockStyle.None 'TODO It's how we prepare for this which is the problem 

                pbx.SizeMode = PictureBoxSizeMode.Zoom


        End Select
        If iScreenstate <> Screenstate.Fitted Then
        End If
    End Sub
    Public Sub FadeInLabel(lbl As Label)
        lbl.Visible = True
        'lbl.BackColor = TransparencyKey
        'Label1.ForeColor = Color.
        'for i = 0 To 255
        '    For j = 0 To 10000 : Next
        'Next
    End Sub
    Public Sub MovePic(mouse As Point, blanker As PictureBox, inside As Control, outside As Control)
        Dim x As Long
        Dim y As Long
        ' mouse = PointToScreen(mouse)
        x = mouse.X
        y = mouse.Y
        'Move left
        Dim xdist As Long
        Dim ydist As Long

        If bImageDimensionState = PicState.TooWide Or bImageDimensionState = PicState.Overscan Then
            If x > blanker.Right Then
                'Move right
                xdist = x - blanker.Right
                'Large pic width
                If inside.Width > outside.Width Then
                    If inside.Left > outside.Width - inside.Width Then
                        inside.Left = inside.Left - xdist / lSpeedfactor
                    End If
                Else 'Small pic width
                    If inside.Left > 0 Then
                        inside.Left = inside.Left - xdist / lSpeedfactor
                    End If

                End If

            ElseIf x < blanker.Left Then
                xdist = x - blanker.Left

                'Large pic width
                If inside.Width > outside.Width Then
                    If inside.Left < 0 Then
                        inside.Left = inside.Left - xdist / lSpeedfactor
                    End If
                Else 'small pic width
                    If inside.Left < outside.Width - inside.Width Then
                        inside.Left = inside.Left - xdist / lSpeedfactor
                    End If

                End If
            End If
        End If

        If bImageDimensionState = PicState.TooTall Or bImageDimensionState = PicState.Overscan Then

            If y < blanker.Top Then 'Move down
                ydist = y - blanker.Top
                'Large pic height
                If inside.Height > outside.Height Then
                    If inside.Top < 0 Then
                        inside.Top = inside.Top - ydist / lSpeedfactor
                    End If
                Else
                    If inside.Top < outside.Height - inside.Height Then
                        inside.Top = inside.Top - ydist / lSpeedfactor
                    End If
                End If
                '

            ElseIf y > blanker.Bottom Then 'Move Up
                ydist = y - blanker.Bottom
                'Large pic height
                If inside.Height > outside.Height Then
                    If inside.Top > outside.Height - inside.Height Then
                        inside.Top = inside.Top - 3 * ydist / lSpeedfactor
                    End If
                Else
                    If inside.Top > 0 Then
                        inside.Top = inside.Top - 3 * ydist / lSpeedfactor
                    End If
                End If

            End If


        End If
        inside.Refresh()

    End Sub

    Public Sub ZoomPicture(pbx As PictureBox, blnEnlarge As Boolean, Percentage As Decimal)
        'hack zooming a fitted picture.

        'If iScreenstate = Screenstate.Fitted Then
        '    Dim width As Long
        '    Dim height As Long
        '    width = pbx.Width
        '    height = pbx.Height

        SetState(pbx, Screenstate.Zoomed)

        '    pbx.Width = width
        '    pbx.Height = height
        '    PlacePic(pbx)
        'End If

        Dim Enlarge As Decimal = 1 + Percentage / 100
        Dim Reduce As Decimal = 1 - Percentage / 100
        If blnEnlarge Then
            iZoomFactor = iZoomFactor * Enlarge
            pbx.Width = pbx.Width * Enlarge
            pbx.Left = pbx.Left - (ePicMousePoint.X - pbx.Left) * Percentage / 100

            pbx.Top = pbx.Top - (ePicMousePoint.Y - pbx.Top) * Percentage / 100
            pbx.Height = pbx.Height * Enlarge

        Else
            iZoomFactor = iZoomFactor * Reduce
            pbx.Width = pbx.Width * Reduce 'Todo Need limitations to prevent going smaller than container asymmetrically.
            pbx.Left = pbx.Left + (ePicMousePoint.X - pbx.Left) * Percentage / 100
            pbx.Top = pbx.Top + (ePicMousePoint.Y - pbx.Top) * Percentage / 100
            pbx.Height = pbx.Height * Reduce
        End If
        '  pbx.Refresh()
        'bImageDimensionState = ClassifyImage(pnlFS.Width, pnlFS.Height, pbx.Width, pbx.Height)


    End Sub


End Module

