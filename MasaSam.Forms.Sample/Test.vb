

Public Class Test

    Private Sub ImageQueueStart(flist As List(Of String), startpoint As Integer, size As Integer)
        For i As Integer = startpoint To startpoint + size
            imgList.Add(ImageFromNumber(flist, i))
        Next
    End Sub
    Private Function ImageFromNumber(flist As List(Of String), i As Integer) As Image
        Dim pic As New IO.FileInfo(flist(i))
        Return Image.FromFile(pic.FullName)

    End Function

    Private Sub GetImages(flist As List(Of String))
        '
        'Actually much better if it gets passed a list of names and then
        'only loads, say six images (experiment) and does a 'tankroll' through, loading the others in as it goes
        '(Use a different thread?)
        ' PictureBox1.Dock = DockStyle.Fill

        Dim max As Integer = mediaLoopSize

        Dim i As Integer = 0
        Dim k As Integer = 0

        ProgressBar1.Maximum = flist.Count - 1
        ProgressBar1.Visible = True
        Try
            ImageQueueStart(flist, 0, max)
            'While k <= max And i <= flist.Count - 1 ' And My.Computer.Info.AvailableVirtualMemory > 100000000

            '    Dim pic As New System.IO.FileInfo(flist(i))
            '    If LCase(pic.Extension) = ".jpg" Or LCase(pic.Extension) = ".jpeg" Then
            '        imgQueue.Add(Image.FromFile(pic.FullName))
            '        imgQueue(i).Tag = pic.Name
            '        ProgressBar1.Value = k
            '        k += 1

            '    End If
            '    i += 1
            'End While
        Catch ex As OutOfMemoryException

            Exit Try
            'Catch ex As IndexOutOfRangeException
            '    Exit Try
        Catch ex As Exception
            MessageBox.Show(ex.Message, ex.GetType.ToString)
        End Try
        TrackBar1.Maximum = flist.Count - 1
        ProgressBar1.Visible = False
        MsgBox("This is 0")
        PreparePic(pbx1, pbxBlanker, imgList(0))

    End Sub







    Private Sub Test_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        '  Me.MdiParent = frmMain
        For Each s In frmMain.lbxFiles.Items
            LboxFiles.Add(s)
        Next

        GetImages(LboxFiles)

    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Dim K As Integer = TrackBar1.Value
        K = K Mod mediaLoopSize
        Label1.Text = K & " of " & imgList.Count
        PreparePic(pbx1, pbxBlanker, imgList(K))

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles pnlFS.Click, pbx1.Click
        iScreenstate = (iScreenstate + 1) Mod 2
        PreparePic(pbx1, pbxBlanker, ImageFromNumber(LboxFiles, TrackBar1.Value))
    End Sub

    Private Sub Test_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel ', pbx1.MouseWheel 'Hack mousewheel detected twice.
        'PictureFunctions.Mousewheel()
        'ePicMousePoint.X = e.X
        'ePicMousePoint.Y = e.Y
        'If sender.Equals(pbx1) Then
        '    ePicMousePoint.X = ePicMousePoint.X + pbx1.Left
        '    ePicMousePoint.Y = ePicMousePoint.Y + pbx1.Top

        'End If
        '' MsgBox(e.Delta.ToString)
        'If ShiftDown And iScreenstate <> Screenstate.Fitted Then
        '    ' iScreenstate = Screenstate.Zoomed
        '    If CtrlDown Then
        '        ZoomPicture(pbx1, e.Delta > 0, 10) 'TODO Options
        '    Else
        '        ZoomPicture(pbx1, e.Delta > 0, 2)

        '    End If

        'Else
        '    TrackBar1.Value = Math.Max(Math.Min(TrackBar1.Value - Math.Sign(e.Delta) * (1), TrackBar1.Maximum), 0)
        '    TrackBar1_Scroll(Me, e)
        'End If

    End Sub

    Private Sub MovePic(mouse As Point, inside As Control, outside As Control)
        '    Dim x As Long
        '    Dim y As Long
        '    ' mouse = PointToScreen(mouse)
        '    x = mouse.X
        '    y = mouse.Y
        '    'Move left
        '    Dim xdist As Long
        '    Dim ydist As Long

        '    If bImageDimensionState = PicState.TooWide Or bImageDimensionState = PicState.Overscan Then
        '        If x > pbxBlanker.Right Then
        '            'Move right
        '            xdist = x - pbxBlanker.Right
        '            'Large pic width
        '            If inside.Width > outside.Width Then
        '                If inside.Left > outside.Width - inside.Width Then
        '                    inside.Left = inside.Left - xdist / lSpeedfactor
        '                End If
        '            Else 'Small pic width
        '                If inside.Left > 0 Then
        '                    inside.Left = inside.Left - xdist / lSpeedfactor
        '                End If

        '            End If

        '        ElseIf x < pbxBlanker.Left Then
        '            xdist = x - pbxBlanker.Left

        '            'Large pic width
        '            If inside.Width > outside.Width Then
        '                If inside.Left < 0 Then
        '                    inside.Left = inside.Left - xdist / lSpeedfactor
        '                End If
        '            Else 'small pic width
        '                If inside.Left < outside.Width - inside.Width Then
        '                    inside.Left = inside.Left - xdist / lSpeedfactor
        '                End If

        '            End If
        '        End If
        '    End If

        '    If bImageDimensionState = PicState.TooTall Or bImageDimensionState = PicState.Overscan Then

        '        If y < pbxBlanker.Top Then 'Move down
        '            ydist = y - pbxBlanker.Top
        '            'Large pic height
        '            If inside.Height > outside.Height Then
        '                If inside.Top < 0 Then
        '                    inside.Top = inside.Top - ydist / lSpeedfactor
        '                End If
        '            Else
        '                If inside.Top < outside.Height - inside.Height Then
        '                    inside.Top = inside.Top - ydist / lSpeedfactor
        '                End If
        '            End If
        '            '

        '        ElseIf y > pbxBlanker.Bottom Then 'Move Up
        '            ydist = y - pbxBlanker.Bottom
        '            'Large pic height
        '            If inside.Height > outside.Height Then
        '                If inside.Top > outside.Height - inside.Height Then
        '                    inside.Top = inside.Top - 3 * ydist / lSpeedfactor
        '                End If
        '            Else
        '                If inside.Top > 0 Then
        '                    inside.Top = inside.Top - 3 * ydist / lSpeedfactor
        '                End If
        '            End If

        '        End If


        '    End If
        '    inside.Refresh()

    End Sub
    Private Sub Test_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control
        frmMain.HandleKeys(sender, e)
        If e.KeyCode = Keys.Escape Then
            For Each i In imgList
                i.Dispose()
            Next
            'picList.Clear()
            Me.Dispose()
            Me.Close()
        End If
    End Sub
    Private Sub pbx1_MouseMove(sender As Object, e As MouseEventArgs) Handles pnlFS.MouseMove, pbx1.MouseMove
        Dim mouse As Point
        mouse.X = e.X
        mouse.Y = e.Y
        If sender.Equals(pbx1) Then
            mouse.X = mouse.X + pbx1.Left
            mouse.Y = mouse.Y + pbx1.Top
        End If
        If Not ShiftDown Then MovePic(mouse, pbx1, pnlFS)
        '        MovePic(mouse, pbx1, pnlFS) 'Todo OPTION zoom prevents move.
    End Sub
    Private Sub Test_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        ShiftDown = e.Shift
        CtrlDown = e.Control
    End Sub


End Class