Public Class Thumbnails
    Private Sub LoadThumbnails(Flist As List(Of String), e As PaintEventArgs)
        Dim pics(Flist.Count) As PictureBox

        Dim i As Int16 = 0
        For Each f In Flist
            If frmMain.FindType(f) = 0 Then
                pics(i) = New PictureBox
                FlowLayoutPanel1.Controls.Add(pics(i))
                With pics(i)
                    '                    If i > 0 Then DisposePic(pics(i - 1))
                    .Image = GetImage(f)
                    .Height = 200

                    .Width = .Image.Width / .Image.Height * .Height
                    .SizeMode = PictureBoxSizeMode.StretchImage

                End With
                Me.Update()
            End If
            i += 1
        Next
    End Sub


    Private Sub Thumbnails_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        LoadThumbnails(Showlist, e)
    End Sub

End Class