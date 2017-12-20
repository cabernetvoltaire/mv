Imports System.ComponentModel

Public Class Thumbnails


    Public Sub LoadThumbnails(Flist As List(Of String), e As PaintEventArgs)
        Dim pics(Flist.Count) As PictureBox

        Dim i As Int16 = 0
        For Each f In Flist
            If frmMain.FindType(f) = 0 Then

                pics(i) = New PictureBox
                FlowLayoutPanel1.Controls.Add(pics(i))
                With pics(i)
                    .Height = 175
                    .Image = Example_GetThumb(e, .Height, f)
                    .Width = .Image.Width / .Image.Height * .Height
                    .SizeMode = PictureBoxSizeMode.StretchImage
                End With
            End If
            i += 1
        Next
    End Sub

    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function

    Public Function Example_GetThumb(ByVal e As PaintEventArgs, h As Long, f As String) As Image
        Dim myCallback As New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
        Dim myBitmap As New Bitmap(f)
        Dim ratio As Single = myBitmap.Height / myBitmap.Width
        Dim myThumbnail As Image = myBitmap.GetThumbnailImage(h, h * ratio, myCallback, IntPtr.Zero)
        myBitmap.Dispose()
        Return myThumbnail
        'e.Graphics.DrawImage(myThumbnail, 150, 75)
    End Function


    Private Sub Thumbnails_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        LoadThumbnails(Showlist, e)
    End Sub

End Class