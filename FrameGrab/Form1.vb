


Imports System.Runtime.InteropServices

    Imports System.Drawing.Imaging



    Public Class Form1



        Private FileName As String

        Private FGM As IGraphBuilder

        Private MC As IMediaControl

    Private WC As IVideoWindow

    Private Bmp As Bitmap
    Public Sub getVideoThumbnail(ByVal file As String, ByVal position As Double, ByVal imageLoc As String, ByVal iconSize As Int32)

        Try
            Dim win As New Form1
            Dim vid As Microsoft.DirectX.AudioVideoPlayback.Video

            vid = New Microsoft.DirectX.AudioVideoPlayback.Video(file)
            Dim width As Int32 = vid.DefaultSize.Width
            Dim height As Int32 = vid.DefaultSize.Height
            win.Width = width
            win.Height = height
            vid.Owner = win.PictureBox1
            vid.Play()
            '            vid.SeekCurrentPosition(position, SeekPositionFlags.absolutePosition)
            vid.Pause()
            Dim thumb As New Bitmap(width, height)
            win.PictureBox1.DrawToBitmap(thumb, New Rectangle(0, 0, width, height))

            Dim newWidth As Integer
            Dim newHeight As Integer
            If width > height Then
                newWidth = iconSize
                newHeight = (iconSize * height) / width
            Else
                newHeight = height
                newWidth = (iconSize * width) / height
            End If
            Dim newThumb As New Bitmap(newWidth, newHeight)
            Dim graphics As Graphics = Graphics.FromImage(DirectCast(newThumb, Image))
            graphics.InterpolationMode = Drawing2D.InterpolationMode.Bicubic

            graphics.DrawImage(thumb, 0, 0, newWidth, newHeight)
            graphics.Dispose()

            newThumb.Save(imageLoc)
        Catch ex As Exception
            MessageBox.Show("An Error (" & ex.ToString & ") Occured while retrieving Thumbnail", "Error Retrieving Thumbnail.", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Failed to Retrieve Thumbnail")
            My.Computer.Clipboard.SetText(ex.ToString)
        End Try
    End Sub











    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        getVideoThumbnail("E:\Pictures\Mvi_5201.avi", 1523, "", "")
    End Sub
    End Class

