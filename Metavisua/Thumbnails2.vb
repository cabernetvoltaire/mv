Imports System.Threading

Public Class Thumbnails
    Public flp2 As New FlowLayoutPanel
    Public t As Thread
    Private flist As List(Of String)
    Private mSize As Int16 = 100
    Private Sub LoadThumbnails()
        Dim flp As New FlowLayoutPanel


        flp.BackColor = Color.Azure
        flp.Dock = DockStyle.Fill
        flp.BringToFront()
        flp.Visible = True

        Dim pics(flist.Count) As PictureBox
        Dim i As Int16 = 0


        For Each f In flist
            Try
                If FindType(f) = Filetype.Pic Then
                    Dim finfo = New IO.FileInfo(f)
                    pics(i) = New PictureBox
                    flp.Controls.Add(pics(i))
                    flp.Update()

                    With pics(i)

                        .Height = mSize

                        If finfo.Extension = ".gif" Then
                            .Image = Image.FromFile(f)

                        Else
                            .Image = GetThumb(p, .Height, f)
                        End If
                        .Width = .Image.Width / .Image.Height * .Height
                        .SizeMode = PictureBoxSizeMode.StretchImage
                        .Tag = f
                        AddHandler .MouseEnter, AddressOf pb_Click
                        '.Image.Dispose()

                    End With
                    'Me.Update()
                End If
                i += 1
            Catch ex As System.InvalidOperationException
                Continue For
            End Try
        Next


    End Sub
    Private Sub pb_Click(sender As Object, e As EventArgs)
        Dim pb = DirectCast(sender, PictureBox)
        strCurrentFilePath = pb.Tag

        frmMain.lbxFiles.SelectionMode = SelectionMode.One
        frmMain.tmrPicLoad.Enabled = True

    End Sub

    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function

    Public Function GetThumb(ByVal e As PaintEventArgs, h As Long, f As String) As Image
        Dim myCallback As New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
        Dim myBitmap As New Bitmap(f)
        Dim ratio As Single = myBitmap.Height / myBitmap.Width
        Dim myThumbnail As Image = myBitmap.GetThumbnailImage(h, h * ratio, myCallback, IntPtr.Zero)
        myBitmap.Dispose()
        Return myThumbnail
        'e.Graphics.DrawImage(myThumbnail, 150, 75)
    End Function

    Private p As PaintEventArgs
    Private Sub Thumbnails_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        p = e


        Loadthumbs()


    End Sub

    Private Sub Loadthumbs()
        Me.Controls.Add(flp2)
        t = New Thread(New ThreadStart(Sub() LoadThumbnails()))
        t.IsBackground = True
        t.Start()

    End Sub



    Public Property FileList() As List(Of String)
        Get
            Return flist
        End Get
        Set(ByVal value As List(Of String))
            flist = value

        End Set
    End Property
    Public Property ThumbnailHeight() As Int16
        Get
            Return mSize
        End Get
        Set(ByVal value As Int16)
            mSize = value
        End Set
    End Property
End Class