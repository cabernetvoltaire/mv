Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form2_Click(sender As Object, e As EventArgs) Handles Me.Click
        ' AxWindowsMediaPlayer1.Ctlcontrols.play()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        AxWindowsMediaPlayer1.URL = "Q:\Watch\Keepers\40.wmv"
    End Sub
End Class