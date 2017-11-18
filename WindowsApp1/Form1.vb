Imports MasaSam.Forms.Controls

Public Class Form1


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        AxWindowsMediaPlayer1.URL = "Q:\Adobe Flash Player 13_04_2017 01_54_38"
        AxWindowsMediaPlayer1.Ctlcontrols.play()
    End Sub
End Class
