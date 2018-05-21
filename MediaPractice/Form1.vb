Imports AxWMPLib
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MHandler.Player = AxWindowsMediaPlayer1
        'MHandler.URL = "Q:\Poppers2WIP.wmv"
        'MHandler.Player.Ctlcontrols.play()
        'AxWindowsMediaPlayer1.URL = MHandler.URL
        'AxWindowsMediaPlayer1.Ctlcontrols.play()
    End Sub

    Private Sub Form1_Click(sender As Object, e As EventArgs) Handles Me.Click

        Dim mHandler As New MediaHandler
        mHandler.Position = mHandler.currentMedia.duration * Rnd()
        mHandler.Player.Ctlcontrols.play()
        Me.FlowLayoutPanel1.Controls.Add(mHandler)

    End Sub
End Class

