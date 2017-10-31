Public Class Buttons


    Public Sub InitialiseButtons()
        Dim alph As String = "ABCDEFGH"
        For i As Byte = 0 To 7
            btnDest(i).Text = "f" & Str(i + 5)
            lblDest(i).Text = alph(i)
        Next
    End Sub




    Private Sub Buttons_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

    End Sub

    Private Sub Buttons_Load(sender As Object, e As EventArgs) Handles Me.Load
        InitialiseButtons()
    End Sub


End Class