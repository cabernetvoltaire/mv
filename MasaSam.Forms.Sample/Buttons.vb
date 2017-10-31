Public Class Buttons
    Public btnDest() As Button = {btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8}
    Public lblDest() As Label = {lbl1, lbl2, lbl3, lbl4, lbl5, lbl6, lbl7, lbl8}

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