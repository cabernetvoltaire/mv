Imports System.Text.RegularExpressions

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Main()
    End Sub
    Public Sub Main()
        Dim input As String = TextBox1.Text
        Dim pattern As String = TextBox2.Text
        'If Regex.IsMatch(input) Then
        '    TextBox3.AppendText(item)

        'End If 
    End Sub
End Class
