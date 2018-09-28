Public Class Listform
    Private Sub ShowList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FileBox1.FolderPath = Media.MediaDirectory
        MainForm.lbxFiles.SelectionMode = SelectionMode.One
    End Sub
    Public Sub OnSelectionChange() Handles FileBox1.FileSelected
        MainForm.tmrPicLoad.Enabled = True
    End Sub

    Private Sub Listform_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

    End Sub

    Private Sub Listform_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        MainForm.frmMain_KeyDown(sender, e)
    End Sub
End Class