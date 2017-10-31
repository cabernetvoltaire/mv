Imports System.ComponentModel
Imports System.IO
Public Class Buttons


    Public Sub InitialiseButtons()
        Dim alph As String = "ABCDEFGH"
        For i As Byte = 0 To 7
            btnDest(i).Text = "F" & Str(i + 5)
            lblDest(i).Text = alph(i)
        Next
    End Sub


    Private Sub Buttons_Load(sender As Object, e As EventArgs) Handles Me.Load
        For i As Byte = 0 To 7
            lblDest(i).Font = New Font(lblDest(i).Font, FontStyle.Bold)
        Next
        blnButtonsLoaded = True
        InitialiseButtons()
        Me.CenterToScreen()
        Me.Top = frmMain.Height * 0.9
    End Sub
    Public Sub AssignButton(i As Byte, strPath As String)
        btnFolderNames(i) = strPath
        Dim f As New DirectoryInfo(strPath)
        lblDest(i).Text = f.Name
    End Sub
    Public Sub HandleKeypress(sender As Object, e As KeyEventArgs)
        Dim i As Byte = e.KeyCode - Keys.F5
        If btnFolderNames(i) = "" Or e.Shift Then
            AssignButton(i, CurrentFolderPath)
        Else
            CurrentFolderPath = btnFolderNames(i)
            frmMain.tvMain2.SelectedFolder = CurrentFolderPath
        End If
    End Sub

    Private Sub Buttons_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        blnButtonsLoaded = False
    End Sub


    Private Sub Buttons_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        frmMain.HandleKeys(sender, e)
        e.Handled = True
    End Sub
End Class