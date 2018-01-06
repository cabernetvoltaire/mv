Imports System.ComponentModel

Public Class FolderSelect
    Private newAlpha As Int16
    Public Property Alpha() As Int16
        Get
            Return newAlpha
        End Get
        Set(ByVal value As Int16)
            newAlpha = value
        End Set
    End Property
    Private newButtonNumber As Byte
    Public Property ButtonNumber() As Byte
        Get
            Return newButtonNumber
        End Get
        Set(ByVal value As Byte)
            newButtonNumber = value
        End Set
    End Property

    Private newFolder As String
    Public Property Folder() As String
        Get
            newFolder = fst1.SelectedFolder
            Return newFolder
        End Get
        Set(ByVal value As String)


            newFolder = value
            '  fst1.SelectedFolder = newFolder
            Label1.Text = value
        End Set
    End Property
    Private newChosenFolder As Integer
    Public ReadOnly Property ChosenFolder() As Integer
        Get
            Dim s As String
            If TextBox1.Text <> "" Then
                s = Folder & "\" & TextBox1.Text
            Else
                s = Folder

            End If
            If Not My.Computer.FileSystem.DirectoryExists(s) Then
                My.Computer.FileSystem.CreateDirectory(s)
            End If
            newChosenFolder = s
            Return newChosenFolder
        End Get
    End Property

    Private Sub btnAssign_Click(sender As Object, e As EventArgs) Handles btnAssign.Click
        Me.Close()
    End Sub
    Private Sub Label1_DoubleClick(sender As Object, e As EventArgs) Handles Label1.DoubleClick
        fst1.SelectedFolder = newFolder
    End Sub

    Private Sub FolderSelect_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Folder <> strButtonFilePath(ButtonNumber, Alpha, 1) Then
            AssignButton(ButtonNumber, Alpha, 1, Folder, True)
        End If
    End Sub
    Private Sub fst1_Paint(sender As Object, e As PaintEventArgs) Handles fst1.Paint
        fst1.SelectedFolder = newFolder
    End Sub
End Class