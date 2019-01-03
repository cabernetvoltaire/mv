Imports System.ComponentModel
Imports MasaSam.Forms.Controls

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
    Public KeepDisplaying As Boolean = True
    Public Property ButtonNumber() As Byte
        Get
            Return newButtonNumber
        End Get
        Set(ByVal value As Byte)
            newButtonNumber = value
        End Set
    End Property
    Public Property Showing As Boolean = True
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
            If value <> "" Then

                PlayMedia(value)
            End If
        End Set
    End Property
    Public Property Button As Button
    Private Sub PlayMedia(value As String)
        Dim x As New IO.DirectoryInfo(value)
        Dim PreMH As New MediaHandler
        Dim count As Integer = x.GetFiles.Count
        If count = 0 Then

        Else

            Dim i As Integer = Int(Rnd() * (count - 1))

            Dim f As IO.FileInfo

            f = x.GetFiles.ToArray(i)
            ' med.Bookmark = med.Duration * Rnd()
            PreMH.Player = PreviewWMP
            PreMH.MediaPath = f.FullName
            If PreMH.MediaType = Filetype.Movie Or PreMH.MediaType = Filetype.Pic Then

                If f.Exists Then
                    PreviewWMP.URL = f.FullName
                    PreviewWMP.settings.mute = True
                    PreviewWMP.uiMode = "none"
                    PreviewWMP.stretchToFit = True
                    'PreviewWMP.Visible = True
                    PreviewWMP.Ctlcontrols.play()
                End If
            End If

        End If
    End Sub

    Private newChosenFolder As Integer
    Public ReadOnly Property ChosenFolder() As Integer
        Get
            Dim s As String
            If TextBox1.Text <> "" Then
                s = Folder & "\" & TextBox1.Text
            Else
                s = Folder

            End If

            newChosenFolder = s
            Return newChosenFolder
        End Get
    End Property
    Public Sub ChangeMedia()
        PlayMedia(newFolder)
    End Sub

    Private Sub btnAssign_Click(sender As Object, e As EventArgs) Handles btnAssign.Click
        AssignButton(ButtonNumber, Alpha, 1, Folder, True)
        Me.Close()
    End Sub
    Private Sub Label1_DoubleClick(sender As Object, e As EventArgs) Handles Label1.DoubleClick
        fst1.SelectedFolder = newFolder
    End Sub

    Private Sub FolderSelect_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If Not My.Computer.FileSystem.DirectoryExists(Folder) Then
            CreateNewDirectory(fst1, Folder, False)
        End If
    End Sub
    Private Sub fst1_Paint(sender As Object, e As PaintEventArgs) Handles fst1.Paint
        fst1.SelectedFolder = newFolder
    End Sub

    Private Sub FolderSelect_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
        Me.Close()
    End Sub

    Private Sub fst1_DirectorySelected(sender As Object, e As DirectoryInfoEventArgs) Handles fst1.DirectorySelected
        Folder = e.Directory.FullName
    End Sub
End Class