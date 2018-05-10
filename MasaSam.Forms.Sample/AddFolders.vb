﻿Imports System.ComponentModel

Public Class AddFolders
    Private newFolder As String
    Public Property Folder() As String
        Get
            Return newFolder
        End Get
        Set(ByVal value As String)
            newFolder = value
        End Set
    End Property

    Private Sub CreateFolders()
        For Each l In TextBox1.Lines
            Dim s As String = Folder & "\" & l
            IO.Directory.CreateDirectory(s)
            frmMain.tvMain2.RefreshTree(Folder)
        Next
    End Sub

    Private Sub AddFolders_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        CreateFolders()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                Me.Close()
        End Select
    End Sub

    Private Sub AddFolders_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class