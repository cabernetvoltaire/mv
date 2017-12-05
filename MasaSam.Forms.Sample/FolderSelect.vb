﻿Imports System.ComponentModel

Public Class FolderSelect
    Private newFolder As String
    Public Property Folder() As String
        Get
            newFolder = FileSystemTree1.SelectedFolder
            Return newFolder
        End Get
        Set(ByVal value As String)
            newFolder = value
            FileSystemTree1.SelectedFolder = value
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

    Private Sub FolderSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnAssign_Click(sender As Object, e As EventArgs) Handles btnAssign.Click

        Me.Close()
    End Sub
End Class