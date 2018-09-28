Public Class FileBox
    ' Inherits ListBox
    Public Event FileSelected()

    Private mFolderPath As String
    Public Property FolderPath() As String
        Get
            Return mFolderPath
        End Get
        Set(ByVal value As String)
            Static lastpath
            mFolderPath = value
            If value <> lastpath Then
                Try
                    ListBox1.Items.Clear()
                    Dim x = New IO.DirectoryInfo(mFolderPath)
                    For Each m In x.EnumerateFiles
                        ListBox1.Items.Add(m.FullName)
                    Next

                Catch ex As Exception
                    ReportFault("Filebox.FolderPath", ex.Message)
                End Try
            End If
        End Set
    End Property

    Private mList As List(Of String)
    Public Property List() As List(Of String)
        Get
            Return mList
        End Get
        Set(ByVal value As List(Of String))
            mList = value
            If Not IsNothing(mList) Then

                For Each s In mList
                    ListBox1.Items.Add(s)
                Next
            End If
        End Set
    End Property

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Media.MediaPath = ListBox1.Items(ListBox1.SelectedIndex)
        RaiseEvent FileSelected()
    End Sub

    Private Sub FileBox_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
    End Sub
End Class
