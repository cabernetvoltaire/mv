Public Class Group
    Private mFolder As IO.DirectoryInfo
    Public Property Folder() As IO.DirectoryInfo
        Get
            Return mFolder
        End Get
        Set(ByVal value As IO.DirectoryInfo)
            mFolder = value
        End Set
    End Property

    Private Sub DivideFolder(SortType As Byte)
        For Each f In mFolder.EnumerateFiles()
            Select Case SortType

            End Select

        Next
    End Sub
End Class
