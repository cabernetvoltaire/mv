﻿''' <summary>
''' Given a folder, move all files whose names contain the name of a subfolder to that subfolder
''' </summary>

Class DateMove
    Public Event FilesMoved(sender As Object, e As EventArgs)
    Private mFolder As String

    Private mFolders As New List(Of IO.DirectoryInfo)

    Public Enum DMY As Byte
        Year
        Month
        Day
        Hour
        Minute


    End Enum
    Private MonthNames As String() = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
    Private SizeNames As String() = {"Size 0- Tiny", "Size 1-Small", "Size 2-Medium", "Size 3-Large", "Size 4-Very Large", "Size 5-Gigantic"}
    Public Property Folder() As String
        Get
            Return mFolder
        End Get
        Set(ByVal value As String)
            mFolder = value
        End Set
    End Property

    Public Sub New()
        mRecursive = False
    End Sub

    Private mRecursive As Boolean
    Public Property Recursive() As Boolean
        Get
            Return mRecursive
        End Get
        Set(ByVal value As Boolean)
            mRecursive = value
        End Set
    End Property

    Public Sub FilterByDate(FolderName As String, Recurse As Boolean, Choice As DMY)

        'For each file
        'Move to folder with the year, creating it first

        Dim s As New IO.DirectoryInfo(FolderName)
        Dim i As Integer
        For Each f In s.GetFiles
            With GetDate(f)
                Select Case Choice
                    Case DMY.Year
                        i = .Year
                    Case DMY.Month
                        i = .Month
                    Case DMY.Day
                        i = .Day


                    Case DMY.Hour
                        i = .Hour
                    Case DMY.Minute
                        i = .Minute
                End Select
            End With
            Dim folname As String = Str(i)
            If Choice = DMY.Month Then
                folname = MonthName(i)
            End If
            'If f.Directory.EnumerateDirectories(Str(i)) Is Nothing Then
            f.Directory.CreateSubdirectory(folname & "\")
            'End If
            f.MoveTo(f.DirectoryName & "\" & folname & "\" & f.Name)
        Next
        RaiseEvent FilesMoved(Nothing, Nothing)

    End Sub
    Public Sub FilterBySize(FolderName As String, Recurse As Boolean)

        'For each file
        'Move to folder with the year, creating it first

        Dim s As New IO.DirectoryInfo(FolderName)
        Dim i As Integer
        For Each f In s.GetFiles
            For m = 5 To 10
                If f.Length < 10 ^ m And f.Length >= 10 ^ (m - 1) Then
                    f.Directory.CreateSubdirectory(SizeNames(m - 5) & "\")
                    f.MoveTo(f.DirectoryName & "\" & SizeNames(m - 5) & "\" & f.Name)
                End If
            Next

            'If f.Directory.EnumerateDirectories(Str(i)) Is Nothing Then
        Next
        RaiseEvent FilesMoved(Nothing, Nothing)

    End Sub
    Private Sub GetFoldersBelow(folderpath As String, recurse As Boolean)

        Dim s As New IO.DirectoryInfo(folderpath)
        For Each m In s.GetDirectories
            If recurse Then GetFoldersBelow(m.FullName, recurse)
            mFolders.Add(m)
        Next

    End Sub
End Class

