Imports Microsoft.VisualBasic
Public Class GroupingForm
    Private mFolderPath As String
    Private mFolder As IO.DirectoryInfo
    Public Property FolderPath() As String
        Get
            Return mFolderPath
        End Get
        Set(ByVal value As String)
            mFolderPath = value
            mFolder = New IO.DirectoryInfo(mFolderPath)
            MakeList(mFolder.EnumerateFiles, 3, 20)
            '            Subgroups(x.EnumerateFiles, 13)
        End Set
    End Property
    Public Property Min As Integer = 3
    Public Property Max As Integer = 20
    Private Property Starts As New List(Of String)
    Private Sub Subgroups(SearchList As IEnumerable(Of IO.FileInfo), length As Integer)
        'Make list of different start strings of given length '
        ListStarts(SearchList, length)
        CreateRow(SearchList)

    End Sub

    'List all the starts of length MAXLENGTH, and count the number of each
    'For each start which has a count smaller than MINCOUNT, make a new start of length one shorter and repeat
    'Go through the list of files, examining MAXLENGTH strings
    'Tally matches with the check string, 
    'When the check changes, add a row IFF the tally is between MINCOUNT and MAXCOUNT
    '
    'Reduce the length by 1 and repeat

    Private Sub MakeList(Slist As IEnumerable(Of IO.FileInfo), mincount As Integer, maxcount As Integer)
        Dim check As String = ""
        Dim count As Integer = 1
        Dim total As Integer

        For i = 40 To 1 Step -1
            For Each s In Slist
                If LCase(Strings.Left(s.Name, i)) <> LCase(check) Then
                    'Check changed
                    If count >= mincount And count <= maxcount Then
                        Dim val As String = Starts.Find(Function(value As String)
                                                            Return InStr(lcase(value), lcase(check)) <> 0
                                                        End Function)

                        If val <> "" Then
                            '                    MsgBox(LCase(check) & " already added")
                        Else
                            DataGridView1.Rows.Add(New String() {check, count})
                            total = total + count
                            Starts.Add(LCase(check))
                            '                   MsgBox(check & " being added")
                        End If
                    End If
                    check = Strings.Left(s.Name, i)
                    count = 1

                Else
                    count += 1
                End If
            Next
        Next
        DataGridView1.Rows.Add(New String() {"TOTAL", total})

    End Sub
    Private Sub CreateDirectories()
        For Each row In DataGridView1.Rows
            Dim x As String = row(0)

        Next
    End Sub

    Private Sub ListStarts(SearchList As IEnumerable(Of IO.FileInfo), length As Integer)
        Starts.Clear()
        Dim check As String = ""
        For Each s In SearchList
            If LCase(Strings.Left(s.Name, length)) <> LCase(check) Then
                check = Strings.Left(s.Name, length)
                If CountStarters(SearchList, check) > Min Then
                    Starts.Add(check)
                Else

                End If
            End If
        Next
    End Sub
    Private Sub CreateRow(list As IEnumerable(Of IO.FileInfo))
        For Each s In Starts
            Dim m As Integer = (CountStarters(list, s))
            DataGridView1.Rows.Add(New String() {s, m})
        Next

    End Sub
    Private Sub GroupingForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        FolderPath = Media.MediaDirectory
    End Sub

    Private Sub GroupingForm_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        For Each row In DataGridView1.Rows
            If row.Equals(DataGridView1.Rows(DataGridView1.Rows.Count - 1)) Then Exit For
            Dim m As String = row.Cells(0).Value.ToString
            For Each x In mFolder.EnumerateFiles
                If InStr(x.Name, m) = 1 Then
                    'Create folder m if doesn't exist
                    'move x to m
                    Dim subfolder As New IO.DirectoryInfo(mFolder.FullName & "\" & m & "\")
                    If subfolder.Exists Then
                    Else
                        subfolder.Create()
                    End If
                    x.MoveTo(subfolder.FullName & "\" & x.Name)
                End If
            Next
        Next

    End Sub

    Private Function CountStarters(Slist As IEnumerable(Of IO.FileInfo), start As String) As Integer
        Dim count As Integer
        For Each m In Slist
            If m.Name.StartsWith(start) Then
                count += 1
            End If
        Next
        Return count
    End Function
End Class