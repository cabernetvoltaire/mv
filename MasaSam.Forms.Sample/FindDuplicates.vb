Public Class FindDuplicates
    Public deletelist As New List(Of String)
    Public uniquelist As New List(Of String)
    Public sortedList As New List(Of String)
    Public uniqueduplist As New List(Of String)
    Public duplist As New List(Of String)
    Private Sub FindDuplicates_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        For i = 0 To 11
            PreviewWMP(i).Width = 250
            PreviewWMP(i).Height = 250
            PreviewWMP(i).Left = 250 * i
            PreviewWMP(i).Top = 0 + 250 * CInt(i > 6)
        Next

        'Just fills the DGV and sorts it
        Dim row As String()
        For Each f In Showlist
            If System.IO.File.Exists(f) Then
                Dim finfo As New System.IO.FileInfo(f)
                row = {False, finfo.Name, finfo.Length, finfo.CreationTime, finfo.FullName}
                Me.dgv1.Rows.Add(row)
            End If

        Next
        Me.dgv1.Sort(dgv1.Columns(2), System.ComponentModel.ListSortDirection.Ascending)
        For Each srow In dgv1.Rows
            If srow.cells(4).value <> Nothing Then
                Dim finfo2 As New IO.FileInfo(srow.Cells(4).value)
                sortedList.Add(finfo2.FullName)
            End If
        Next
        HighlightPossDups()

        For Each file In sortedList
            lbxsorted.Items.Add(file)

        Next
        lblSorted.Text = lblSorted.Text & " (" & sortedList.Count & " files)"


        Dim blnDontIncludeLast = True
        Dim lastlength As Long
        Dim lastinfo As System.IO.FileInfo = Nothing
        Dim finfo3 As System.IO.FileInfo = Nothing
        For Each file In sortedList
            finfo3 = My.Computer.FileSystem.GetFileInfo(file)
            Dim currlength As Long
            currlength = finfo3.Length
            If currlength = lastlength Then
                deletelist.Add(finfo3.FullName)
            Else
                uniquelist.Add(finfo3.FullName)
            End If
            lastinfo = finfo3
            lastlength = lastinfo.Length
        Next
        MakeSortedUnique()
    End Sub
    Private Sub HighlightPossDups()
        Dim blnDontIncludeLast = True
        Dim lastlength As Long
        Dim lastpath As String = ""
        Dim lastinfo As IO.FileInfo = Nothing

        For Each file In sortedList

            Dim finfo As New IO.FileInfo(file)
            Dim currlength As Long
            currlength = finfo.Length
            Dim currpath As String = finfo.FullName
            If currlength = lastlength And currpath <> lastpath Then
                If Not blnDontIncludeLast Then
                    duplist.Add(lastinfo.FullName)

                End If
                duplist.Add(file)
                blnDontIncludeLast = True
            Else
                blnDontIncludeLast = False
            End If
            lastinfo = finfo
            lastpath = finfo.FullName
            lastlength = finfo.Length
        Next
        lbxsorted.Items.Clear()
        For Each f In duplist
            lbxsorted.Items.Add(f)
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        HighlightPossDups()
    End Sub



    Private Sub dgv1_Keydown(sender As Object, e As KeyEventArgs) Handles dgv1.KeyDown
        frmMain.HandleKeys(sender, e)
        Select Case e.KeyCode
            Case Keys.Down, Keys.Up, Keys.Left, Keys.Right, Keys.Escape
            Case Keys.Escape
                Me.Hide()
            Case Else
                e.Handled = True

        End Select
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs)
        'creates flist from only one file from each. 

    End Sub



    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        For Each file In deletelist
            Dim finfo As New IO.FileInfo(file)
            Dim fname As String
            fname = finfo.Name

            Label2.Text = "Deleted " & file
            Try
                'IO.File.Move(file, "E:\Experiment\" & fname)
                My.Computer.FileSystem.DeleteFile(file, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)
            Catch ex As IO.IOException

            End Try
        Next
    End Sub




    Private Sub lbxdelete_DoubleClick(sender As Object, e As EventArgs) Handles lbxdelete.DoubleClick
        For Each file In deletelist
            lbxdelete.Items.Add(file)
        Next
        lblDelete.Text = lblDelete.Text & " (" & deletelist.Count & " files)"

    End Sub

    Private Sub lbxunique_DoubleClick(sender As Object, e As EventArgs) Handles lbxunique.DoubleClick
        For Each file In uniquelist
            lbxunique.Items.Add(file)
        Next
        lblUnique.Text = lblUnique.Text & " (" & uniquelist.Count & " files)"
        Showlist = uniquelist

    End Sub


    Private Sub lbxunique_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxunique.SelectedIndexChanged, lbxdelete.SelectedIndexChanged, lbxsorted.SelectedIndexChanged, lbxDuplicates.SelectedIndexChanged
        With sender
            strCurrentFilePath = .Items(.SelectedIndex)
            If sender.Equals(lbxsorted) Or sender.Equals(lbxunique) Then
                finddups(strCurrentFilePath)
            End If
            ' = frmMain.strCurrentFilePath
        End With
    End Sub

    Private Sub finddups(strFilePath As String)
        'Finds duplicates of the given path within the current Showlist
        'and adds them to duplist, which is then written to lbxDuplicates
        'The appopriate controls then show the files in lbxDuplicates



        duplist.Clear()
        lbxDuplicates.Items.Clear()
        Dim i As Integer = 0
        Dim finfo As New IO.FileInfo(strFilePath)
        Dim length As Long = finfo.Length
        For i = 0 To sortedList.Count - 1
            Dim f As New IO.FileInfo(sortedList.Item(i))
            If f.Length = length Then
                duplist.Add(f.FullName)
                lbxDuplicates.Items.Add(f.FullName)
                '  ElseIf f.Length > length Then
                '     Exit For
            End If
        Next
        'For i = 0 To 3
        '    If lbxDuplicates.Items(i) Is Nothing Then Exit For
        '    Dim strpath As String = lbxDuplicates.Items(i).ToString
        '    If lbxDuplicates.Items(i).ToString <> "" Then
        '        PreviewWMP(i).URL = strpath
        '    End If
        'Next
        For i = 0 To 11
            PreviewWMP(i).URL = ""
            PreviewWMP(i).Visible = False

        Next
        i = 0

        For Each row In lbxDuplicates.Items
            Dim strpath As String = row.ToString
            If i <= 11 Then
                PreviewWMP(i).URL = strpath
                PreviewWMP(i).Visible = True
                lbxDeleteList.Items.Add(row)
            Else

            End If
            i += 1
        Next

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) 
        lbxDeleteList.Items.Add(strCurrentFilePath)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        For Each file In lbxDeleteList.Items
            System.IO.File.Delete(file)
            'lbxDeleteList.Items.Remove(file)
        Next
    End Sub

    Private Sub lbxsorted_DoubleClick(sender As Object, e As EventArgs) Handles lbxsorted.DoubleClick
        MakeSortedUnique()
    End Sub

    Private Sub MakeSortedUnique()
        Dim blnDontIncludeLast = True
        Dim lastlength As Long
        Dim lastinfo As System.IO.FileInfo = Nothing
        Dim finfo3 As System.IO.FileInfo = Nothing
        For Each file In sortedList

            finfo3 = My.Computer.FileSystem.GetFileInfo(file)
            Dim currlength As Long
            currlength = finfo3.Length
            If currlength = lastlength Then
                lbxsorted.Items.Remove(finfo3.FullName)
            Else
            End If
            lastinfo = finfo3
            lastlength = lastinfo.Length
        Next
    End Sub

    Private Sub SplitContainer1_Panel2_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainer1.Panel2.Paint

    End Sub
End Class