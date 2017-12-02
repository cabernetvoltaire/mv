Imports System.IO
Public Class FindDuplicates
    Public deletelist As New List(Of String)
    Public uniquelist As New List(Of String)
    Public sortedList As New List(Of String)
    Public uniqueduplist As New List(Of String)
    Public duplist As New List(Of String)
    Public blnClicktoSave As Boolean = True
    Dim DeleteArray(1000, 100) As String
    Dim blnDelete(1000, 100) As Boolean
    Private Sub FindDuplicates_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ArrangePreviews()
        'Add current showlist, in size sorted order, to lbxSorted
        sortedList = SetPlayOrder(PlayOrder.Length, Showlist)
        FillShowbox(lbxsorted, FilterState.All, sortedList)
        uniquelist = ExtractDups(sortedList)
        MsgBox("There are " & uniquelist.Count & " unique files having duplicates, out of " & Showlist.Count)

        FillShowbox(lbxunique, FilterState.All, uniquelist)
        'Create a list of unique files, by ignoring any which have the same length as the previous one.
        '        HighlightPossDups()


    End Sub



    ''' <summary>
    ''' Takes 'sortedlist' and creates deletelist and uniquelist, based on filesize
    ''' </summary>
    Private Sub DeletesandUniques()
        Dim blnDontIncludeLast = True
        Dim lastlength As Long
        Dim lastinfo As FileInfo = Nothing
        Dim finfo3 As FileInfo = Nothing
        For Each file In sortedList
            'Get a file
            finfo3 = My.Computer.FileSystem.GetFileInfo(file)
            Dim currlength As Long
            currlength = finfo3.Length
            If currlength = lastlength Then
                'Same as before - delete
                deletelist.Add(finfo3.FullName)
            Else
                'Otherwise add it the uniques
                uniquelist.Add(finfo3.FullName)
            End If
            lastinfo = finfo3
            lastlength = lastinfo.Length
        Next
    End Sub

    ''' <summary>
    ''' Places 12 Preview controls
    ''' </summary>
    Private Shared Sub ArrangePreviews()
        For i = 0 To 11
            PreviewWMP(i).Width = 250
            PreviewWMP(i).Height = 250
            PreviewWMP(i).Left = 250 * i
            PreviewWMP(i).Top = 0 + 250 * CInt(i > 6)
            Try

                PreviewWMP(i).uiMode = "None"
            Catch ex As System.Runtime.InteropServices.InvalidComObjectException
                Continue For
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Finds duplicates of the given file within the current Showlist and adds them to duplist, which is then written to lbxDuplicates. The appropriate controls then show the files in lbxDuplicates
    ''' </summary>
    ''' <param name="strFilePath"></param>
    Private Sub finddups(strFilePath As String)

        duplist.Clear()
        lbxDuplicates.Items.Clear()
        Dim i As Integer = 0
        If Not File.Exists(strFilePath) Then
            MsgBox("File not found")
            Exit Sub
        End If

        Dim finfo As New IO.FileInfo(strFilePath)

            Dim length As Long = finfo.Length
        For i = 0 To sortedList.Count - 1
            Dim f As New IO.FileInfo(sortedList.Item(i))
            If f.Length = length Then duplist.Add(f.FullName)

        Next
        'Fill the duplicates box
        FillShowbox(lbxDuplicates, FilterState.All, duplist)
        For i = 0 To 11
            Try
                PreviewWMP(i).URL = ""
                PreviewWMP(i).Visible = False

            Catch ex As System.Runtime.InteropServices.InvalidComObjectException
                Continue For
            End Try

        Next
        'Show and split
        i = 0
        deletelist.Clear()
        For Each row In lbxDuplicates.Items
            Dim strpath As String = row.ToString
            If i <= 11 Then
                Try
                    PreviewWMP(i).URL = strpath
                    PreviewWMP(i).Visible = True

                Catch ex As System.Runtime.InteropServices.InvalidComObjectException
                End Try

                If i = 0 Then
                    If Not lbxSave.Items.Contains(row) Then lbxSave.Items.Add(row)
                Else
                    If Not lbxDeleteList.Items.Contains(row) Then lbxDeleteList.Items.Add(row)
                End If
                i += 1
            End If

        Next

    End Sub






    ''' <summary>
    ''' Removes duplicates from List1 and responds with List2
    ''' </summary>
    Private Function ExtractDups(List1 As List(Of String)) As List(Of String)
        Dim uniquedups As New List(Of String)
        Dim lastlength As Long = 0
        Dim lastinfo As FileInfo = Nothing
        Dim finfo3 As FileInfo = Nothing
        Dim blnSame As Boolean
        For Each file In List1

            finfo3 = New FileInfo(file)
            Dim currlength As Long
            currlength = finfo3.Length
            If currlength = lastlength Then
                If lastinfo IsNot Nothing Then
                    If Not blnSame Then
                        uniquedups.Add(lastinfo.FullName)
                        blnSame = True
                    End If
                End If
            Else
                blnSame = False
            End If
            lastinfo = finfo3
            lastlength = lastinfo.Length
        Next
        Return uniquedups
    End Function

    Private Sub lbxunique_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles lbxunique.SelectedIndexChanged, lbxDeleteList.SelectedIndexChanged, lbxSave.SelectedIndexChanged, lbxDuplicates.SelectedIndexChanged, lbxsorted.SelectedIndexChanged
        Dim path As String = sender.Items(sender.SelectedIndex)
        finddups(path)
    End Sub


    Private Sub lbxDeleteList_DoubleClick(sender As Object, e As EventArgs) Handles lbxDeleteList.DoubleClick
        FillShowbox(lbxDeleteList, FilterState.All, deletelist)


    End Sub

    Private Sub btnDeleteFiles_Click(sender As Object, e As EventArgs) Handles btnDeleteFiles.Click
        If MsgBox("Are you sure you want to delete all these files?", MsgBoxStyle.Critical, "DELETE FILES?") = MsgBoxResult.Ok Then
            MsgBox("Click to delete files")
            For Each file In lbxDeleteList.Items
                My.Computer.FileSystem.DeleteFile(file, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)
                RemoveFromListBox(file, lbxDeleteList, deletelist)
                RemoveFromListBox(file, lbxsorted, sortedList)
                RemoveFromListBox(file, lbxDuplicates, duplist)

            Next
            MsgBox("Finished")
        End If
    End Sub



    Private Sub lbxDuplicates_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lbxDuplicates.MouseDoubleClick
        If blnClicktoSave Then
            If Not lbxSave.Items.Contains(lbxDuplicates.SelectedItem) Then lbxSave.Items.Add(lbxDuplicates.SelectedItem)
            If lbxDeleteList.Items.Contains(lbxDuplicates.SelectedItem) Then lbxDeleteList.Items.Remove(lbxDuplicates.SelectedItem)
        Else
            If lbxSave.Items.Contains(lbxDuplicates.SelectedItem) Then lbxSave.Items.Remove(lbxDuplicates.SelectedItem)
            If Not lbxDeleteList.Items.Contains(lbxDuplicates.SelectedItem) Then lbxDeleteList.Items.Add(lbxDuplicates.SelectedItem)

        End If

    End Sub

    Private Sub lbxSave_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lbxSave.MouseDoubleClick
        lbxSave.Items.Remove(lbxSave.SelectedItem)
        If Not lbxDeleteList.Items.Contains(lbxSave.SelectedItem) Then lbxDeleteList.Items.Add(lbxSave.SelectedItem)

    End Sub

    Private Sub lbxDeleteList_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lbxDeleteList.MouseDoubleClick
        lbxDeleteList.Items.Remove(lbxDeleteList.SelectedItem)
        If Not lbxSave.Items.Contains(lbxDeleteList.SelectedItem) Then lbxSave.Items.Add(lbxDeleteList.SelectedItem)

    End Sub

    Private Sub AutoList_Click(sender As Object, e As EventArgs) Handles AutoList.Click
        With lbxunique
            While .SelectedIndex < .Items.Count - 1
                .SelectedItem = .Items(.SelectedIndex + 1)
                frmMain.Update()
            End While
        End With
    End Sub
End Class