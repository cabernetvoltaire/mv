Imports System.IO
Public Class FindDuplicates
    Public deletelist As New List(Of String)
    Public uniquelist As New List(Of String)
    Public sortedList As New List(Of String)
    Public uniqueduplist As New List(Of String)
    Public duplist As New List(Of String)
    Private Sub FindDuplicates_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ArrangePreviews()
        sortedList = SetPlayOrder(PlayOrder.Length, Showlist)


        HighlightPossDups()

        lblInfo.Text = "Sorted (" & Showlist.Count & " files)"
        FillShowbox(lbxsorted, FilterState.All, Showlist)
        Dim blnDontIncludeLast = True
        Dim lastlength As Long
        Dim lastinfo As FileInfo = Nothing
        Dim finfo3 As FileInfo = Nothing
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
        uniquelist = ExtractDups(Showlist)
        MsgBox("Uniques have" & uniquelist.Count & " files, out of " & Showlist.Count)
        FillShowbox(lbxunique, FilterState.All, uniquelist)

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
        Next
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






    ''' <summary>
    ''' Finds duplicates of the given file within the current Showlist and adds them to duplist, which is then written to lbxDuplicates. The appropriate controls then show the files in lbxDuplicates
    ''' </summary>
    ''' <param name="strFilePath"></param>
    Private Sub finddups(strFilePath As String)

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
                If InStr(row, txtFilter.Text) <> 0 Then
                    If deletelist.Contains(row) Then
                    Else
                        deletelist.Add(row)
                    End If

                Else

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

    Private Sub lbxunique_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles lbxunique.SelectedIndexChanged, lbxDeleteList.SelectedIndexChanged
        Dim path As String = sender.Items(sender.SelectedIndex)
        finddups(path)
    End Sub

    Private Sub lbxunique_DoubleClick(sender As Object, e As EventArgs) Handles lbxunique.DoubleClick
    End Sub

    Private Sub lbxDeleteList_DoubleClick(sender As Object, e As EventArgs) Handles lbxDeleteList.DoubleClick
        FillShowbox(lbxDeleteList, FilterState.All, deletelist)

    End Sub

    Private Sub btnDeleteFiles_Click(sender As Object, e As EventArgs) Handles btnDeleteFiles.Click
        If MsgBox("Are you sure you want to delete all these files?", MsgBoxStyle.Critical, "DELETE FILES?") = MsgBoxResult.Ok Then
            MsgBox("Deleting files (not really)")
        End If
    End Sub
End Class