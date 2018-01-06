Imports System.IO
Public Class FindDuplicates
    Public initiallist As New List(Of String)
    Public deletelist As New List(Of String)
    Public uniquelist As New List(Of String)
    Public sortedList As New List(Of String)
    Public uniqueduplist As New List(Of String)
    Public duplist As New List(Of String)
    Public blnClicktoSave As Boolean = True
    Dim DeleteArray(1000, 100) As String
    Dim blnDelete(1000, 100) As Boolean

    Dim PreviewWMP(12) As AxWMPLib.AxWindowsMediaPlayer


    Private Sub FindDuplicates_Shown(sender As Object, e As EventArgs) Handles Me.Shown


        ArrangePreviews()
        If Showlist.Count = 0 Then
            initiallist = FileboxContents
        Else
            initiallist = Showlist
        End If
        'Add current showlist, in size sorted order, to lbxSorted
        sortedList = SetPlayOrder(PlayOrder.Length, initiallist)
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
    Private Sub ArrangePreviews()
        Dim size As Integer = 125
        For i = 0 To 11
            PreviewWMP(i) = New AxWMPLib.AxWindowsMediaPlayer

            Panel3.Controls.Add(PreviewWMP(i))
            With PreviewWMP(i)
                If i = 0 Then
                    .Width = size
                    .Height = size
                    .Left = 0
                    .Top = 0
                    .uiMode = "None"
                Else
                    .Width = size
                    .Height = size
                    .Left = (size + 25) * i + 50
                    .Top = 0 + size * CInt(i > 6)
                    .uiMode = "None"

                End If
                AddHandler .MouseMoveEvent, AddressOf previewover
            End With
        Next
    End Sub
    Private Sub previewover(sender As Object, e As AxWMPLib._WMPOCXEvents_MouseMoveEvent)
        ToolTipDups.SetToolTip(sender, sender.url)
    End Sub
    ''' <summary>
    ''' Finds duplicates of the given file within the current Showlist and adds them to duplist, which is then written to lbxDuplicates. The appropriate controls then show the files in lbxDuplicates
    ''' </summary>
    ''' <param name="strFilePath"></param>
    Private Sub finddups(strFilePath As String, Verbose As Boolean)
        'Need to split the algorith from the control handling.
        duplist.Clear()
        lbxDuplicates.Items.Clear()
        Dim i As Integer = 0
        If Not File.Exists(strFilePath) Then
            If Verbose Then MsgBox("File not found")
            Exit Sub
        End If

        Dim finfo As New IO.FileInfo(strFilePath)

        Dim length As Long = finfo.Length
        For i = 0 To sortedList.Count - 1
            Dim f As New IO.FileInfo(sortedList.Item(i))
            If Not f.Exists Then Continue For
            If f.Length = length Then duplist.Add(f.FullName)

        Next
        'Fill the duplicates box
        If Verbose Then
            FillShowbox(lbxDuplicates, FilterState.All, duplist)
            For i = 0 To 11
                Try
                    ' PreviewWMP(i).URL = ""
                    PreviewWMP(i).Visible = False
                    GC.Collect()
                Catch ex As System.Runtime.InteropServices.InvalidComObjectException
                    MsgBox(ex.Message)
                    Continue For
                End Try

            Next
        End If
        'Show and split
        i = 0
        deletelist.Clear()

        For Each row In lbxDuplicates.Items
            Dim strpath As String = row.ToString
            If i <= 11 Then

                If Verbose Then
                    Try
                        PreviewWMP(i).URL = strpath
                        PreviewWMP(i).Visible = True
                        Me.Update()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                        Continue For

                    End Try
                End If
            End If

            If i = 0 Then
                If Verbose And Not lbxSave.Items.Contains(row) Then lbxSave.Items.Add(row)
            Else
                If Not lbxDeleteList.Items.Contains(row) Then
                    deletelist.Add(row)
                    If Verbose Then lbxDeleteList.Items.Add(row)
                End If
            End If
            i += 1
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
        Dim lbx As ListBox = sender
        If lbx.SelectedIndex < 0 Then Exit Sub
        Dim path As String = lbx.Items(lbx.SelectedIndex)
        finddups(path, True)
    End Sub




    Private Sub btnDeleteFiles_Click(sender As Object, e As EventArgs) Handles btnDeleteFiles.Click
        If MsgBox("Are you sure you want to delete all these files?", MsgBoxStyle.Critical, "DELETE FILES?") = MsgBoxResult.Ok Then
            MsgBox("Click to delete files")
            Dim s As New List(Of String)
            For Each file In lbxDeleteList.Items
                s.Add(file)

            Next
            MoveFiles(s, "", lbxDeleteList)
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
        If Not lbxSave.Items.Contains(lbxDeleteList.SelectedItem) Then lbxSave.Items.Add(lbxDeleteList.SelectedItem)
        lbxDeleteList.Items.Remove(lbxDeleteList.SelectedItem)

    End Sub

    Private Sub AutoList_Click(sender As Object, e As EventArgs) Handles AutoList.Click
        For Each path In lbxunique.Items
            finddups(path, False)
        Next

    End Sub

    Private Sub lbxDuplicates_MouseEnter(sender As Object, e As EventArgs) Handles lbxDuplicates.MouseEnter

    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTipDups.Popup

    End Sub

    Private Sub lbxDuplicates_MouseHover(sender As Object, e As EventArgs) Handles lbxDuplicates.MouseHover
        MouseHoverInfo(lbxDuplicates, ToolTipDups)
    End Sub


End Class