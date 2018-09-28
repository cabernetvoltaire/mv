Module ListboxHandler
    Private WithEvents lbxf As ListBox = MainForm.lbxFiles
    Private WithEvents lbxs As ListBox = MainForm.lbxShowList
    Public Event FileListSelectionChanged()
    Public Event ShowListSelectionChange()

    Public Sub ListBoxChanged(sender As Object, e As EventArgs) Handles lbxf.SelectedIndexChanged

        With sender
            Dim i As Long = .SelectedIndex
            If .Items.count = 0 Then
                .items.add("If there is nothing showing here, check your filters")
            ElseIf i >= 0 Then
                Media.MediaPath = .items(i)
            End If
        End With
    End Sub
End Module
