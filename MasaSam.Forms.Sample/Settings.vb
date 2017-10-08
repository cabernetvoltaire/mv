Module Settings
    Public Sub PreferencesSave()
        With My.Computer.Registry.CurrentUser
            '.CreateSubKey(VertSplit)
            '.CreateSubKey("HorSplit")
            '.CreateSubKey(Folder)
            '.CreateSubKey(File)

            .SetValue("VertSplit", frmMain.ctrTreeandFiles.SplitterDistance)
            .SetValue("HorSplit", frmMain.ctrFilesandPics.SplitterDistance)
            .SetValue("Folder", CurrentFolderPath)
            .SetValue("File", strCurrentFilePath)
            .SetValue("Filter", CurrentFilterState)
        End With

    End Sub
    Public Sub PreferencesGet()
        With My.Computer.Registry.CurrentUser

            frmMain.ctrTreeandFiles.SplitterDistance = .GetValue("VertSplit", frmMain.ctrTreeandFiles.Height / 4)
            frmMain.ctrFilesandPics.SplitterDistance = .GetValue("HorSplit", frmMain.ctrTreeandFiles.Width / 2)

            CurrentFolderPath = .GetValue("Folder", "C:\")
            If Not IO.Directory.Exists(CurrentFolderPath) Then CurrentFolderPath = "C:\"

            frmMain.tssMoveCopy.Text = CurrentFolderPath
            strCurrentFilePath = .GetValue("File")
            If Not IO.File.Exists(strCurrentFilePath) Then strCurrentFilePath = ""

            CurrentFilterState = .GetValue("Filter", 0)
        End With


    End Sub

End Module
