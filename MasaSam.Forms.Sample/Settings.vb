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

            frmMain.ctrTreeandFiles.SplitterDistance = .GetValue("VertSplit")
            frmMain.ctrFilesandPics.SplitterDistance = .GetValue("HorSplit")
            CurrentFolderPath = .GetValue("Folder")
            frmMain.tssMoveCopy.Text = CurrentFolderPath
            strCurrentFilePath = .GetValue("File")
            CurrentFilterState = .GetValue("Filter")
        End With


    End Sub

End Module
