Module Settings
    Public Enum ExifOrientations As Byte
        Unknown = 0
        TopLeft = 1
        TopRight = 2
        BottomRight = 3
        BottomLeft = 4
        LeftTop = 5
        RightTop = 6
        RightBottom = 7
        LeftBottom = 8
    End Enum

    Public Enum CtrlFocus As Byte
        Tree = 0
        Files = 1
        ShowList = 2
    End Enum
    Public PFocus As Byte = CtrlFocus.Tree

    Public Const OrientationId As Integer = &H112
    Public blnSpeedRestart As Boolean = False
    Public iSSpeeds() As Integer = {1500, 900, 50}
    Public iPlaybackSpeed() As Decimal = {0.03, 0.5, 0.75}
    Public currentWMP As New AxWMPLib.AxWindowsMediaPlayer
    Public LastPlayed As New Stack(Of String)
    Public blnAutoAdvanceFolder As Boolean = True
    Public blnRandomStartAlways As Boolean = True
    Public blnRestartSlideShowFlag As Boolean = False
    Public blnCopyMode As Boolean = False
    Public blnChooseRandomFile As Boolean = True
    Public blnTVCurrent As Boolean

    Public strButtonFile As String
    Public iCurrentAlpha As Integer = 0
    Public ChosenPlayOrder As Byte = 0
    Public Sub PreferencesSave()

        With My.Computer.Registry.CurrentUser
            .SetValue("VertSplit", frmMain.ctrTreeandFiles.SplitterDistance)
            .SetValue("HorSplit", frmMain.ctrFilesandPics.SplitterDistance)
            .SetValue("Folder", CurrentFolderPath)
            .SetValue("File", strCurrentFilePath)
            .SetValue("Filter", CurrentFilterState)
            .SetValue("LastButtonFolder", strButtonFile)
        End With

    End Sub
    Public Sub PreferencesGet()
        With My.Computer.Registry.CurrentUser
            frmMain.ctrTreeandFiles.SplitterDistance = .GetValue("VertSplit", frmMain.ctrTreeandFiles.Height / 4)
            frmMain.ctrFilesandPics.SplitterDistance = .GetValue("HorSplit", frmMain.ctrTreeandFiles.Width / 2)
            CurrentFolderPath = .GetValue("Folder", "C:\")
            strCurrentFilePath = .GetValue("File")
            CurrentFilterState = .GetValue("Filter", 0)
            strButtonFile = .GetValue("LastButtonFolder", "")
        End With
        If Not IO.Directory.Exists(CurrentFolderPath) Then CurrentFolderPath = "C:\"
        If Not IO.File.Exists(strCurrentFilePath) Then strCurrentFilePath = ""
        frmMain.tssMoveCopy.Text = CurrentFolderPath


    End Sub

End Module
