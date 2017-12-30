Public Module Settings
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
    Public iSSpeeds() As Integer = {1500, 900, 200}
    Public iPlaybackSpeed() As Decimal = {0.25, 0.6, 0.8}
    Public currentWMP As New AxWMPLib.AxWindowsMediaPlayer
    Public LastPlayed As New Stack(Of String)
    Public LastFolder As New Stack(Of String)
    Public blnAutoAdvanceFolder As Boolean = True
    Public blnRestartSlideShowFlag As Boolean = False
    Public blnCopyMode As Boolean = False
    Public Property blnSecondScreen As Boolean = True
    Public Property LastShowList As String
    Public Property blnLink As Boolean
    Public Property blnMoveMode As Boolean = False
    Public Property lastselection As String
    Public Property blnJumpToMark As Boolean = False
    Public blnRandomStartPoint = True
    Public blnRandomStartAlways As Boolean = True
    Public blnRandomAdvance As Boolean = True
    Public blnLoopPlay As Boolean = True
    Public blnChooseRandomFile As Boolean = True


    Public iPropjump As Integer = 10
    Public iQuickJump As Integer = 30
    Public lCurrentDisplayIndex As Long = 0
    Public PlaybackSpeed As Double = 1
    Public lngInterval = 50
    Public lngMediaDuration As Long
    Public lngMark As Long
    Public strCurrentFilePath As String = ""
    Public strExt As String
    Public CurrentFolderPath As String = "E:\"
    Public FileboxContents As New List(Of String)
    Public FBCShown As Boolean()
    Public fType As Filetype

    Public Showlist As New List(Of String)
    Public Oldlist As New List(Of String)

    Public blnDontShowRepeats As Boolean = True
    Public Sublist As New List(Of String)
    Public currentPicBox As New PictureBox
    Public Autozoomrate As Decimal = 0.4
    Public strVisibleButtons(8) As String
    Public NofShown As Int16
    Public blnButtonsLoaded As Boolean = False
    Public ssspeed As Integer = 200
    Public CountSelections As Int16

    Public blnFullScreen As Boolean


    Public strPlayOrder() As String = {"Original", "Random", "Name", "Path Name", "Date/Time", "Size", "Type"}
    Public Property lngListSizeBytes As Long
    Public blnTVCurrent As Boolean

    Public strButtonFile As String
    Public iCurrentAlpha As Integer = 0
    Public ChosenPlayOrder As Byte = 0
    Public Sub PreferencesSave()

        With My.Computer.Registry.CurrentUser
            .SetValue("VertSplit", frmMain.ctrFileBoxes.SplitterDistance)
            .SetValue("HorSplit", frmMain.ctrMainFrame.SplitterDistance)
            .SetValue("Folder", CurrentFolderPath)
            .SetValue("File", strCurrentFilePath)
            .SetValue("Filter", CurrentFilterState)
            .SetValue("LastButtonFolder", strButtonFile)
            .SetValue("LastAlpha", iCurrentAlpha)

        End With

    End Sub
    Public Sub PreferencesGet()
        With My.Computer.Registry.CurrentUser
            frmMain.ctrFileBoxes.SplitterDistance = .GetValue("VertSplit", frmMain.ctrFileBoxes.Height / 4)
            frmMain.ctrMainFrame.SplitterDistance = .GetValue("HorSplit", frmMain.ctrFileBoxes.Width / 2)
            'CurrentFolderPath = .GetValue("Folder", "C:\")
            ChangeFolder(.GetValue("Folder", "C:\"), False)
            strCurrentFilePath = .GetValue("File")
            CurrentFilterState = .GetValue("Filter", 0)
            strButtonFile = .GetValue("LastButtonFolder", "")

            iCurrentAlpha = .GetValue("LastAlpha", 0)

        End With
        If Not IO.Directory.Exists(CurrentFolderPath) Then ChangeFolder("C:\", True)
        If Not IO.File.Exists(strCurrentFilePath) Then strCurrentFilePath = ""
        frmMain.tssMoveCopy.Text = CurrentFolderPath
        frmMain.RandomFunctionsToggle(False)

    End Sub

End Module
