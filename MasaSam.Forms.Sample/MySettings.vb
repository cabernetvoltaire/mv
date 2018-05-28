
Public Module Mysettings


    Public PFocus As Byte = CtrlFocus.Tree
    Public Property ZoneSize As Decimal = 0.4
    Public Const OrientationId As Integer = &H112
    Public currentWMP As New AxWMPLib.AxWindowsMediaPlayer
    Public WithEvents Media As New MediaHandler
    Public WithEvents NavigateMoveState As New StateHandler
    Public WithEvents CurrentFilterState As New FilterHandler
    Public WithEvents PlayOrder As New SortHandler
    Public WithEvents StartPoint As New StartPointHandler
    Public WithEvents Random As New RandomHandler
#Region "Options"

    Public iSSpeeds() As Integer = {1500, 600, 200}
    Public iPlaybackSpeed() As Integer = {3, 10, 20}
    Public iPropjump As Integer = 10
    Public iQuickJump As Integer = 30
    Public lCurrentDisplayIndex As Long = 0
    Public LastPlayed As New Stack(Of String)
    Public LastFolder As New Stack(Of String)
    Public Property blnCopyMode As Boolean = False
    Public Property blnMoveMode As Boolean = True
    Public Property LastShowList As String
    Public Property blnJumpToMark As Boolean = False
    Public blnRandomStartAlways As Boolean = True
    '  Public blnRandomAdvance(3) As Boolean
    Public blnLoopPlay As Boolean = True
    'Public blnChooseRandomFile As Boolean = True
    Public PlaybackSpeed As Double = 30
    Public strCurrentFilePath As String = ""
    Public CurrentFolderPath As String = "E:\"
    Public FavesFolderPath As String = "Q:\Favourites"
    Public strButtonFile As String

    Public Autozoomrate As Decimal = 0.4
    Public iCurrentAlpha As Integer = 0
#End Region

#Region "Internal"
    Public Property blnSecondScreen As Boolean = True
    'Public blnAutoAdvanceFolder As Boolean = True
    Public blnRestartSlideShowFlag As Boolean = False
    Public Property blnLink As Boolean
    Public Property lastselection As String


    Public lngInterval = 10
    Public lngMediaDuration As Long
    Public lngMark As Long
    Public strExt As String
    Public FileboxContents As New List(Of String)
    Public FBCShown As Boolean()
    Public fType As Filetype

    Public Showlist As New List(Of String)
    Public Oldlist As New List(Of String)

    Public blnDontShowRepeats As Boolean = True
    Public Sublist As New List(Of String)
    Public currentPicBox As New PictureBox
    Public strVisibleButtons(8) As String
    Public NofShown As Int16
    Public blnButtonsLoaded As Boolean = False
    Public ssspeed As Integer = 200
    'Public CountSelections As Int16
    Public blnFullScreen As Boolean
    Public Property lngListSizeBytes As Long
    'Public blnTVCurrent As Boolean

    Public ChosenPlayOrder As Byte = 0
#End Region
    Public Sub PreferencesSave()

        With My.Computer.Registry.CurrentUser
            .SetValue("VertSplit", frmMain.ctrFileBoxes.SplitterDistance)
            .SetValue("HorSplit", frmMain.ctrMainFrame.SplitterDistance)
            .SetValue("Folder", CurrentFolderPath)
            .SetValue("File", strCurrentFilePath)
            .SetValue("Filter", frmMain.CurrentFilterState.State)
            .SetValue("LastButtonFolder", strButtonFile)
            .SetValue("LastAlpha", iCurrentAlpha)

        End With

    End Sub
    Public Sub PreferencesGet()
        frmMain.ctrPicAndButtons.SplitterDistance = 9 * frmMain.ctrPicAndButtons.Height / 10
        With My.Computer.Registry.CurrentUser
            frmMain.ctrFileBoxes.SplitterDistance = .GetValue("VertSplit", frmMain.ctrFileBoxes.Height / 4)
            frmMain.ctrMainFrame.SplitterDistance = .GetValue("HorSplit", frmMain.ctrFileBoxes.Width / 2)
            CurrentFolderPath = .GetValue("Folder", frmMain.MyPictures)
            strCurrentFilePath = .GetValue("File")
            frmMain.CurrentFilterState.State = .GetValue("Filter", 0)
            strButtonFile = .GetValue("LastButtonFolder", "")

            iCurrentAlpha = .GetValue("LastAlpha", 0)

        End With
        If Not IO.Directory.Exists(CurrentFolderPath) Then CurrentFolderPath = "C:\"
        If Not IO.File.Exists(strCurrentFilePath) Then strCurrentFilePath = ""
        frmMain.tssMoveCopy.Text = CurrentFolderPath
        '   frmMain.RandomFunctionsToggle(False)

    End Sub

End Module
