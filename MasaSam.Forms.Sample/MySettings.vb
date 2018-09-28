
Public Module Mysettings


    Public PFocus As Byte = CtrlFocus.Tree
    Public Property ZoneSize As Decimal = 0.4
    Public Const OrientationId As Integer = &H112
    Public currentWMP As New AxWMPLib.AxWindowsMediaPlayer
    Public WithEvents Media As New MediaHandler
    'Public WithEvents NavigateMoveState As New StateHandler
    'Public WithEvents CurrentFilterState As New FilterHandler
    'Public WithEvents PlayOrder As New SortHandler
    'Public WithEvents StartPoint As New StartPointHandler
    'Public WithEvents Random As New RandomHandler
#Region "Options"

    Public iSSpeeds() As Integer = {1000, 300, 200}
    Public iPlaybackSpeed() As Integer = {3, 15, 45}

    Public LastPlayed As New Stack(Of String)
    Public LastFolder As New Stack(Of String)
    Public Property LastShowList As String
    Public Property blnJumpToMark As Boolean = False
    Public blnLoopPlay As Boolean = True
    'Public blnChooseRandomFile As Boolean = True
    Public PlaybackSpeed As Double = 30
    Public FavesFolderPath As String = "Q:\Favourites"
    Public strButtonFile As String = "Q:\Alpha2.msb"

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
            .SetValue("VertSplit", MainForm.ctrFileBoxes.SplitterDistance)
            .SetValue("HorSplit", MainForm.ctrMainFrame.SplitterDistance)

            .SetValue("Folder", Media.MediaDirectory)
            .SetValue("File", Media.MediaPath)
            .SetValue("Filter", MainForm.CurrentFilterState.State)
            .SetValue("SortOrder", MainForm.PlayOrder.State)
            .SetValue("StartPoint", MainForm.StartPoint.State)
            .SetValue("State", MainForm.NavigateMoveState.State)
            .SetValue("LastButtonFolder", strButtonFile)
            .SetValue("LastAlpha", iCurrentAlpha)
        End With

    End Sub
    Public Sub PreferencesGet()
        MainForm.ctrPicAndButtons.SplitterDistance = 9 * MainForm.ctrPicAndButtons.Height / 10
        With My.Computer.Registry.CurrentUser
            'Media.MediaDirectory = .GetValue("Folder", System.Environment.GetFolderPath(Environment.SpecialFolder.MyPictures))
            Media.MediaPath = .GetValue("File", "C:\exiftool.exe")

            MainForm.ctrFileBoxes.SplitterDistance = .GetValue("VertSplit", MainForm.ctrFileBoxes.Height / 4)
            MainForm.ctrMainFrame.SplitterDistance = .GetValue("HorSplit", MainForm.ctrFileBoxes.Width / 2)
            MainForm.CurrentFilterState.State = .GetValue("Filter", 0)
            MainForm.PlayOrder.State = .GetValue("SortOrder", 0)
            MainForm.StartPoint.State = .GetValue("StartPoint", 0)
            MainForm.NavigateMoveState.State = .GetValue("State", 0)
            strButtonFile = .GetValue("LastButtonFolder", "Q:\Alpha2.msb")

            iCurrentAlpha = .GetValue("LastAlpha", 0)

        End With
        '     If Not IO.Directory.Exists(Media.MediaDirectory) Then Media.MediaDirectory = "C:\"
        If Not IO.File.Exists(Media.MediaPath) Then Media.MediaPath = ""
        MainForm.tssMoveCopy.Text = Media.MediaDirectory
        '   frmMain.RandomFunctionsToggle(False)

    End Sub
    Public Sub OnMediaChanged(sender As Object, e As EventArgs) Handles Media.MediaChanged

        ChangeFolder(Media.MediaDirectory, True)

    End Sub

End Module
