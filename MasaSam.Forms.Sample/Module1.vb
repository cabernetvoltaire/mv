Module General

    Public PreviewWMP() As AxWMPLib.AxWindowsMediaPlayer = {FindDuplicates.WMP1, FindDuplicates.WMP2, FindDuplicates.WMP3, FindDuplicates.WMP4, FindDuplicates.WMP5,
     FindDuplicates.WMP6, FindDuplicates.WMP7, FindDuplicates.WMP8, FindDuplicates.WMP9, FindDuplicates.WMp10, FindDuplicates.WMP11, FindDuplicates.WMP12}
    Public blnRandomStartPoint = False
    Public PlaybackSpeed As Double = 1
    Public lngInterval = 50
    Public lngMediaDuration As Long
    Public lngMark As Long
    Public iPropjump As Integer = 15
    Public iQuickJump As Integer = 20
    Public strCurrentFilePath As String = ""
    Public strExt As String
    Public Property CurrentFolderPath As String = "E:\"
    Public FileboxContents As New List(Of String)
    Public FBCShown As New List(Of Boolean)
    Public blnDontShowRepeats As Boolean = True
    Public lCurrentDisplayIndex As Long = 0
    Public fType As Filetype
    Public blnRandom As Boolean = False
    Public Showlist As New List(Of String)
    Public Sublist As New List(Of String)
    Public currentPicBox As New PictureBox
    Public Autozoomrate As Decimal = 0.2

    Public ssspeed As Integer = 200

    Public blnFullScreen As Boolean

    Public Orientation() As String = {"Unknown", "TopLeft", "TopRight", "BottomRight", "BottomLeft", "LeftTop", "RightTop", "RightBottom", "LeftBottom"}
    Public Enum Filetype As Byte
        Pic
        Movie
        Doc
        Gif
        Xcel
        Browsable
        Unknown
    End Enum
    Public SlideShowSpeeds() As Integer = {50, 100, 500, 1000, 2000, 9000}
    Public VideoSpeeds() As Integer = {10, 20, 50, 75, 100}
    Public Function TimeOperation(blnStart As Boolean) As TimeSpan
        Static StartTime As Date
        If blnStart Then
            StartTime = Now
            Return Now - StartTime
        Else
            Return Now - StartTime
        End If
    End Function
End Module
