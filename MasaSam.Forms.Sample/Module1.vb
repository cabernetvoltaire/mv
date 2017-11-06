Imports System.IO
Module General
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
    Private Const OrientationId As Integer = &H112
    Public blnSpeedRestart As Boolean = False
    Public iSSpeeds() As Integer = {1500, 900, 50}
    Public iPlaybackSpeed() As Decimal = {0.03, 0.5, 0.75}
    Public currentWMP As New AxWMPLib.AxWindowsMediaPlayer
    Public LastPlayed As New Stack(Of String)
    Public blnAutoAdvanceFolder As Boolean = True
    Public blnRandomStartAlways As Boolean = True
    Public blnRestartSlideShowFlag As Boolean = False
    Public Property blnTVCurrent As Boolean
    Public Property strButtonFile As String
    Public Property blnChooseRandomFile As Boolean = True
    Public iCurrentAlpha As Integer = 0
    Public ChosenPlayOrder As Byte = 0
    Public Function ImageOrientation(ByVal img As Image) As ExifOrientations
        ' Get the index of the orientation property.
        Dim orientation_index As Integer = Array.IndexOf(img.PropertyIdList, OrientationId)

        ' If there is no such property, return Unknown.
        If (orientation_index < 0) Then Return ExifOrientations.Unknown

        ' Return the orientation value.
        Return DirectCast(img.GetPropertyItem(OrientationId).Value(0), ExifOrientations)
    End Function

    Public strPlayOrder() As String = {"Original", "Random", "Name", "Path Name", "Date/Time", "Size", "Type"}
    Public Property lngListSizeBytes As Long
    Public PreviewWMP() As AxWMPLib.AxWindowsMediaPlayer = {FindDuplicates.WMP1, FindDuplicates.WMP2, FindDuplicates.WMP3, FindDuplicates.WMP4, FindDuplicates.WMP5,
     FindDuplicates.WMP6, FindDuplicates.WMP7, FindDuplicates.WMP8, FindDuplicates.WMP9, FindDuplicates.WMp10, FindDuplicates.WMP11, FindDuplicates.WMP12}
    Public btnDest() As Button = {frmMain.btn1, frmMain.btn2, frmMain.btn3, frmMain.btn4, frmMain.btn5, frmMain.btn6, frmMain.btn7, frmMain.btn8}
    Public lblDest() As Label = {frmMain.lbl1, frmMain.lbl2, frmMain.lbl3, frmMain.lbl4, frmMain.lbl5, frmMain.lbl6, frmMain.lbl7, frmMain.lbl8}
    Public Enum PlayOrder As Byte
        Original
        Random
        Name
        PathName
        Time
        Length
        Type
    End Enum
    Public lngShowlistLines As Long = 0

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
    Public Autozoomrate As Decimal = 0.4
    Public strVisibleButtons(8) As String

    Public blnButtonsLoaded As Boolean
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

    ''' <summary>
    ''' Copies list from sorted list2
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="list2"></param>
    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of String, String))
        list.Clear()
        For Each m As KeyValuePair(Of String, String) In list2
            list.Add(m.Value)
        Next
    End Sub
    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Long, String))
        list.Clear()
        For Each m As KeyValuePair(Of Long, String) In list2
            list.Add(m.Value)
        Next
    End Sub
    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Date, String))
        list.Clear()
        For Each m As KeyValuePair(Of Date, String) In list2
            list.Add(m.Value)
        Next
    End Sub

    Public Sub Buttons_Load()
        For i As Byte = 0 To 7
            lblDest(i).Font = New Font(lblDest(i).Font, FontStyle.Bold)
        Next
        ' blnButtonsLoaded = True
        InitialiseButtons()
    End Sub
    Public Sub AssignButton(i As Byte, j As Integer, strPath As String)
        strVisibleButtons(i) = strPath
        Dim f As New DirectoryInfo(strPath)
        strButtonFilePath(i, j, 1) = strPath

        lblDest(i).Text = f.Name
        strButtonCaptions(i, j, 1) = f.Name
    End Sub
    ''' <summary>
    ''' Assigns names to buttons
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Public Sub HandleFunctionKeyDown(sender As Object, e As KeyEventArgs)
        Dim i As Byte = e.KeyCode - Keys.F5
        If strVisibleButtons(i) = "" Or e.Shift Then
            AssignButton(i, iCurrentAlpha, CurrentFolderPath) 'Just assign

        Else
            CurrentFolderPath = strVisibleButtons(i) 'Switch to this folder
            frmMain.tvMain2.SelectedFolder = CurrentFolderPath
        End If
    End Sub
    ''' <summary>
    ''' Fill a listbox with a list, according to a filter
    ''' </summary>
    ''' <param name="lbx"></param>
    ''' <param name="Filter"></param>
    ''' <param name="showlist"></param>
    Public Sub FillShowbox(lbx As ListBox, Filter As Byte, showlist As List(Of String))
        If showlist.Count = 0 Then Exit Sub

        frmMain.CollapseShowlist(False)
        lbx.Items.Clear()

        For Each s In showlist
            lbx.Items.Add(s)
        Next
        lbx.TabStop = True

    End Sub
    Public Function SetPlayOrder(Order As Byte, List As List(Of String)) As List(Of String)
        Dim NewListS As New SortedList(Of String, String)
        Dim NewListL As New SortedList(Of Long, String)
        Dim NewListD As New SortedList(Of Date, String)
        For Each f In List
            If Len(f) > 247 Then Continue For
            Dim file As New FileInfo(f)
            Try
                Select Case Order
                    Case PlayOrder.Name
                        NewListS.Add(file.Name & file.FullName, file.FullName)
                    Case PlayOrder.Length
                        NewListL.Add(file.Length + Len(file.FullName), file.FullName)
                    Case PlayOrder.Time
                        Dim time = file.LastWriteTime.AddMilliseconds(Rnd(100))
                        NewListD.Add(time, file.FullName)
                    Case PlayOrder.PathName
                        NewListS.Add(file.FullName, file.FullName)

                    Case PlayOrder.Type
                        NewListS.Add(file.Extension & file.Name & Str(Rnd(100)), file.FullName)
                    Case PlayOrder.Random
                        Try
                            NewListS.Add(Str(Rnd(List.Count)), file.FullName)

                        Catch ex As System.ArgumentException
                            NewListS.Add(Str(Rnd(List.Count)), file.FullName)

                        End Try
                End Select
            Catch ex As System.ArgumentException 'TODO could do better than this. 
                Continue For

            Catch ex As System.IO.PathTooLongException
                Continue For
            End Try
        Next

        If NewListD.Count <> 0 Then
            CopyList(List, NewListD)

        ElseIf NewListS.Count <> 0 Then
            CopyList(List, NewListS)
        ElseIf NewListL.Count <> 0 Then
            CopyList(List, NewListL)

        End If

        Return List




    End Function



End Module
