Imports System.IO

Module ButtonHandling
    Public buttons As New ButtonSet

    Public nletts As Int16 = 36
    '   Public layers(nletts) As Byte
    ' Public currentlayer As Byte
    Public strButtonFilePath(8, nletts, 1) As String
    ' Public fButtonDests(8, nletts, 3) As IO.DirectoryInfo
    Private iAlphaCount = nletts
    Public strButtonCaptions(8, nletts, 1)

    Public btnDest() As Button = {frmMain.btn1, frmMain.btn2, frmMain.btn3, frmMain.btn4, frmMain.btn5, frmMain.btn6, frmMain.btn7, frmMain.btn8}

    Public lblDest() As Label = {frmMain.lbl1, frmMain.lbl2, frmMain.lbl3, frmMain.lbl4, frmMain.lbl5, frmMain.lbl6, frmMain.lbl7, frmMain.lbl8}
    Public Sub Buttons_Load()
        buttons.CurrentLetter = "A"
        For i As Byte = 0 To 7
            lblDest(i).Font = New Font(lblDest(i).Font, FontStyle.Bold)
        Next
        ' blnButtonsLoaded = True
        InitialiseButtons()
    End Sub

    Public Sub UpdateButton(strPath As String, strDest As String)
        For i = 0 To nletts - 1
            For j = 0 To 7
                Dim s As String = strButtonFilePath(j, i, 1)
                If Not s Is Nothing Then
                    If s = strPath Then
                        strButtonFilePath(j, i, 1) = strDest
                    ElseIf InStr(strPath, s) <> 0 Then
                        strButtonFilePath(j, i, 1) = Replace(strButtonFilePath(j, i, 1), strPath, strDest)
                    End If
                End If
            Next
        Next

    End Sub


    ''' <summary>
    ''' Assigns all the buttons in a generation, beneath sPath, to the letter iAlpha
    ''' </summary>
    Public Sub AssignLinear(sPath As String, iAlpha As Integer, blnNext As Boolean)
        Dim d As New DirectoryInfo(sPath)
        Dim i As Byte
        Dim di() As DirectoryInfo
        di = d.GetDirectories
        Dim n = d.GetDirectories.Count - 1
        If n > 0 Then
            For i = 0 To n
                Dim k As Byte
                k = i Mod 8

                AssignButton(k, iAlpha + Int(i / 8), 1, di(i).FullName)
            Next
        End If

        KeyAssignmentsStore(strButtonFile)

    End Sub
    Public Sub AssignAlphabetic(blntest As Boolean)

        Dim dlist As New List(Of String)
        Dim d As New DirectoryInfo(Media.MediaDirectory)
        FindAllFoldersBelow(d, dlist, True, True)
        ' dlist = SetPlayOrder(PlayOrder.Name, dlist)
        dlist.Sort()
        Dim n(nletts) As Integer
        Dim layer As Int16 = 1
        For i = 0 To dlist.Count - 1
            Dim s As String = dlist.Item(i)
            Dim sht As String = New DirectoryInfo(s).Name
            Dim l As String = UCase(sht(0))
            Dim k As Int16 = ButtfromAsc(Asc(l))
            If k >= 0 AndAlso k < nletts Then
                If (n(k) Mod 8) = 0 Then
                    layer += 1
                    ReDim Preserve strButtonFilePath(8, nletts, layer)
                Else
                End If
                AssignButton(n(k), k, layer, s)
                n(k) += 1

            End If

        Next
        KeyAssignmentsStore(strButtonFile)
    End Sub
    Public Sub AssignTree(strStart As String)

        Dim dlist As New SortedList(Of String, DirectoryInfo)
        Dim plist As New SortedList(Of String, DirectoryInfo)
        Dim exclude As String = ""
        exclude = InputBox("String to exclude from folders?", "")
        Dim d As New DirectoryInfo(strStart)
        For Each di In d.EnumerateDirectories
            dlist.Add(di.Name, di)
        Next
        Dim i As Int16 = 0

        Dim n(nletts) As Integer
        While i < 8 AndAlso (n(0) < 8 OrElse n(1) < 8 OrElse n(2) < 8 OrElse n(3) < 8 OrElse n(4) < 8 OrElse n(5) < 8 OrElse n(6) < 8 OrElse n(7) < 8 OrElse n(8) < 8 OrElse n(9) < 8 OrElse n(10) < 8 OrElse n(11) < 8 OrElse n(12) < 8 OrElse n(1) < 8 OrElse n(13) < 8 OrElse n(14) < 8 OrElse n(15) < 8 OrElse n(17) < 8 OrElse n(18) < 8 OrElse n(19) < 8 OrElse n(20) < 8 OrElse n(21) < 8 OrElse n(22) < 8 OrElse n(23) < 8 OrElse n(24) < 8)

            For Each di In dlist.Values
                Dim l As String = UCase(di.Name(0))
                Dim ButtonNumber As Int16 = ButtfromAsc(Asc(l))
                If ButtonNumber >= 0 AndAlso ButtonNumber < nletts Then
                    If exclude <> "" AndAlso InStr(di.Name, exclude) <> 0 Then
                    Else
                        If n(ButtonNumber) < 8 Then
                            AssignButton(n(ButtonNumber), ButtonNumber, 1, di.FullName)
                            n(ButtonNumber) += 1
                        End If
                    End If
                End If


            Next

            plist = FindNextTier(dlist)
            dlist = plist
            i += 1
        End While


        KeyAssignmentsStore(strButtonFile)
    End Sub
    Public Function FindNextTier(tier As SortedList(Of String, DirectoryInfo)) As SortedList(Of String, DirectoryInfo)
        Dim i As Int16 = 0
        Dim nexttier As New SortedList(Of String, DirectoryInfo)
        For Each d In tier.Values
            Try
                For Each di In d.EnumerateDirectories
                    nexttier.Add(di.Name + Str(i), di)
                    i += 1

                Next
            Catch ex As Exception
                Continue For
            End Try
        Next
        Return nexttier
    End Function


    Public Sub AssignButton(ByVal i As Byte, ByVal j As Integer, ByVal k As Byte, ByVal strPath As String, Optional blnStore As Boolean = False)
        Dim f As New DirectoryInfo(strPath)
        With buttons.CurrentRow.Row(i)
            .Path = strPath
            .Label = f.Name
        End With
        If strVisibleButtons(i) <> "" And NavigateMoveState.State = StateHandler.StateOptions.Navigate Then
            If Not MsgBox("Replace button assignment for F" & i + 5 & " with " & f.Name & "?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then Exit Sub
        End If
        strVisibleButtons(i) = strPath
        strButtonFilePath(i, j, k) = strPath

        lblDest(i).Text = f.Name
        strButtonCaptions(i, j, k) = f.Name
        UpdateButtonAppearance()
        If blnStore Then
            KeyAssignmentsStore(strButtonFile)
        End If
    End Sub
    Public Sub ChangeButtonLetter(e As KeyEventArgs)

        frmMain.lblAlpha.Text = e.KeyCode.ToString
        iCurrentAlpha = ButtfromAsc(e.KeyCode)
        UpdateButtonAppearance()

    End Sub
    Public Function ButtfromAsc(asc As Integer) As Integer
        Dim n As Integer
        If asc <= 57 Then
            n = asc - 48 + 26

        Else
            n = asc - 65
        End If
        Return n
    End Function
    Public Function AscfromButt(button As Integer) As Integer
        Dim asc As Integer
        If button <= 25 Then
            asc = button + 65
        Else
            asc = button - 26 + 48
        End If
        Return asc
    End Function
    ''' <summary>
    ''' Loads just the current row of buttons
    ''' </summary>
    Public Sub LoadCurrentButtonSet(layer As Byte)
        For i = 0 To 7
            frmMain.lblAlpha.Text = Chr(AscfromButt(iCurrentAlpha)).ToString
            Dim s As String
            Dim f As String = strButtonFilePath(i, iCurrentAlpha, layer)

            strVisibleButtons(i) = f
            s = strButtonCaptions(i, iCurrentAlpha, 1)
            If s <> "" Then
                lblDest(i).Text = s
                If My.Computer.FileSystem.DirectoryExists(f) Then
                    If CtrlDown Then
                        lblDest(i).ForeColor = Color.Red
                    ElseIf ShiftDown Then
                        lblDest(i).ForeColor = Color.Blue

                    Else
                        lblDest(i).ForeColor = Color.Black
                    End If

                Else
                    lblDest(i).ForeColor = Color.Gray

                End If

            Else
                lblDest(i).Text = "ABCDEFGH"(i)

            End If
        Next
    End Sub
    ''' <summary>
    ''' This is for when we hold down CTRL or SHIFT and the appearance of the buttons schanges. 
    ''' </summary>
    Public Sub UpdateButtonAppearance()
        frmMain.lblAlpha.Text = buttons.CurrentLetter
        For i = 0 To 7
            frmMain.ToolTip1.SetToolTip(btnDest(i), buttons.CurrentRow.Row(i).Path)
        Next
        For i = 0 To 7
            frmMain.lblAlpha.Text = Chr(AscfromButt(iCurrentAlpha)).ToString
            Dim s As String
            Dim f As String = strButtonFilePath(i, iCurrentAlpha, 1)
            frmMain.ToolTip1.SetToolTip(btnDest(i), f)

            strVisibleButtons(i) = f
            s = strButtonCaptions(i, iCurrentAlpha, 1)
            If s <> "" Then
                lblDest(i).Text = s
                If My.Computer.FileSystem.DirectoryExists(f) Then
                    If frmMain.NavigateMoveState.State = StateHandler.StateOptions.Move Xor CtrlDown Then
                        lblDest(i).ForeColor = Color.Red
                    ElseIf ShiftDown Then
                        lblDest(i).ForeColor = Color.Blue
                    Else
                        lblDest(i).ForeColor = Color.Black
                    End If

                Else
                    lblDest(i).ForeColor = Color.Gray

                End If

            Else
                lblDest(i).Text = "ABCDEFGH"(i)

            End If
        Next
    End Sub
    Public Sub InitialiseButtons()
        Dim alph As String = "ABCDEFGH"
        For i As Byte = 0 To 7
            With btnDest(i)
                .Text = buttons.CurrentRow.Row(i).FaceText
                '.Text = "F" & Str(i + 5)

                AddHandler .Click, AddressOf ButtonClick
                '  AddHandler .MouseEnter, AddressOf ButtonMouse
            End With

            lblDest(i).Text = buttons.CurrentRow.Row(i).Label
        Next
    End Sub
    Public Sub InitialiseButtons(row As ButtonRow)
        For i As Byte = 0 To 7
            With btnDest(i)
                .Text = "F" & Str(i + 5)

                AddHandler .Click, AddressOf ButtonClick
            End With

            lblDest(i).Text = row.Row(i).Label
        Next
    End Sub
    Private Sub ButtonClick(sender As Object, e As MouseEventArgs)
        'If e.Button = MouseButtons.Right Then
        FolderSelect.Show()
        'End If
    End Sub

    Public Sub NewButtonList()
        ClearCurrentButtons()
        SaveButtonlist()

    End Sub
    Public Sub ClearCurrentButtons()
        Array.Clear(strButtonFilePath, 0, strButtonFilePath.Length)
        Array.Clear(strButtonCaptions, 0, strButtonCaptions.Length)
        Array.Clear(strVisibleButtons, 0, strVisibleButtons.Length)
        For i = 0 To 7
            lblDest(i).Text = ""

        Next
    End Sub
    Public Sub KeyAssignmentsRestore(Optional filename As String = "")

        Dim intIndex, intLetter As Integer
        Dim path As String
        If filename = "" Then
            path = frmMain.LoadButtonList()
        Else
            path = filename
        End If
        If path = "" Then Exit Sub
        ClearCurrentButtons()
        'Get the file path
        Dim fs As New StreamReader(New FileStream(path, FileMode.OpenOrCreate, FileAccess.Read))
        Dim s As String
        Dim subs As String()

        Do While fs.Peek <> -1
            s = fs.ReadLine
            subs = s.Split("|")

            If subs.Length <> 4 Then
                'MsgBox("Not a button file")
                Exit Sub
            Else
                intIndex = Val(subs(0))
                intLetter = Val(subs(1))
                strButtonFilePath(intIndex, intLetter, 1) = (subs(2))
                strButtonCaptions(intIndex, intLetter, 1) = (subs(3))
            End If
        Loop

        UpdateButtonAppearance()

        fs.Close()
        Exit Sub


    End Sub
    Public Sub KeyAssignmentsStore(path As String)
        Dim intLoop As Integer
        Dim iLetter As Integer
        Dim Okay As Boolean
        Dim strEncrypted As String
        If path = "" Then
            Okay = False
        Else
            Dim f As New FileInfo(path)
            If f.Exists Then Okay = True
        End If
        If Not Okay Then

            With frmMain.SaveFileDialog1
                .DefaultExt = "msb"
                .Filter = "Metavisua button files|*.msb|All files|*.*"
                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    path = .FileName
                Else
                    Exit Sub

                End If
            End With
        End If

        Dim fs As New StreamWriter(New FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
            For iLetter = 0 To iAlphaCount - 1
                For intLoop = 0 To 7

                    If strButtonFilePath(intLoop, iLetter, 1) <> "" Then
                        strEncrypted = intLoop & "|" & iLetter & "|" & strButtonFilePath(intLoop, iLetter, 1) & "|" & strButtonCaptions(intLoop, iLetter, 1)
                        fs.WriteLine(strEncrypted)
                    End If
                Next
            Next
            strButtonFile = path
            fs.Close()


        PreferencesSave()

    End Sub

    Private Sub SaveASButton(path As String)
        Throw New NotImplementedException()
    End Sub

    Public Sub SetDirectory(strFilePath As String, KeyCode As Integer)
        'If it's already been assigned, then take it out of the list
        Dim Index As Integer
        Index = KeyCode - Keys.F4

        'Assign that path to the array
        strButtonFilePath(Index, iCurrentAlpha, 1) = strFilePath

    End Sub
End Module
