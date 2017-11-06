﻿Imports System.IO

Module ButtonHandling
    Public strButtonFilePath(8, 26, 1) As String
    Public fButtonDests(8, 26, 1) As IO.DirectoryInfo
    Private iAlphaCount = 26
    Public strButtonCaptions(8, 26, 1)


    Public Sub ChangeButtonLetter(e As KeyEventArgs)
        frmMain.lblAlpha.Text = e.KeyCode.ToString
        iCurrentAlpha = e.KeyCode - Keys.A
        LoadCurrentButtonSet()

    End Sub

    Public Sub LoadCurrentButtonSet()
        For i = 0 To 7
            Dim s As String
            strVisibleButtons(i) = strButtonFilePath(i, iCurrentAlpha, 1)
            s = strButtonCaptions(i, iCurrentAlpha, 1)
            If s <> "" Then
                lblDest(i).Text = s

            Else
                lblDest(i).Text = "ABCDEFGH"(i)

            End If
        Next
    End Sub

    Public Sub InitialiseButtons()
        Dim alph As String = "ABCDEFGH"
        For i As Byte = 0 To 7
            btnDest(i).Text = "F" & Str(i + 5)
            lblDest(i).Text = alph(i)
        Next
    End Sub
    Private Sub AssignButton(Index As Integer, strDraggedPath As String)
        Dim strCaption As String

        'Assign dropped path to this button.
        strButtonFilePath(Index, iCurrentAlpha, 1) = strDraggedPath
        lblDest(Index).Enabled = IO.Directory.Exists(strDraggedPath)
        'Give it a caption.
        'Default Caption will be the folder name
        strCaption = Right(strButtonFilePath(Index, iCurrentAlpha, 1), "\")

        lblDest(Index).Text = strCaption
        strButtonCaptions(Index, iCurrentAlpha, 1) = strCaption
        'Assign the key to this.
        SetDirectory(strDraggedPath, Index)
        KeyAssignmentsStore(strButtonFile)

    End Sub
    Public Sub KeyAssignmentsRestore()

        Dim intIndex, intLetter As Integer
        Dim path As String
        path = frmMain.LoadButtonList()
        'Get the file path
        Dim fs As New StreamReader(New FileStream(path, FileMode.OpenOrCreate, FileAccess.Read))
        Dim s As String
        Do While fs.Peek <> -1
            s = fs.ReadLine
            intIndex = Val(s.Split("|")(0))
            intLetter = Val(s.Split("|")(1))
            strButtonFilePath(intIndex, intLetter, 1) = (s.Split("|")(2))
            strButtonCaptions(intIndex, intLetter, 1) = (s.Split("|")(3))
        Loop

        LoadCurrentButtonSet()

        fs.Close()
        Exit Sub


    End Sub
    Public Sub KeyAssignmentsStore(path As String)
        Dim intLoop As Integer
        Dim iLetter As Integer
        Dim strEncrypted As String


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