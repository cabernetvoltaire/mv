Module Shortcuts
    Public Class ShortcutHandler
        Public Sub New()

        End Sub
        ''' <summary>
        ''' Creates a shortcut to sTargetPath, places it in sSHortCutPath, with a short name of sShortcutName
        ''' </summary>
        ''' <param name="sTargetPath"></param>
        ''' <param name="sShortCutPath"></param>
        ''' <param name="sShortCutName"></param>
        Public Sub Create_ShortCut(ByVal sTargetPath As String, sShortCutPath As String, ByVal sShortCutName As String)


            ' Requires reference to Windows Script Host Object Model

            Dim oShell As IWshRuntimeLibrary.WshShell
            Dim oShortcut As IWshRuntimeLibrary.WshShortcut
            Dim sName As String
            oShell = New IWshRuntimeLibrary.WshShell
            Dim f As IO.FileInfo
            f = New IO.FileInfo(sShortCutPath)
            If f.Extension <> "lnk" Then
                sName = sShortCutPath & "\" & sShortCutName & ".lnk"
            Else
                sName = sShortCutPath

            End If
            If IO.File.Exists(sName) Then IO.File.Delete(sName)

            oShortcut = oShell.CreateShortcut(sName)

            With oShortcut
                .TargetPath = sTargetPath

                .Save()
            End With

            oShortcut = Nothing
            oShell = Nothing
        End Sub

        Public Sub Assign_ShortCutPath(ByVal sTargetPath As String, sShortCutPath As String, sShortCutName As String, oShortcut As IWshRuntimeLibrary.WshShortcut)

            ' Requires reference to Windows Script Host Object Model
            Dim sName As String
            sName = sShortCutPath & "\" & sShortCutName & ".lnk"

            With oShortcut
                .TargetPath = sTargetPath
                .Save()
            End With

            oShortcut = Nothing
        End Sub
    End Class




    Public Sub ReAssign_ShortCutPath(ByVal sTargetPath As String, sShortCutPath As String, oShortcut As IWshRuntimeLibrary.WshShortcut)

        ' Requires reference to Windows Script Host Object Model


        With oShortcut

            .TargetPath = sTargetPath
            .Save
        End With

        oShortcut = Nothing
    End Sub


    Function GetTargetPath(ByVal FileName As String)
        '    If RightString(FileName, ".") <> "lnk" Then
        '        sReport "Not a shortcut"
        'Exit Function
        '    End If

        '    Dim Obj As Object

        'Set Obj = CreateObject("WScript.Shell")


        'Dim Shortcut As Object

        'Set Shortcut = Obj.CreateShortcut(FileName)

        'GetTargetPath = Shortcut.TargetPath



        '    Shortcut.Save


    End Function

    Public Function ShortcutFolderCopy(fol As String, sDest As String)
        '    Dim sName As String
        '    Dim DestFol As folder
        '    sName = fol.ShortName
        'Set DestFol = fs2.CreateFolder(sDest + "\" + sName)
        'Dim f As file
        '    For Each f In fol.Files

        '        FileMoveToFolder f, DestFol, FOF_SILENT

        'Next


    End Function

End Module
