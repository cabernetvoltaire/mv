﻿Module Shortcuts
    Public Class ShortcutHandler
        Public Sub New()
            Dim TargetPath As String = ""
            Dim ShortCutPath As String = ""
            Dim ShortCutName As String = ""

        End Sub
        Private sTargetPath As String
        Private sShortcutPath As String
        Private sShortcutName As String
        Public Property TargetPath() As String
            Get
                Return sTargetPath
            End Get
            Set(ByVal value As String)
                sTargetPath = value
            End Set
        End Property
        Public Property ShortcutPath() As String
            Get
                Return sShortcutPath
            End Get
            Set(ByVal value As String)
                sShortcutPath = value
            End Set
        End Property
        Public Property ShortcutName() As String
            Get
                Return sShortcutName
            End Get
            Set(ByVal value As String)
                sShortcutName = value
            End Set
        End Property
        ''' <summary>
        ''' Creates a shortcut to sTargetPath, places it in sSHortCutPath, with a short name of sShortcutName
        ''' </summary>

        Public Sub Create_ShortCut()


            ' Requires reference to Windows Script Host Object Model

            Dim oShell As IWshRuntimeLibrary.WshShell
            Dim oShortcut As IWshRuntimeLibrary.WshShortcut
            Dim sName As String
            oShell = New IWshRuntimeLibrary.WshShell
            Dim f As IO.FileInfo
            f = New IO.FileInfo(sShortcutPath)
            If f.Extension <> "lnk" Then
                sName = sShortcutPath & "\" & sShortcutName & ".lnk"
            Else
                sName = sShortcutPath

            End If
            If IO.File.Exists(sName) Then IO.File.Delete(sName)

            oShortcut = oShell.CreateShortcut(sName)

            With oShortcut
                Dim d As New IO.DirectoryInfo(sShortcutPath)
                If d.Exists Then
                Else
                    d.Create()
                End If

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
