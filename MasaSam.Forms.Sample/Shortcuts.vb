﻿Imports System.IO
Imports IWshRuntimeLibrary
Module Shortcuts

    Public Class ShortcutHandler
        Public Sub New()
            Dim TargetPath As String = ""
            Dim ShortCutPath As String = ""
            Dim ShortCutName As String = ""

        End Sub
        Public Sub New(TargetPath, ShortCutPath, shortCutName)
            sTargetPath = TargetPath
            sShortcutPath = ShortCutPath
            sShortcutName = shortCutName
        End Sub
        Private oShell As WshShell
        Private oShortcut As WshShortcut
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
        ''' Creates a shortcut to TargetPath, places it in ShortCutPath, with a short name of ShortcutName
        ''' </summary>

        Public Sub Create_ShortCut()

            Dim sName As String
            oShell = New WshShell
            Dim f As IO.FileInfo
            f = New IO.FileInfo(sShortcutPath)
            If f.Extension <> ".lnk" Then
                sName = sShortcutPath & "\" & sShortcutName & ".lnk"
            Else
                sName = sShortcutPath

            End If
            Dim exf As New IO.FileInfo(sName)
            If exf.Exists Then exf.Delete()

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

        Public Sub Assign_ShortCutPath()

            ' Requires reference to Windows Script Host Object Model
            Dim sName As String
            sName = sShortcutPath & "\" & sShortcutName & ".lnk"

            With oShortcut
                .TargetPath = sTargetPath
                .Save()
            End With

            oShortcut = Nothing
        End Sub
    End Class




    Public Sub ReAssign_ShortCutPath(ByVal sTargetPath As String, sShortCutPath As String)
        Dim d As New DirectoryInfo(sShortCutPath)
        CreateLink(sTargetPath, d.Parent.FullName)
    End Sub
    Public Function LinkTargetExists(Linkfile As String) As Boolean
        Dim f As String
        f = LinkTarget(Linkfile)
        If f = "" Then
            Return False
            Exit Function
        End If
        Dim Finfo = New FileInfo(f)
        If Finfo.Exists Then
            Return True
        Else
            Return False
        End If

    End Function


End Module
