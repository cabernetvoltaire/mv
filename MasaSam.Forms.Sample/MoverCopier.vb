Imports System.Threading
Imports System.IO
Public Class MoverCopier
    Public Event FolderMoved(Path As String)
    Public Event FileMoved(DirPath As String)
    Public t As Thread
    Public fm As New FavouritesMinder(FavesFolderPath)

    Private mMedia As MediaHandler
    Private mFolder As New IO.DirectoryInfo(mMedia.MediaDirectory)
    Public Property Media() As MediaHandler
        Get
            Return mMedia
        End Get
        Set(ByVal value As MediaHandler)
            mMedia = value
            mFolder = New DirectoryInfo(mMedia.MediaDirectory)
        End Set
    End Property

    ''' <summary>
    ''' Simply moves a folder to its parent folder
    ''' </summary>
    ''' <param name="folder"></param>
    Public Sub PromoteFolder(folder As DirectoryInfo)
        folder.MoveTo(folder.Parent.Parent.FullName & "\" & folder.Name)
        RaiseEvent FolderMoved(folder.FullName)
    End Sub
    Private Sub MoveFolder(strdest As String)
        'Starts a thread
        t = New Thread(New ThreadStart(Sub() FolderMover(mFolder.FullName, strdest, True)))
        t.IsBackground = True
        t.SetApartmentState(ApartmentState.STA)

        t.Start()
    End Sub
    Public Sub FolderMover(strDir As String, strDest As String, blnOverride As Boolean)
        Try
            If strDest = "" Then mFolder.Delete()
            With My.Computer.FileSystem
                Dim s As String = mFolder.Name
                Dim par As New DirectoryInfo(mFolder.Parent.FullName)
                Dim destdir = New DirectoryInfo(strDest)
                Select Case NavigateMoveState.State
                    Case StateHandler.StateOptions.Copy

                        .CopyDirectory(strDir, strDest & "\" & s, FileIO.UIOption.OnlyErrorDialogs)

                    Case StateHandler.StateOptions.Move

                        Dim flist As New List(Of String)
                        GetFiles(mFolder, flist)
                        'Check the favourites
                        fm.DestinationPath = strDest
                        fm.CheckFiles(flist)
                        .MoveDirectory(strDir, strDest & "\" & s, FileIO.UIOption.OnlyErrorDialogs)

                    Case StateHandler.StateOptions.MoveLeavingLink

                        'Create link directory?
                        .MoveDirectory(strDir, strDest & "\" & s, FileIO.UIOption.OnlyErrorDialogs)

                    Case StateHandler.StateOptions.CopyLink
                        'Creat link directory?

                End Select
                UpdateButton(strDir, strDest & "\" & s) 'todo doesnt handle sub-tree
            End With
        Catch ex As Exception
            '            MsgBox(ex.Message) '
        End Try

    End Sub

    Private Sub MoveFolder(source As DirectoryInfo, destinationpath As String)
        Dim destdir As New DirectoryInfo(destinationpath)
        Dim filelist As New List(Of String)
        GetFiles(source, filelist)

        If destdir.Exists = False Then My.Computer.FileSystem.CreateDirectory(destinationpath)
        MoveFiles(filelist, destinationpath)

    End Sub

    ''' <summary>
    ''' Converts a dir into a list of filepaths
    ''' </summary>
    ''' <param name="dir"></param>
    ''' <param name="flist"></param>
    Private Sub GetFiles(dir As DirectoryInfo, flist As List(Of String))
        For Each m In dir.EnumerateFiles()
            flist.Add(m.FullName)
        Next
        For Each x In dir.EnumerateDirectories
            GetFiles(x, flist)
        Next
    End Sub

End Class
