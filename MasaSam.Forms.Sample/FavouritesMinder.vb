Public Class FavouritesMinder
    Private mFavesList As New List(Of String)
    Public Property FavesList() As List(Of String)
        Get
            Return mFavesList
        End Get
        Set(ByVal value As List(Of String))
            mFavesList = value
        End Set
    End Property

    Private mDestPath As String
    Public Property DestinationPath() As String
        Get
            Return mDestPath
        End Get
        Set(ByVal value As String)
            mDestPath = value
        End Set

    End Property

    Public Sub New(path As String)
        Dim faves As New IO.DirectoryInfo(path)
        If Not faves.Exists Then faves = IO.Directory.CreateDirectory(Environment.SpecialFolder.MyPictures & "\Favourites")
        For Each f In faves.GetFiles("*.lnk", IO.SearchOption.AllDirectories)

            mFavesList.Add(f.FullName)
        Next
    End Sub

    Public Sub CheckFile(f As IO.FileInfo)
        If f.Extension = ".lnk" Then
        Else
            Check(f, mDestPath)

        End If
    End Sub
    Public Sub CheckFiles(f As List(Of String))
        For Each m In f
            Dim k As New IO.FileInfo(m)
            Check(k, mDestPath)
        Next
    End Sub

    Public Sub Check(f As IO.FileInfo, destinationpath As String)
        Dim m As String = FavesList.Find(Function(x) x.Contains(f.Name))
        Dim bk As Long = 0
        If m <> "" Then
            Dim minfo As New IO.FileInfo(m)

            If InStr(m, "%") <> 0 Then

                Dim s As String()
                s = m.Split("%")
                bk = Val(s(1))
            Else
                bk = 0
            End If
            Dim sch As New ShortcutHandler(destinationpath, minfo.Directory.FullName, f.Name)
            sch.Create_ShortCut(bk)

        End If
    End Sub
End Class
