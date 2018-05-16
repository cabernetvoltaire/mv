
Public Class Grouper
    Private mPath As String
    Private mList As List(Of String)
    Public Sub New(Path As String)
        mPath = Path
        mList = IO.Directory.GetFiles(mPath).ToList
    End Sub
    Public Property InputList() As List(Of String)
        Get
            Return mList
        End Get
        Set(ByVal value As List(Of String))

            mList = value
        End Set
    End Property

    Private mSublists As List(Of List(Of String))
    Public ReadOnly Property Sublists() As List(Of List(Of String))
        Get
            Return GroupByName()
        End Get
    End Property
    Private Function GroupByName() As List(Of List(Of String))
        Dim mSublists As New List(Of List(Of String))
        Dim i As Int16
        While i <= mList.Count - 1
            Dim s As String = mList(i)
            Dim n = Len(s)
            While n > 1
                Dim trylist As List(Of String)
                Dim c As String = Left(s, n)
                trylist = mList.FindAll(Function(x) x.Contains(c))

                If trylist.Count > 1 And trylist.Count < mList.Count Then
                    mSublists.Add(trylist)
                    For Each j In trylist
                        mList.Remove(j)
                    Next
                    'trylist.Clear()
                Else
                    n = n - 1
                End If
            End While
            i += 1
        End While
        Return mSublists
    End Function


End Class
