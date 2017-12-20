Public Class Duplicates2
    Private newFlist As List(Of String)
    Private uniques(2000) As String
    Private duplicates(2000) As String
    Private Sub AnalyzeDuplicates()
        Dim lastlength As Long = 0
        Dim lastinfo As IO.FileInfo = Nothing
        Dim finfo3 As IO.FileInfo = Nothing
        Dim blnSame As Boolean = False
        Dim i As Long = 0
        For Each file In newFlist
            finfo3 = New IO.FileInfo(file) 'file is current file
            Dim currlength As Long = finfo3.Length
            If currlength = lastlength Then 'Last and current are duplicates
                If lastinfo IsNot Nothing Then 'for the first lopp

                    If Not blnSame Then 'Last was first different so add it
                        uniques(i) = lastinfo.FullName
                        'The duplicates of ith unique are kept in ith duplicate string
                        duplicates(i) = finfo3.FullName
                        blnSame = True
                        i += 1
                    Else
                        duplicates(i) = duplicates(i) + "|" + finfo3.FullName 'concatenate all duplicates
                    End If

                End If
            Else
                blnSame = False
            End If
            lastinfo = finfo3
            lastlength = lastinfo.Length
        Next
    End Sub
    Public ReadOnly Property Keeps As String()
        Get
            Return uniques
        End Get

    End Property
    Public ReadOnly Property Dupes() As String()
        Get
            Return duplicates
        End Get
    End Property


    Public Property Flist() As List(Of String)
        Get
            Return newFlist
        End Get
        Set(ByVal value As List(Of String))
            newFlist = value
        End Set
    End Property
    Private Sub FillListBox(lbx As ListBox, ByVal s() As String)
        For i = 0 To s.Length - 1
            If s(i) IsNot Nothing Then
                lbx.Items.Add(s(i))
            End If
        Next
    End Sub
    Private Sub Duplicates2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' FillShowbox(lbxsorted, CurrentFilterState, Flist)
        Flist = SetPlayOrder(PlayOrder.Length, Flist)
        AnalyzeDuplicates()
        FillListBox(lbxunique, uniques)
        FillListBox(lbxduplicates, duplicates)
    End Sub


End Class