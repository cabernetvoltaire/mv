Public Class Duplicates2
    Private newFlist As List(Of String)
    Private uniques As New List(Of String)
    Private duplicates() As List(Of String)

    Private Sub AnalyzeDuplicates()
        Dim lastlength As Long = 0
        Dim lastinfo As IO.FileInfo = Nothing
        Dim finfo3 As IO.FileInfo = Nothing
        Dim blnSame As Boolean = False
        Dim i As Long = 0
        For Each file In newFlist
            finfo3 = New IO.FileInfo(file)
            Dim currlength As Long = finfo3.Length
            If currlength = lastlength Then 'Last and current are duplicates
                If lastinfo IsNot Nothing Then 'for the first lopp

                    If Not blnSame Then 'Last was first different so add it
                        uniques.Add(lastinfo.FullName)
                        blnSame = True
                    End If
                    If blnSame Then
                        duplicates(i).Add(lastinfo.FullName)
                    End If
                    i += 1
                End If
            Else
                blnSame = False
            End If
            lastinfo = finfo3
            lastlength = lastinfo.Length
        Next
    End Sub


    Public Property Flist() As List(Of String)
        Get
            Return newFlist
        End Get
        Set(ByVal value As List(Of String))
            newFlist = value
        End Set
    End Property
    Private Sub FillListBox(lbx As ListBox, ByVal s() As String)
        For i = 0 To s.Length
            lbx.Items.Add(s(i))
        Next
    End Sub
    Private Sub Duplicates2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FillShowbox(lbxsorted, CurrentFilterState, Flist)
        AnalyzeDuplicates()
        FillShowbox(lbxunique, CurrentFilterState, uniques)
        FillShowbox(lbxduplicates, CurrentFilterState, duplicates(0))
    End Sub

    Private Sub lbxunique_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxunique.SelectedIndexChanged
        If lbxunique.SelectedIndex < duplicates.Length Then

            FillShowbox(lbxduplicates, CurrentFilterState, duplicates(lbxunique.SelectedIndex))
        End If

    End Sub
End Class