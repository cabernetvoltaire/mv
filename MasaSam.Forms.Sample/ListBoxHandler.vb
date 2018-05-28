Public Class ListBoxHandler
    Private mListBox As ListBox
    Public Event ListChanged()
    Public Property Listbox() As ListBox
        Get
            Return mListBox
        End Get
        Set(ByVal value As ListBox)
            mListBox = value
        End Set
    End Property

    Private mList As List(Of String)
    Public Property List() As List(Of String)
        Get
            Return mList
        End Get
        Set(ByVal value As List(Of String))
            Dim b As List(Of String)

            mList = value
            If b.Equals(mList) Then RaiseEvent ListChanged()
        End Set
    End Property

    Private Sub PopulateList()
        mListBox.Items.Clear()

        For Each s In mList
            mListBox.Items.Add(s)
        Next
    End Sub

End Class
