
Public Class ButtonSet
    Public WithEvents CurrentSet As New List(Of ButtonRow)
    Public Event LetterChanged As EventHandler
    Private mCurrentRow As ButtonRow
    Public Property CurrentRow() As ButtonRow
        Get
            Return mCurrentRow
        End Get
        Set(ByVal value As ButtonRow)
            mCurrentRow = value
        End Set
    End Property
    Private mCurrentLetter As Char
    Public Property CurrentLetter() As Char
        Get
            Return mCurrentLetter
        End Get
        Set(ByVal value As Char)
            Dim b = mCurrentLetter
            mCurrentLetter = value
            mCurrentRow = CurrentSet(Asc(mCurrentLetter) - Asc("A"))
            If b <> mCurrentLetter Then RaiseEvent LetterChanged(Me, Nothing)
            'Change current row
        End Set
    End Property
    Public Sub New()
        Dim alph As String = "ABCEDFGHIJKLMNOPQRSTUVWXYZ0123456789"
        Dim Rows(Len(alph)) As ButtonRow
        For i = 0 To Len(alph) - 1
            Rows(i) = New ButtonRow
            Rows(i).Letter = alph(i)
            CurrentSet.Add(Rows(i))
        Next
        CurrentRow = CurrentSet(0)
    End Sub

End Class

