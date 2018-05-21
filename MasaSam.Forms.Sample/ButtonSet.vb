
Public Class ButtonSet
    Public WithEvents CurrentSet As List(Of ButtonRow)

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
            mCurrentLetter = value
        End Set
    End Property
    Public Sub New()
        Dim alph As String = "ABCEDFGHIJKLMNOPQRSTUVWXYZ0123456789"
        Dim Rows(Len(alph)) As ButtonRow
        For i = 0 To Len(alph) - 1
            Rows(i).Letter = alph(i)
            CurrentSet.Add(Rows(i))
        Next
    End Sub

End Class

