
Public Class ButtonSet
    Public WithEvents CurrentSet As New List(Of ButtonRow)
    Public Event LetterChanged(l As Keys)
    Private alph As String = "ABCEDFGHIJKLMNOPQRSTUVWXYZ0123456789"

    Private mCurrentRow As New ButtonRow
    Public Property CurrentRow() As ButtonRow
        Get
            Return mCurrentRow
        End Get
        Set(ByVal value As ButtonRow)
            If Not mCurrentRow.Equals(value) Then
                mCurrentRow = value
            End If
        End Set
    End Property
    Private mCurrentLetter As Keys
    Public Property CurrentLetter() As Keys
        Get
            Return mCurrentLetter
        End Get
        Set(ByVal value As Keys)
            Dim b = mCurrentLetter
            If b <> value Then

                Dim c = InStr(alph, Chr(value))
                If c = 0 Then
                Else
                    mCurrentLetter = value
                    mCurrentRow = CurrentSet(c - 1)
                    RaiseEvent LetterChanged(value)
                End If
            Else
            End If

        End Set
    End Property
    Public Sub New()
        Dim Rows(Len(alph)) As ButtonRow
        For i = 0 To Len(alph) - 1
            Rows(i) = New ButtonRow
            Rows(i).Letter = CType(Asc(alph(i)), Keys)
            CurrentSet.Add(Rows(i))
        Next
        CurrentRow = CurrentSet(0)
    End Sub

End Class

