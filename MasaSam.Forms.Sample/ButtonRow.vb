Public Class ButtonRow
    Public Row(8) As MVButton
    Public Event CurrentChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Sub New()
        Dim S As String = "ABCDEFGH"
        For i = 0 To 7
            Row(i) = New MVButton
            Row(i).FaceText = "f" & Str(i + 5)
            Row(i).Label = S(i)
        Next

        'Current = False
        'Letter = "A"
    End Sub
    Private mCurrent As Boolean
        Public Property Current() As Boolean
            Get
                Return mCurrent
            End Get
            Set(ByVal value As Boolean)
                mCurrent = value
                RaiseEvent CurrentChanged(Me, New EventArgs)


            End Set
        End Property
        Private mLetter As Char
        Public Property Letter() As Char
            Get
                Return mLetter
            End Get
            Set(ByVal value As Char)

                mLetter = value
            '     Key = GetKey(value)
        End Set
        End Property

    Private Function GetKey(value As Char) As Integer
        Dim alpha As String = "ABCDEFGHIJKLMNOPWRSTUVWXYZ"
        Dim i As Byte = InStr(alpha, value)
        If i = 0 Then
            i = InStr("0123456789", value)
            Return i + Keys.D0
        Else
            Return i + Keys.A
        End If
        Throw New NotImplementedException()
    End Function

    Private mKey As Keys
    Public Property Key() As Integer
        Get
            Return mKey
        End Get
        Set(ByVal value As Integer)

            mKey = value
            mLetter = Chr(value)
        End Set
    End Property


    Public Sub InitialiseButtons()
            Dim alph As String = "ABCDEFGH"
            For i As Byte = 0 To 7
                Row(i).FaceText = "F" & Str(i + 5)
                Row(i).Label = alph(i)
            Next
        End Sub
    End Class
