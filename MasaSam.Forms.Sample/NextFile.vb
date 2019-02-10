Public Class NextFile
    Private mListCount As Integer
    Private mListbox As New ListBox
    Public Property Listbox() As ListBox
        Get
            Return mListbox
        End Get
        Set(ByVal value As ListBox)
            mListbox = value
            mListCount = mListbox.Items.Count
        End Set
    End Property

    Public Property CurrentItem As String
    Public Property Forwards As Boolean = True
    Private mCurrentIndex As Integer
    Public Property CurrentIndex() As Integer
        Get
            Return mCurrentIndex
        End Get
        Set(ByVal value As Integer)
            mCurrentIndex = value
            If value >= 0 And value <= mListCount Then
                CurrentItem = mListbox.Items(mCurrentIndex)
            End If

        End Set
    End Property
    Private Property mNextItem As String
    Public Property NextItem As String
        Get
            If Randomised Then
                mNextItem = Listbox.Items(Int(Rnd() * (mListCount - 1)))
            Else
                If mListCount > 1 Then
                    If Forwards Then
                        mNextItem = Listbox.Items((mCurrentIndex + 1) Mod (mListCount - 1))
                    Else
                        If mCurrentIndex = 0 Then
                            mCurrentIndex = mListCount - 1
                        Else
                            mNextItem = Listbox.Items((mCurrentIndex - 1) Mod (mListCount - 1))
                        End If
                    End If
                Else
                    mNextItem = Listbox.Items(0)

                End If
            End If
            Return mNextItem
        End Get
        Set(value As String)

        End Set
    End Property

    Private Property mPreviousItem As String
    Public Property PreviousItem As String
        Get
            If Randomised Then
                mPreviousItem = Listbox.Items(Int(Rnd() * (mListCount - 1)))
            Else
                If mListCount > 1 Then
                    If Forwards Then
                        mPreviousItem = Listbox.Items((mCurrentIndex - 1) Mod (mListCount - 1))
                    Else
                        If mCurrentIndex = 0 Then
                            mCurrentIndex = mListCount - 1
                        Else
                            mPreviousItem = Listbox.Items((mCurrentIndex + 1) Mod (mListCount - 1))
                        End If
                    End If
                Else
                    mPreviousItem = Listbox.Items(0)

                End If
            End If
            Return mNextItem
        End Get
        Set(value As String)

        End Set
    End Property


    Public Property Randomised As Boolean


End Class
