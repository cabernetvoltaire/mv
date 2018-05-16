Public Class StartPointHandler
    Public Enum StartTypes As Byte
        Beginning
        NearBeginning
        NearEnd
        Random
        ParticularPercentage
        ParticularAbsolute
    End Enum

    Public Event StateChanged(s As StartTypes)
    Public Event JumpKey()
    Private mOrder = {"Beginning", "Near Beginning", "Near End", "Random", "Particular(%)", "Particular(s)"}
    Private mDescList As New List(Of String)
    Public Sub New(Optional ByVal StartPercentage As Byte = 50, Optional ByVal StartAbsolute As Byte = 65)
        mState = StartTypes.Beginning
        mPercentage = StartPercentage
        mAbsolute = StartAbsolute
    End Sub
    Public ReadOnly Property Descriptions As List(Of String)
        Get
            For i = 0 To 5
                mDescList.Add(mOrder(i))
            Next
            Descriptions = mDescList
        End Get
    End Property
    Public ReadOnly Property Description() As String
        Get
            Return mOrder(mState)
        End Get
    End Property
    Private mDuration As Long
    Public Property Duration() As Long
        Get
            Return mDuration

        End Get
        Set(ByVal value As Long)
            mDuration = value
        End Set
    End Property
    Private mDistance As Long
    Public Property Distance() As Long
        Get
            Return mDistance
        End Get
        Set(ByVal value As Long)
            mDistance = value
        End Set
    End Property

    Private mAbsolute As Long
    Public Property Absolute() As Long
        Get
            Return mAbsolute
        End Get
        Set(ByVal value As Long)
            mAbsolute = value
        End Set
    End Property
    Private mPercentage As Byte
    Public Property Percentage() As Byte
        Get
            Return mPercentage
        End Get
        Set(ByVal value As Byte)
            mPercentage = value
        End Set
    End Property
    Private mStartPoint As Long
    Public Property StartPoint() As Long
        Get

            Return GetStartPoint()
        End Get
        Set(ByVal value As Long)
            mStartPoint = value
        End Set
    End Property
    Private mState As Byte
    Public Property State() As Byte
        Get
            Return mState
        End Get
        Set(ByVal value As Byte)
            mState = value
            mStartPoint = GetStartPoint()
            RaiseEvent StateChanged(mState)
        End Set
    End Property

    Private Function GetStartPoint()
        Select Case mState
            Case StartTypes.Beginning
                mStartPoint = 0
            Case StartTypes.NearBeginning
                mStartPoint = mDistance
            Case StartTypes.NearEnd
                mStartPoint = mDuration - mDistance
            Case StartTypes.ParticularAbsolute
                mStartPoint = Math.Min(mAbsolute, mDuration / 2)

            Case StartTypes.ParticularPercentage
                mStartPoint = mPercentage / 100 * mDuration
            Case StartTypes.Random
                mStartPoint = Rnd() * mDuration
        End Select
        Return mStartPoint
    End Function
End Class
