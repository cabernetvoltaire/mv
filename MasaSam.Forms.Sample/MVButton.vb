Public Class MVButton

    Private mPath As String
    Public Property Path() As String
        Get
            Return mPath
        End Get
        Set(ByVal value As String)
            mPath = value
        End Set
    End Property
    Private mFaceText As String
    Public Property FaceText() As String
        Get
            Return mFaceText
        End Get
        Set(ByVal value As String)
            mFaceText = value
        End Set
    End Property
    Private mlblText As String
    Public Property Label() As String
        Get
            Return mlblText
        End Get
        Set(ByVal value As String)
            mlblText = value
        End Set
    End Property
    Private mActive As Boolean
    Public Property Active() As Boolean
        Get
            Return mActive
        End Get
        Set(ByVal value As Boolean)
            mActive = value
        End Set
    End Property

    Private mColour As Color
    Public Property Colour() As Color
        Get
            Return mColour
        End Get
        Set(ByVal value As Color)
            mColour = value
        End Set
    End Property


End Class
