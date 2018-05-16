Public Class FilterHandler
    Public Enum FilterState As Byte
        All
        Piconly
        Vidonly
        PicVid
        LinkOnly
        NoPicVid
    End Enum

    Public Event StateChanged(s As FilterState)
    Private mColour = {Color.MintCream, Color.LemonChiffon, Color.LightPink, Color.LightSeaGreen, Color.LightBlue, Color.PaleTurquoise}
    Private mDescription = {"All files", "Only pictures", "Only videos", "Only pictures and videos", "Only links", "No pictures or videos"}
    Public ReadOnly Property Colour() As Color
        Get
            Return mColour(mState)
        End Get
    End Property
    Private mDescList As New List(Of String)

    Public ReadOnly Property Descriptions As List(Of String)
        Get
            For i = 0 To 5
                mDescList.Add(mDescription(i))
            Next
            Descriptions = mDescList
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return mDescription(mState)
        End Get
    End Property
    Private mState As Byte
    Public Property State() As Byte
        Get
            Return mState
        End Get
        Set(ByVal value As Byte)
            mState = value

            RaiseEvent StateChanged(mState)
        End Set
    End Property
    Public Sub IncrementState()
        Me.State = (Me.State + 1) Mod (FilterState.NoPicVid)
    End Sub
End Class
