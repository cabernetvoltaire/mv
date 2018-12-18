﻿Public Class FilterHandler
    Public Enum FilterState As Byte
        All
        PicVid

        Piconly
        Vidonly
        LinkOnly
        NoPicVid
    End Enum

    Public Event StateChanged(sender As Object, e As EventArgs)
    Private mColour = {Color.MintCream, Color.LemonChiffon, Color.LightPink, Color.LightSeaGreen, Color.LightBlue, Color.PaleTurquoise}
    Private mDescription = {"All files", "Only pictures and videos", "Only pictures", "Only videos", "Only links", "No pictures or videos"}
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

    Public Sub New()
        Me.State = FilterState.All
    End Sub
    Public Property OldState As Byte
    Public Property State() As Byte
        Get
            Return mState
        End Get
        Set(ByVal value As Byte)
            If value = FilterState.LinkOnly Then OldState = mState
            mState = value

            RaiseEvent StateChanged(Me, New EventArgs)
        End Set
    End Property
    Public Sub IncrementState()
        Me.State = (Me.State + 1) Mod (FilterState.NoPicVid + 1)
    End Sub
End Class
