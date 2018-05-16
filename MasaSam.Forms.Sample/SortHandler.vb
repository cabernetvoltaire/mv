﻿Public Class SortHandler
    Public Enum Order As Byte
        Original
        Random
        Name
        PathName
        DateTime
        Size
        Type
    End Enum

    Public Event StateChanged(s As Order)
    Private mOrder = {"Original", "Random", "Name", "Path Name", "Date/Time", "Size", "Type"}

    Private mDescList As New List(Of String)

    Public ReadOnly Property Descriptions As List(Of String)
        Get
            For i = 0 To 6
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
        Me.State = (Me.State + 1) Mod (Order.Type)
    End Sub

    Public Sub Toggle()
        If Me.State = Order.Original Then
            Me.State = Order.Random
        Else
            Me.State = Order.Original
        End If
    End Sub
End Class
