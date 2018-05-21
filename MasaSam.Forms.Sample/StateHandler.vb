Public Class StateHandler
    Public Enum StateOptions As Byte
        Navigate
        Move
        MoveLeavingLink
        Copy
        CopyLink

    End Enum
    Public Event StateChanged(sender As Object, e As EventArgs)
    Public mColour() As Color = ({Color.Aqua, Color.Orange, Color.Purple, Color.Gray, Color.AliceBlue})

    Private mDescription = {"Navigate Mode", "Move Mode", "Move (Leave link)", "Copy", "Copy Link"}
    Public ReadOnly Property Colour() As Color
        Get
            Return mColour(mState)
        End Get
    End Property
    Public ReadOnly Property Description() As String
        Get
            Return mDescription(mState)
        End Get
    End Property
    Private mState As Byte

    Public Sub New()
        Me.State = StateOptions.Navigate
    End Sub

    Public Property State() As Byte
        Get
            Return mState
        End Get
        Set(ByVal value As Byte)
            mState = value
            Select Case mState
                Case StateOptions.Navigate
                    blnMoveMode = False
                    blnCopyMode = False
                Case StateOptions.Move
                    blnMoveMode = True
                    blnCopyMode = False
                Case StateOptions.MoveLeavingLink
                Case StateOptions.Copy
                    blnMoveMode = True
                    blnCopyMode = True
                Case StateOptions.CopyLink

            End Select
            RaiseEvent StateChanged(Me, New EventArgs)
        End Set
    End Property
    Public Sub ToggleState()
        Select Case Me.State
            Case StateOptions.Navigate
                Me.State = StateOptions.Move
            Case Else
                Me.State = StateOptions.Navigate
        End Select

    End Sub
    Public Sub IncrementState()
        Me.State = (Me.State + 1) Mod (StateOptions.CopyLink)
    End Sub
End Class

