Public Class StateHandler
    Public Enum StateOptions As Byte
        Navigate
        Move
        MoveLeavingLink
        Copy
        CopyLink

    End Enum
    Public Event StateChanged(s As StateOptions)

    Private mState As String
    Public Property State() As String
        Get
            Return mState
        End Get
        Set(ByVal value As String)
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
            RaiseEvent StateChanged(mState)
        End Set
    End Property
    Public Sub IncrementState()
        Me.State = (Me.State + 1) Mod (StateOptions.CopyLink)
    End Sub
End Class
