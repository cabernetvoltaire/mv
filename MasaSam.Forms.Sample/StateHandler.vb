Public Class StateHandler
    Public Enum StateOptions As Byte
        Navigate
        Move
        MoveLeavingLink
        Copy
        CopyLink

    End Enum
    Public Event StateChanged(sender As Object, e As EventArgs)
    Public mColour() As Color = ({Color.Aqua, Color.Orange, Color.LightPink, Color.LightGreen, Color.AliceBlue})

    Private mDescription = {"Navigate Mode", "Move Mode", "Move (Leave link)", "Copy", "Copy Link"}
    Private mInstructions = {"Function keys navigate to folder." & vbCrLf & "[SHIFT] + Fn moves file. " & vbCrLf & "[CTRL] + [SHIFT] +Fn moves folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key",
        "Function keys move file to folder." & vbCrLf & "[SHIFT] + Fn moves current folder to folder. " & vbCrLf & "[CTRL] + [SHIFT] +Fn navigates to folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key",
        "Function keys navigate to folder." & vbCrLf & "[SHIFT] + Fn moves file, leaving link. " & vbCrLf & "[CTRL] + [SHIFT] +Fn moves folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key",
        "Function keys navigate to folder." & vbCrLf & "[SHIFT] + Fn creates a copy in the destination" & vbCrLf & "[CTRL] + [SHIFT] +Fn does same for folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key",
        "Function keys navigate to folder." & vbCrLf & "[SHIFT] + Fn creates a link in the destination. " & vbCrLf & "[CTRL] + [SHIFT] +Fn does same for folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key"
    }
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
    Public ReadOnly Property Instructions() As String
        Get
            Return mInstructions(mState)
        End Get
    End Property




    Public Sub New()
        Me.State = StateOptions.Navigate
    End Sub

    Public Property State() As Byte
        Get
            Return mState
        End Get
        Set(ByVal value As Byte)
            mState = value
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

