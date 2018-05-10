Public Module KeyHandling
    'Numpad options (not available on Laptop)
    Public KeyFullscreen = Keys.Scroll
    Public KeyRandomize = Keys.Pause
    Public KeyTraverseTreeBack = Keys.Subtract
    Public KeyTraverseTree = Keys.Add
    Public KeyJumpAutoT = Keys.Divide
    Public KeyMuteToggle = Keys.Decimal
    Public KeyNextFile = Keys.PageDown
    Public KeyPreviousFile = Keys.PageUp
    Public KeyZoomIn = Keys.OemPeriod
    Public KeyFolderJump = Keys.Multiply
    Public KeyCycleScope = Keys.Oem7
    Public KeyCycleMode = Keys.OemOpenBrackets
    Public KeyToggleSpeed = Keys.NumPad0
    Public KeySpeed1 = Keys.NumPad1
    Public KeySpeed2 = Keys.NumPad2
    Public KeySpeed3 = Keys.NumPad3
    Public KeySmallJumpDown = Keys.NumPad5
    Public KeySmallJumpUp = Keys.NumPad6
    Public KeyJumpToPoint = Keys.NumPad7
    Public KeyBigJumpBack = Keys.NumPad8
    Public KeyBigJumpOn = Keys.NumPad9
    Public KeyMarkPoint = Keys.NumPad4

    Public LKeyFullscreen = Keys.F + Keys.Alt
    Public LKeyRandomize = Keys.R + Keys.Alt
    Public LKeyTraverseTreeBack = Keys.Alt + Keys.OemMinus
    Public LKeyTraverseTree = Keys.Alt + Keys.Oemplus
    Public LKeyJumpAutoT = Keys.Alt + Keys.Divide
    Public LKeyMuteToggle = Keys.Alt + Keys.Decimal
    Public LKeyNextFile = Keys.Alt + Keys.N
    Public LKeyPreviousFile = Keys.Alt + Keys.P
    Public LKeyZoomIn = Keys.Alt + Keys.Z
    Public LKeyFolderJump = Keys.Alt + Keys.Multiply
    Public LKeyToggleSpeed = Keys.Alt + Keys.D0
    Public LKeySpeed1 = Keys.Alt + Keys.D1
    Public LKeySpeed2 = Keys.Alt + Keys.D2
    Public LKeySpeed3 = Keys.Alt + Keys.D3
    Public LKeySmallJumpDown = Keys.Alt + Keys.D5
    Public LKeySmallJumpUp = Keys.Alt + Keys.D6
    Public LKeyJumpToPoint = Keys.Alt + Keys.D7
    Public LKeyBigJumpBack = Keys.Alt + Keys.D8
    Public LKeyBigJumpOn = Keys.Alt + Keys.D9
    Public LKeyMarkPoint = Keys.Alt + Keys.D4


    'Common Keys
    Public KeyBackUndo = Keys.Back
    Public KeyReStartSS = Keys.Space
    Public KeyLoopToggle = Keys.Insert
    Public KeyTrueSize = Keys.Oemplus
    Public KeyCycleRandom = Keys.Pause
    Public KeyRotate = Keys.OemCloseBrackets
    Public KeyRotateBack = Keys.OemOpenBrackets
    Public KeyFilter = Keys.OemQuestion
    Public KeyAddFile = Keys.Oemtilde

    Public KeyZoomOut = Keys.Oemcomma
    Public KeyToggleThumbs = Keys.Oem3
    Public KeyToggleButtons = Keys.OemSemicolon
    Public KeyEscape = Keys.Escape
    Public KeySelect = Keys.F3
    Public KeyDelete = Keys.Delete
    Public KeyMoveToggle = Keys.Oem7


    Public Enum FilterState
        All
        Piconly
        Vidonly
        PicVid
        LinkOnly
        NoPicVid
    End Enum


    'Public CurrentFolder As New IO.DirectoryInfo(C:\)
    Public Sub ControlSetFocus(control As Control)
        ' Set focus to the control, if it can receive focus.
        Exit Sub
        If control.CanFocus Then
            control.Focus()
        End If
    End Sub

    Public Function GetMaxValue _
    (Of TEnum As {IComparable, IConvertible, IFormattable})() As TEnum

        Dim type = GetType(TEnum)

        If Not type.IsSubclassOf(GetType([Enum])) Then _
            Throw New InvalidCastException _
                ("Cannot cast '" & type.FullName & "' to System.Enum.")
        Return [Enum].ToObject(type, [Enum].GetValues(type) _
                        .Cast(Of Integer).Last)
    End Function

End Module

