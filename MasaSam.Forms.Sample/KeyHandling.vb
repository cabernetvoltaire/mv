Public Module KeyHandling

    Public KeyFullscreen As Integer = Keys.Scroll
    Public KeyRandomize As Integer = Keys.Pause
    Public KeyTraverseTreeBack = Keys.Subtract
    Public KeyJumpAutoT = Keys.Divide
    Public KeyBackUndo = Keys.Back
    Public KeyMuteToggle = Keys.Decimal
    Public KeyReStartSS = Keys.Space
    Public KeyNextFile = Keys.PageDown
    Public KeyLoopToggle = Keys.Insert
    Public KeyFolderJump = Keys.Multiply
    Public KeyZoomIn = Keys.OemPeriod

    'Development: Put the rest in here
    Public KeyTrueSize = Keys.Scroll
    Public KeyCycleScope = Keys.Oem7
    Public KeyCycleRandom = Keys.Pause
    Public KeyCycleMode = Keys.OemOpenBrackets
    Public KeyRotate = Keys.OemCloseBrackets
    Public KeyCopyToggle = Keys.OemSemicolon
    Public KeyFilter = Keys.OemQuestion
    Public KeyScrollPic = Keys.OemMinus
    Public KeyTraverseTree = Keys.Add
    Public KeyPreviousFile = Keys.PageUp
    Public KeyToggleSpeed = Keys.NumPad0
    Public KeySpeed1 = Keys.NumPad1
    Public KeySpeed2 = Keys.NumPad2
    Public KeySpeed3 = Keys.NumPad3
    Public KeyMarkPoint = Keys.NumPad4
    Public KeySmallJumpDown = Keys.NumPad5
    Public KeySmallJumpUp = Keys.NumPad6
    Public KeyJumpToPoint = Keys.NumPad7
    Public KeyBigJumpBack = Keys.NumPad8
    Public KeyBigJumpOn = Keys.NumPad9
    Public KeySpeed1Alt = Keys.Z
    Public KeySpeed2Alt = Keys.X
    Public KeySpeed3Alt = Keys.C
    Public KeyMarkPointAlt = Keys.A
    Public KeySmallJumpDownAlt = Keys.S
    Public KeySmallJumpUpAlt = Keys.D
    Public KeyJumpToPointAlt = Keys.Q
    Public KeyBigJumpBackAlt = Keys.W
    Public KeyBigJumpOnAlt = Keys.E
    Public KeyZoomOut = Keys.Oemcomma
    Public KeyToggleThumbs = Keys.Oem3
    Public KeyToggleButtons = Keys.OemOpenBrackets
    Public KeyEscape = Keys.Escape
    Public KeySelect = Keys.F3
    Public KeyDelete = Keys.Delete

    Public Enum FilterState
        All
        Piconly
        Vidonly
        PicVid
        LinkOnly
        NoPicVid
    End Enum
    Public Filterstates As String() = {"All", "Picture only", "Videos only", "Pictures and videos", "Links only", "Not pictures or videos"}
    Public FilterColours As Color() = {Color.LightGray, Color.LightPink, Color.LightSeaGreen, Color.LightSteelBlue, Color.Lime, Color.LightCyan}
    Public CurrentFilterState As Integer = FilterState.All
    'Public CurrentFolder As New IO.DirectoryInfo(C:\)
    Public Sub ControlSetFocus(control As Control)
        ' Set focus to the control, if it can receive focus.
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

    Public Sub GiveKey(K As Keys)
        Select Case K
            Case Keys.Scroll
                MsgBox(KeyFullscreen)
            Case Keys.Pause
                MsgBox(KeyRandomize)
            Case Keys.Subtract
                MsgBox(KeyTraverseTreeBack)
            Case Keys.Divide
                'MsgBox(KeyJumpAutoT)
            Case Keys.Back
                MsgBox(KeyBackUndo)
            Case Keys.Decimal
                'MsgBox(KeyMuteToggle)
            Case Keys.Space
                MsgBox(KeyReStartSS)
            Case Keys.PageDown
                MsgBox(KeyNextFile)
            Case Keys.Insert
                MsgBox(KeyLoopToggle)
            Case Keys.Multiply
                MsgBox(KeyFolderJump)

            Case Keys.OemPeriod
                MsgBox(KeyZoomIn)
            Case Keys.Scroll
                MsgBox(KeyTrueSize)
            Case Keys.Oemtilde
                MsgBox(KeyCycleScope)
            Case Keys.Pause
                MsgBox(KeyCycleRandom)
            Case Keys.OemOpenBrackets
                MsgBox(KeyCycleMode)
            Case Keys.OemCloseBrackets
                MsgBox(KeyRotate)
            Case Keys.OemSemicolon
                MsgBox(KeyCopyToggle)
            Case Keys.OemQuestion
'                MsgBox(KeyFilter)
            Case Keys.OemMinus
                MsgBox(KeyScrollPic)
            Case Keys.Add
                MsgBox(KeyTraverseTree)
            Case Keys.PageUp
                MsgBox(KeyPreviousFile)
            Case Keys.NumPad0
                'MsgBox(KeyToggleSpeed)
            Case Keys.NumPad1
                'MsgBox(KeySpeed1)
            Case Keys.NumPad2
                'MsgBox(KeySpeed2)
            Case Keys.NumPad3
                'MsgBox(KeySpeed3)
            Case Keys.NumPad4
                MsgBox(KeyMarkPoint)
            Case Keys.NumPad5
                MsgBox(KeySmallJumpDown)
            Case Keys.NumPad6
                MsgBox(KeySmallJumpUp)
            Case Keys.NumPad7
                MsgBox(KeyJumpToPoint)
            Case Keys.NumPad8
                MsgBox(KeyBigJumpBack)
            Case Keys.NumPad9
                'MsgBox(KeyBigJumpOn)
            Case Keys.Z
                MsgBox(KeySpeed1Alt)
            Case Keys.X
                MsgBox(KeySpeed2Alt)
            Case Keys.C
                MsgBox(KeySpeed3Alt)
            Case Keys.A
                MsgBox(KeyMarkPointAlt)
            Case Keys.S
                MsgBox(KeySmallJumpDownAlt)
            Case Keys.D
                MsgBox(KeySmallJumpUpAlt)
            Case Keys.Q
                MsgBox(KeyJumpToPointAlt)
            Case Keys.W
                MsgBox(KeyBigJumpBackAlt)
            Case Keys.E
                MsgBox(KeyBigJumpOnAlt)
            Case Keys.Oemcomma
                MsgBox(KeyZoomOut)
            Case Keys.Oem3
                MsgBox(KeyToggleThumbs)
            Case Else
                '       MsgBox(K.KeyCode)
        End Select
    End Sub
End Module

