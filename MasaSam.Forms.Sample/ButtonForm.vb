Public Class ButtonForm
    Public WithEvents buttons As New ButtonSet

    Public sbtns As Button()
    Public lbls As Label()

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub ButtonForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sbtns = {Me.Button4, Me.Button2, Me.btn1, Me.Button1, Me.Button9, Me.Button7, Me.Button5, Me.Button6}
        lbls = {Me.Label4, Me.Label2, Me.lbl1, Me.Label1, Me.Label9, Me.Label7, Me.Label5, Me.Label6}
        TranscribeButtons(buttons.CurrentRow)
    End Sub
    Private Function TranscribeButtons(m As ButtonRow)
        For i = 0 To 7
            sbtns(i).Text = m.Row(i).FaceText
            lbls(i).Text = m.Row(i).Label

        Next

    End Function
    Private Function OnLetterChanged(m As Keys) Handles buttons.LetterChanged
        lblAlpha.Text = Chr(m)
        TranscribeButtons(buttons.CurrentRow)

    End Function

    Private Sub ButtonForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                If e.Shift Then
                    buttons.CurrentRow.Row(e.KeyCode - Keys.F5).Path = Media.MediaDirectory
                    TranscribeButtons(buttons.CurrentRow)

                Else
                    Dim s As String = buttons.CurrentRow.Row(e.KeyCode - Keys.F5).Path
                    If s <> "" Then
                        Media.MediaDirectory = s
                    Else
                        Exit Sub
                    End If
                    ChangeFolder(s)
                    'CancelDisplay()
                    MainForm.tvMain2.SelectedFolder = Media.MediaDirectory
                End If

            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9
                buttons.CurrentLetter = e.KeyCode
            Case Else
                MainForm.frmMain_KeyDown(sender, e)

        End Select


    End Sub
End Class