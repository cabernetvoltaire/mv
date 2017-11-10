<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FullScreen
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FullScreen))
        Me.FSWMP = New AxWMPLib.AxWindowsMediaPlayer()
        Me.fullScreenPicBox = New System.Windows.Forms.PictureBox()
        Me.FSBlanker = New System.Windows.Forms.PictureBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.FSWMP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fullScreenPicBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSBlanker, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FSWMP
        '
        Me.FSWMP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FSWMP.Enabled = True
        Me.FSWMP.Location = New System.Drawing.Point(0, 0)
        Me.FSWMP.Name = "FSWMP"
        Me.FSWMP.OcxState = CType(resources.GetObject("FSWMP.OcxState"), System.Windows.Forms.AxHost.State)
        Me.FSWMP.Size = New System.Drawing.Size(2492, 1144)
        Me.FSWMP.TabIndex = 0
        Me.FSWMP.TabStop = False
        Me.FSWMP.UseWaitCursor = True
        '
        'fullScreenPicBox
        '
        Me.fullScreenPicBox.BackColor = System.Drawing.Color.Black
        Me.fullScreenPicBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.fullScreenPicBox.Location = New System.Drawing.Point(0, 0)
        Me.fullScreenPicBox.Name = "fullScreenPicBox"
        Me.fullScreenPicBox.Size = New System.Drawing.Size(2492, 1144)
        Me.fullScreenPicBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.fullScreenPicBox.TabIndex = 18
        Me.fullScreenPicBox.TabStop = False
        Me.fullScreenPicBox.Visible = False
        '
        'FSBlanker
        '
        Me.FSBlanker.Location = New System.Drawing.Point(688, 457)
        Me.FSBlanker.Name = "FSBlanker"
        Me.FSBlanker.Size = New System.Drawing.Size(1193, 660)
        Me.FSBlanker.TabIndex = 19
        Me.FSBlanker.TabStop = False
        Me.FSBlanker.Visible = False
        '
        'Timer1
        '
        '
        'FullScreen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(2492, 1144)
        Me.Controls.Add(Me.FSBlanker)
        Me.Controls.Add(Me.fullScreenPicBox)
        Me.Controls.Add(Me.FSWMP)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Name = "FullScreen"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Form1"
        CType(Me.FSWMP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fullScreenPicBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSBlanker, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents FSWMP As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents fullScreenPicBox As PictureBox
    Friend WithEvents FSBlanker As PictureBox
    Friend WithEvents Timer1 As Timer
End Class
