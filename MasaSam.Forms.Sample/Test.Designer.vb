<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Test
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Test))
        Me.TrackBar1 = New System.Windows.Forms.TrackBar()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.pnlFS = New System.Windows.Forms.Panel()
        Me.pbxBlanker = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pbx1 = New System.Windows.Forms.PictureBox()
        Me.MainWMP = New AxWMPLib.AxWindowsMediaPlayer()
        Me.tmrMovePic = New System.Windows.Forms.Timer(Me.components)
        Me.tmrLabelFadeIn = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFS.SuspendLayout()
        CType(Me.pbxBlanker, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbx1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MainWMP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TrackBar1
        '
        Me.TrackBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TrackBar1.Location = New System.Drawing.Point(0, 1320)
        Me.TrackBar1.Name = "TrackBar1"
        Me.TrackBar1.Size = New System.Drawing.Size(2763, 80)
        Me.TrackBar1.TabIndex = 5
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 1290)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(2763, 30)
        Me.ProgressBar1.TabIndex = 6
        '
        'pnlFS
        '
        Me.pnlFS.BackColor = System.Drawing.Color.Orange
        Me.pnlFS.Controls.Add(Me.pbxBlanker)
        Me.pnlFS.Controls.Add(Me.Label1)
        Me.pnlFS.Controls.Add(Me.pbx1)
        Me.pnlFS.Controls.Add(Me.MainWMP)
        Me.pnlFS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFS.Location = New System.Drawing.Point(0, 0)
        Me.pnlFS.Name = "pnlFS"
        Me.pnlFS.Size = New System.Drawing.Size(2763, 1290)
        Me.pnlFS.TabIndex = 7
        '
        'pbxBlanker
        '
        Me.pbxBlanker.Location = New System.Drawing.Point(796, 320)
        Me.pbxBlanker.Name = "pbxBlanker"
        Me.pbxBlanker.Size = New System.Drawing.Size(563, 535)
        Me.pbxBlanker.TabIndex = 4
        Me.pbxBlanker.TabStop = False
        Me.pbxBlanker.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Font = New System.Drawing.Font("Calibri", 14.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.LimeGreen
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(108, 41)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Label1"
        Me.Label1.Visible = False
        '
        'pbx1
        '
        Me.pbx1.BackColor = System.Drawing.Color.Maroon
        Me.pbx1.Location = New System.Drawing.Point(1542, 310)
        Me.pbx1.Name = "pbx1"
        Me.pbx1.Size = New System.Drawing.Size(577, 545)
        Me.pbx1.TabIndex = 5
        Me.pbx1.TabStop = False
        '
        'MainWMP
        '
        Me.MainWMP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MainWMP.Enabled = True
        Me.MainWMP.Location = New System.Drawing.Point(0, 0)
        Me.MainWMP.Name = "MainWMP"
        Me.MainWMP.OcxState = CType(resources.GetObject("MainWMP.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MainWMP.Size = New System.Drawing.Size(2763, 1290)
        Me.MainWMP.TabIndex = 6
        Me.MainWMP.TabStop = False
        Me.MainWMP.UseWaitCursor = True
        Me.MainWMP.Visible = False
        '
        'Test
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.Color.BlanchedAlmond
        Me.ClientSize = New System.Drawing.Size(2763, 1400)
        Me.Controls.Add(Me.pnlFS)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.TrackBar1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.KeyPreview = True
        Me.Name = "Test"
        Me.Text = "Test"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.TrackBar1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFS.ResumeLayout(False)
        Me.pnlFS.PerformLayout()
        CType(Me.pbxBlanker, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbx1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MainWMP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TrackBar1 As TrackBar
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents pnlFS As Panel
    Friend WithEvents tmrMovePic As Timer
    Friend WithEvents Label1 As Label
    Friend WithEvents pbxBlanker As PictureBox
    Friend WithEvents pbx1 As PictureBox
    Friend WithEvents tmrLabelFadeIn As Timer
    Friend WithEvents MainWMP As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
