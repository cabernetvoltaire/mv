<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MediaControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MediaControl))
        Me.wmp = New AxWMPLib.AxWindowsMediaPlayer()
        Me.Blanker = New System.Windows.Forms.PictureBox()
        Me.PicBox = New System.Windows.Forms.PictureBox()
        CType(Me.wmp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Blanker, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'wmp
        '
        Me.wmp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wmp.Enabled = True
        Me.wmp.Location = New System.Drawing.Point(0, 0)
        Me.wmp.Name = "wmp"
        Me.wmp.OcxState = CType(resources.GetObject("wmp.OcxState"), System.Windows.Forms.AxHost.State)
        Me.wmp.Size = New System.Drawing.Size(1160, 714)
        Me.wmp.TabIndex = 22
        '
        'Blanker
        '
        Me.Blanker.BackColor = System.Drawing.Color.Maroon
        Me.Blanker.Location = New System.Drawing.Point(146, 155)
        Me.Blanker.Name = "Blanker"
        Me.Blanker.Size = New System.Drawing.Size(762, 384)
        Me.Blanker.TabIndex = 21
        Me.Blanker.TabStop = False
        Me.Blanker.Visible = False
        '
        'PicBox
        '
        Me.PicBox.BackColor = System.Drawing.Color.Black
        Me.PicBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PicBox.Location = New System.Drawing.Point(0, 0)
        Me.PicBox.Margin = New System.Windows.Forms.Padding(0)
        Me.PicBox.Name = "PicBox"
        Me.PicBox.Size = New System.Drawing.Size(1160, 714)
        Me.PicBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PicBox.TabIndex = 19
        Me.PicBox.TabStop = False
        Me.PicBox.Visible = False
        '
        'MediaControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.wmp)
        Me.Controls.Add(Me.Blanker)
        Me.Controls.Add(Me.PicBox)
        Me.Name = "MediaControl"
        Me.Size = New System.Drawing.Size(1160, 714)
        CType(Me.wmp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Blanker, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PicBox As PictureBox
    Friend WithEvents Blanker As PictureBox
    Friend WithEvents wmp As AxWMPLib.AxWindowsMediaPlayer
End Class
