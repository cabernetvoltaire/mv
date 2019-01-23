<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FolderSelect
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FolderSelect))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.fst1 = New MasaSam.Forms.Controls.FileSystemTree()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnAssign = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PreviewWMP = New AxWMPLib.AxWindowsMediaPlayer()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.PreviewWMP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.fst1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.PreviewWMP, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 400.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 82.5!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 17.5!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(634, 993)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'fst1
        '
        Me.fst1.AutoExpandNodes = True
        Me.fst1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.fst1.FileExtensions = "*"
        Me.fst1.Location = New System.Drawing.Point(6, 406)
        Me.fst1.Margin = New System.Windows.Forms.Padding(6)
        Me.fst1.Name = "fst1"
        Me.fst1.RootDrive = Nothing
        Me.fst1.SelectedFolder = Nothing
        Me.fst1.Size = New System.Drawing.Size(622, 477)
        Me.fst1.TabIndex = 0
        Me.fst1.TrackDriveState = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnAssign)
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 892)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(628, 98)
        Me.Panel1.TabIndex = 1
        '
        'btnAssign
        '
        Me.btnAssign.Location = New System.Drawing.Point(461, 40)
        Me.btnAssign.Name = "btnAssign"
        Me.btnAssign.Size = New System.Drawing.Size(168, 38)
        Me.btnAssign.TabIndex = 2
        Me.btnAssign.Text = "A&ssign"
        Me.btnAssign.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(14, 49)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(416, 29)
        Me.TextBox1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 25)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Label1"
        '
        'PreviewWMP
        '
        Me.PreviewWMP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PreviewWMP.Enabled = True
        Me.PreviewWMP.Location = New System.Drawing.Point(3, 3)
        Me.PreviewWMP.Name = "PreviewWMP"
        Me.PreviewWMP.OcxState = CType(resources.GetObject("PreviewWMP.OcxState"), System.Windows.Forms.AxHost.State)
        Me.PreviewWMP.Size = New System.Drawing.Size(628, 394)
        Me.PreviewWMP.TabIndex = 4
        '
        'FolderSelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(634, 993)
        Me.ControlBox = False
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FolderSelect"
        Me.ShowIcon = False
        Me.Text = "FolderSelect"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PreviewWMP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents fst1 As Controls.FileSystemTree
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnAssign As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents PreviewWMP As AxWMPLib.AxWindowsMediaPlayer
End Class
