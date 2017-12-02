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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.FileSystemTree1 = New MasaSam.Forms.Controls.FileSystemTree()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnAssign = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.FileSystemTree1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 82.5!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 17.5!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(706, 840)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'FileSystemTree1
        '
        Me.FileSystemTree1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FileSystemTree1.FileExtensions = "*"
        Me.FileSystemTree1.Location = New System.Drawing.Point(6, 6)
        Me.FileSystemTree1.Margin = New System.Windows.Forms.Padding(6)
        Me.FileSystemTree1.Name = "FileSystemTree1"
        Me.FileSystemTree1.RootDrive = Nothing
        Me.FileSystemTree1.SelectedFolder = Nothing
        Me.FileSystemTree1.Size = New System.Drawing.Size(694, 681)
        Me.FileSystemTree1.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnAssign)
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 696)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(700, 141)
        Me.Panel1.TabIndex = 1
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(190, 49)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(212, 29)
        Me.TextBox1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 25)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Label1"
        '
        'btnAssign
        '
        Me.btnAssign.Location = New System.Drawing.Point(485, 34)
        Me.btnAssign.Name = "btnAssign"
        Me.btnAssign.Size = New System.Drawing.Size(168, 65)
        Me.btnAssign.TabIndex = 2
        Me.btnAssign.Text = "A&ssign"
        Me.btnAssign.UseVisualStyleBackColor = True
        '
        'FolderSelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(706, 840)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "FolderSelect"
        Me.Text = "FolderSelect"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents FileSystemTree1 As Controls.FileSystemTree
    Friend WithEvents Panel1 As Panel
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnAssign As Button
End Class
