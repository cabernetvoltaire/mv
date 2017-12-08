<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FindDuplicates
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FindDuplicates))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lbxsorted = New System.Windows.Forms.ListBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lbxunique = New System.Windows.Forms.ListBox()
        Me.lbxDeleteList = New System.Windows.Forms.ListBox()
        Me.lbxSave = New System.Windows.Forms.ListBox()
        Me.lbxDuplicates = New System.Windows.Forms.ListBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.AutoList = New System.Windows.Forms.Button()
        Me.btnDeleteFiles = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.FlowLayoutPanel3 = New System.Windows.Forms.FlowLayoutPanel()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.lblDelete = New System.Windows.Forms.Label()
        Me.lblUnique = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.lblSorted = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(511, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 25)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Label1"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.lbxsorted)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.FlowLayoutPanel2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.FlowLayoutPanel3)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblInfo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblDelete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblUnique)
        Me.SplitContainer1.Panel2.Controls.Add(Me.FlowLayoutPanel1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblSorted)
        Me.SplitContainer1.Size = New System.Drawing.Size(2614, 1640)
        Me.SplitContainer1.SplitterDistance = 600
        Me.SplitContainer1.TabIndex = 3
        '
        'lbxsorted
        '
        Me.lbxsorted.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbxsorted.FormattingEnabled = True
        Me.lbxsorted.ItemHeight = 24
        Me.lbxsorted.Location = New System.Drawing.Point(0, 0)
        Me.lbxsorted.Name = "lbxsorted"
        Me.lbxsorted.Size = New System.Drawing.Size(600, 1640)
        Me.lbxsorted.TabIndex = 4
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 622.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.lbxunique, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.lbxDeleteList, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.lbxSave, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.lbxDuplicates, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label6, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel2, 2, 3)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 873)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 57.70863!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 42.29137!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(2010, 732)
        Me.TableLayoutPanel1.TabIndex = 36
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label4.Location = New System.Drawing.Point(697, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(688, 25)
        Me.Label4.TabIndex = 34
        Me.Label4.Text = "Duplicates of file clicked"
        '
        'lbxunique
        '
        Me.lbxunique.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbxunique.FormattingEnabled = True
        Me.lbxunique.ItemHeight = 24
        Me.lbxunique.Location = New System.Drawing.Point(3, 28)
        Me.lbxunique.Name = "lbxunique"
        Me.lbxunique.Size = New System.Drawing.Size(688, 378)
        Me.lbxunique.TabIndex = 5
        '
        'lbxDeleteList
        '
        Me.lbxDeleteList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbxDeleteList.FormattingEnabled = True
        Me.lbxDeleteList.ItemHeight = 24
        Me.lbxDeleteList.Location = New System.Drawing.Point(697, 452)
        Me.lbxDeleteList.Name = "lbxDeleteList"
        Me.lbxDeleteList.Size = New System.Drawing.Size(688, 277)
        Me.lbxDeleteList.TabIndex = 8
        '
        'lbxSave
        '
        Me.lbxSave.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbxSave.FormattingEnabled = True
        Me.lbxSave.ItemHeight = 24
        Me.lbxSave.Location = New System.Drawing.Point(3, 452)
        Me.lbxSave.Name = "lbxSave"
        Me.lbxSave.Size = New System.Drawing.Size(688, 277)
        Me.lbxSave.TabIndex = 6
        '
        'lbxDuplicates
        '
        Me.lbxDuplicates.FormattingEnabled = True
        Me.lbxDuplicates.ItemHeight = 24
        Me.lbxDuplicates.Location = New System.Drawing.Point(697, 28)
        Me.lbxDuplicates.Name = "lbxDuplicates"
        Me.lbxDuplicates.Size = New System.Drawing.Size(688, 364)
        Me.lbxDuplicates.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label3.Location = New System.Drawing.Point(3, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(688, 25)
        Me.Label3.TabIndex = 33
        Me.Label3.Text = "Unique Files with Duplicates (Click to see)"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 409)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(407, 25)
        Me.Label5.TabIndex = 35
        Me.Label5.Text = "Files that will NOT be deleted (Click to switch)"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(697, 409)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(380, 25)
        Me.Label6.TabIndex = 36
        Me.Label6.Text = "Files that WILL be deleted (Click to switch)"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.RadioButton2)
        Me.Panel1.Controls.Add(Me.RadioButton1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(1391, 28)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(616, 378)
        Me.Panel1.TabIndex = 37
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(122, 148)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(154, 29)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "RadioButton2"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(122, 90)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(154, 29)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "RadioButton1"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.AutoList)
        Me.Panel2.Controls.Add(Me.btnDeleteFiles)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(1391, 452)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(616, 277)
        Me.Panel2.TabIndex = 38
        '
        'AutoList
        '
        Me.AutoList.Location = New System.Drawing.Point(104, 117)
        Me.AutoList.Name = "AutoList"
        Me.AutoList.Size = New System.Drawing.Size(317, 48)
        Me.AutoList.TabIndex = 33
        Me.AutoList.Text = "Find &All"
        Me.AutoList.UseVisualStyleBackColor = True
        '
        'btnDeleteFiles
        '
        Me.btnDeleteFiles.Location = New System.Drawing.Point(104, 28)
        Me.btnDeleteFiles.Name = "btnDeleteFiles"
        Me.btnDeleteFiles.Size = New System.Drawing.Size(317, 48)
        Me.btnDeleteFiles.TabIndex = 32
        Me.btnDeleteFiles.Text = "Delete &Files"
        Me.btnDeleteFiles.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(682, 950)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(255, 25)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Unique Files with Duplicates"
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.FlowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(2010, 873)
        Me.FlowLayoutPanel2.TabIndex = 30
        '
        'FlowLayoutPanel3
        '
        Me.FlowLayoutPanel3.Location = New System.Drawing.Point(131, 899)
        Me.FlowLayoutPanel3.Name = "FlowLayoutPanel3"
        Me.FlowLayoutPanel3.Size = New System.Drawing.Size(490, 48)
        Me.FlowLayoutPanel3.TabIndex = 9
        '
        'lblInfo
        '
        Me.lblInfo.AutoSize = True
        Me.lblInfo.Location = New System.Drawing.Point(89, 1353)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(0, 25)
        Me.lblInfo.TabIndex = 5
        '
        'lblDelete
        '
        Me.lblDelete.AutoSize = True
        Me.lblDelete.Location = New System.Drawing.Point(969, 524)
        Me.lblDelete.Name = "lblDelete"
        Me.lblDelete.Size = New System.Drawing.Size(94, 25)
        Me.lblDelete.TabIndex = 6
        Me.lblDelete.Text = "To delete"
        '
        'lblUnique
        '
        Me.lblUnique.AutoSize = True
        Me.lblUnique.Location = New System.Drawing.Point(598, 532)
        Me.lblUnique.Name = "lblUnique"
        Me.lblUnique.Size = New System.Drawing.Size(74, 25)
        Me.lblUnique.TabIndex = 5
        Me.lblUnique.Text = "Unique"
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(2010, 397)
        Me.FlowLayoutPanel1.TabIndex = 3
        '
        'lblSorted
        '
        Me.lblSorted.AutoSize = True
        Me.lblSorted.Location = New System.Drawing.Point(183, 403)
        Me.lblSorted.Name = "lblSorted"
        Me.lblSorted.Size = New System.Drawing.Size(100, 25)
        Me.lblSorted.TabIndex = 4
        Me.lblSorted.Text = "SortedList"
        '
        'FindDuplicates
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2614, 1640)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FindDuplicates"
        Me.Text = "FindDuplicates"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents lblInfo As Label
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents lblDelete As Label
    Friend WithEvents lblUnique As Label
    Friend WithEvents lblSorted As Label
    Friend WithEvents lbxDeleteList As ListBox
    Friend WithEvents FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents lbxsorted As ListBox
    Friend WithEvents lbxunique As ListBox
    Friend WithEvents lbxDuplicates As ListBox
    Friend WithEvents lbxSave As ListBox
    Friend WithEvents btnDeleteFiles As Button
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents FlowLayoutPanel3 As FlowLayoutPanel
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents Panel2 As Panel
    Friend WithEvents AutoList As Button
End Class
