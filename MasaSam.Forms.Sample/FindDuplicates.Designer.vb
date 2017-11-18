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
        Me.lbxSave = New System.Windows.Forms.ListBox()
        Me.btnDeleteFiles = New System.Windows.Forms.Button()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.WMP1 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMP2 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMP3 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMP4 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMP5 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMP6 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMP7 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMP8 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMP9 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMp10 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMP11 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.WMP12 = New AxWMPLib.AxWindowsMediaPlayer()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.lbxDeleteList = New System.Windows.Forms.ListBox()
        Me.lblDelete = New System.Windows.Forms.Label()
        Me.lblUnique = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.lbxunique = New System.Windows.Forms.ListBox()
        Me.lbxDuplicates = New System.Windows.Forms.ListBox()
        Me.lblSorted = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel3 = New System.Windows.Forms.FlowLayoutPanel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.FlowLayoutPanel2.SuspendLayout()
        CType(Me.WMP1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMP2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMP3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMP4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMP5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMP6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMP7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMP8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMP9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMp10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMP11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WMP12, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
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
        'btnDeleteFiles
        '
        Me.btnDeleteFiles.Location = New System.Drawing.Point(1391, 452)
        Me.btnDeleteFiles.Name = "btnDeleteFiles"
        Me.btnDeleteFiles.Size = New System.Drawing.Size(317, 48)
        Me.btnDeleteFiles.TabIndex = 32
        Me.btnDeleteFiles.Text = "Delete &Files"
        Me.btnDeleteFiles.UseVisualStyleBackColor = True
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP1)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP2)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP3)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP4)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP5)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP6)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP7)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP8)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP9)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMp10)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP11)
        Me.FlowLayoutPanel2.Controls.Add(Me.WMP12)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.FlowLayoutPanel2.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(2010, 873)
        Me.FlowLayoutPanel2.TabIndex = 30
        '
        'WMP1
        '
        Me.WMP1.Enabled = True
        Me.WMP1.Location = New System.Drawing.Point(3, 3)
        Me.WMP1.Name = "WMP1"
        Me.WMP1.OcxState = CType(resources.GetObject("WMP1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP1.Size = New System.Drawing.Size(298, 398)
        Me.WMP1.TabIndex = 26
        Me.WMP1.TabStop = False
        Me.WMP1.UseWaitCursor = True
        Me.WMP1.Visible = False
        '
        'WMP2
        '
        Me.WMP2.Enabled = True
        Me.WMP2.Location = New System.Drawing.Point(3, 407)
        Me.WMP2.Name = "WMP2"
        Me.WMP2.OcxState = CType(resources.GetObject("WMP2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP2.Size = New System.Drawing.Size(286, 398)
        Me.WMP2.TabIndex = 27
        Me.WMP2.TabStop = False
        Me.WMP2.UseWaitCursor = True
        Me.WMP2.Visible = False
        '
        'WMP3
        '
        Me.WMP3.Enabled = True
        Me.WMP3.Location = New System.Drawing.Point(307, 3)
        Me.WMP3.Name = "WMP3"
        Me.WMP3.OcxState = CType(resources.GetObject("WMP3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP3.Size = New System.Drawing.Size(305, 398)
        Me.WMP3.TabIndex = 28
        Me.WMP3.TabStop = False
        Me.WMP3.UseWaitCursor = True
        Me.WMP3.Visible = False
        '
        'WMP4
        '
        Me.WMP4.Enabled = True
        Me.WMP4.Location = New System.Drawing.Point(307, 407)
        Me.WMP4.Name = "WMP4"
        Me.WMP4.OcxState = CType(resources.GetObject("WMP4.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP4.Size = New System.Drawing.Size(303, 398)
        Me.WMP4.TabIndex = 29
        Me.WMP4.TabStop = False
        Me.WMP4.UseWaitCursor = True
        Me.WMP4.Visible = False
        '
        'WMP5
        '
        Me.WMP5.Enabled = True
        Me.WMP5.Location = New System.Drawing.Point(618, 3)
        Me.WMP5.Name = "WMP5"
        Me.WMP5.OcxState = CType(resources.GetObject("WMP5.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP5.Size = New System.Drawing.Size(298, 398)
        Me.WMP5.TabIndex = 27
        Me.WMP5.TabStop = False
        Me.WMP5.UseWaitCursor = True
        Me.WMP5.Visible = False
        '
        'WMP6
        '
        Me.WMP6.Enabled = True
        Me.WMP6.Location = New System.Drawing.Point(618, 407)
        Me.WMP6.Name = "WMP6"
        Me.WMP6.OcxState = CType(resources.GetObject("WMP6.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP6.Size = New System.Drawing.Size(298, 398)
        Me.WMP6.TabIndex = 30
        Me.WMP6.TabStop = False
        Me.WMP6.UseWaitCursor = True
        Me.WMP6.Visible = False
        '
        'WMP7
        '
        Me.WMP7.Enabled = True
        Me.WMP7.Location = New System.Drawing.Point(922, 3)
        Me.WMP7.Name = "WMP7"
        Me.WMP7.OcxState = CType(resources.GetObject("WMP7.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP7.Size = New System.Drawing.Size(298, 398)
        Me.WMP7.TabIndex = 31
        Me.WMP7.TabStop = False
        Me.WMP7.UseWaitCursor = True
        Me.WMP7.Visible = False
        '
        'WMP8
        '
        Me.WMP8.Enabled = True
        Me.WMP8.Location = New System.Drawing.Point(922, 407)
        Me.WMP8.Name = "WMP8"
        Me.WMP8.OcxState = CType(resources.GetObject("WMP8.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP8.Size = New System.Drawing.Size(298, 398)
        Me.WMP8.TabIndex = 33
        Me.WMP8.TabStop = False
        Me.WMP8.UseWaitCursor = True
        Me.WMP8.Visible = False
        '
        'WMP9
        '
        Me.WMP9.Enabled = True
        Me.WMP9.Location = New System.Drawing.Point(1226, 3)
        Me.WMP9.Name = "WMP9"
        Me.WMP9.OcxState = CType(resources.GetObject("WMP9.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP9.Size = New System.Drawing.Size(298, 398)
        Me.WMP9.TabIndex = 32
        Me.WMP9.TabStop = False
        Me.WMP9.UseWaitCursor = True
        Me.WMP9.Visible = False
        '
        'WMp10
        '
        Me.WMp10.Enabled = True
        Me.WMp10.Location = New System.Drawing.Point(1226, 407)
        Me.WMp10.Name = "WMp10"
        Me.WMp10.OcxState = CType(resources.GetObject("WMp10.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMp10.Size = New System.Drawing.Size(298, 398)
        Me.WMp10.TabIndex = 34
        Me.WMp10.TabStop = False
        Me.WMp10.UseWaitCursor = True
        Me.WMp10.Visible = False
        '
        'WMP11
        '
        Me.WMP11.Enabled = True
        Me.WMP11.Location = New System.Drawing.Point(1530, 3)
        Me.WMP11.Name = "WMP11"
        Me.WMP11.OcxState = CType(resources.GetObject("WMP11.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP11.Size = New System.Drawing.Size(298, 398)
        Me.WMP11.TabIndex = 35
        Me.WMP11.TabStop = False
        Me.WMP11.UseWaitCursor = True
        Me.WMP11.Visible = False
        '
        'WMP12
        '
        Me.WMP12.Enabled = True
        Me.WMP12.Location = New System.Drawing.Point(1530, 407)
        Me.WMP12.Name = "WMP12"
        Me.WMP12.OcxState = CType(resources.GetObject("WMP12.OcxState"), System.Windows.Forms.AxHost.State)
        Me.WMP12.Size = New System.Drawing.Size(298, 398)
        Me.WMP12.TabIndex = 36
        Me.WMP12.TabStop = False
        Me.WMP12.UseWaitCursor = True
        Me.WMP12.Visible = False
        '
        'lblInfo
        '
        Me.lblInfo.AutoSize = True
        Me.lblInfo.Location = New System.Drawing.Point(89, 1353)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(0, 25)
        Me.lblInfo.TabIndex = 5
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
        'lbxDuplicates
        '
        Me.lbxDuplicates.FormattingEnabled = True
        Me.lbxDuplicates.ItemHeight = 24
        Me.lbxDuplicates.Location = New System.Drawing.Point(697, 28)
        Me.lbxDuplicates.Name = "lbxDuplicates"
        Me.lbxDuplicates.Size = New System.Drawing.Size(688, 364)
        Me.lbxDuplicates.TabIndex = 7
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
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(682, 950)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(255, 25)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Unique Files with Duplicates"
        '
        'FlowLayoutPanel3
        '
        Me.FlowLayoutPanel3.Location = New System.Drawing.Point(131, 899)
        Me.FlowLayoutPanel3.Name = "FlowLayoutPanel3"
        Me.FlowLayoutPanel3.Size = New System.Drawing.Size(490, 48)
        Me.FlowLayoutPanel3.TabIndex = 9
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
        Me.TableLayoutPanel1.Controls.Add(Me.btnDeleteFiles, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label6, 1, 2)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 873)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 57.70863!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 42.29137!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(2010, 732)
        Me.TableLayoutPanel1.TabIndex = 36
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
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.FlowLayoutPanel2.ResumeLayout(False)
        CType(Me.WMP1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMP2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMP3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMP4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMP5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMP6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMP7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMP8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMP9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMp10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMP11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WMP12, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
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
    Friend WithEvents WMP1 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents lbxDeleteList As ListBox
    Friend WithEvents WMP4 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents WMP3 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents WMP2 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents WMP5 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents WMP6 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents WMP7 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents WMP8 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents WMP9 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents WMp10 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents WMP11 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents WMP12 As AxWMPLib.AxWindowsMediaPlayer
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
End Class
