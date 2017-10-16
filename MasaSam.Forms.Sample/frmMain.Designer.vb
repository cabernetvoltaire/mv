<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.LoadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ShowListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ListsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddAllPicturesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CurrentOnlyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IncludingSubfoldersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddAllVideosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CurrentOnlyToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.IncludingSubsetsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddAllFilesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CurrentOnlyToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.AllSubFoldersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddPicturesAndVideosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CurrentOnlyToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.AllSubfoldersToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.SlideshowToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.SlowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NormalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FastToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VideoToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.SpeedToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SlowToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.NormalToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.FastToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.RandomStartToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.SaveAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FullScreenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TrueSizeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ButtonsToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.SwitchModesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AssignNextToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SlideShowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VideoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.tsslblFiles = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslblFilter = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslblRandom = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslblSPEED = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslblSTART = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslblZOOM = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslblShowfile = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsslblLastfile = New System.Windows.Forms.ToolStripStatusLabel()
        Me.DateSSL = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tssFolderInfo = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tssMoveCopy = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tssFilter = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel4 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.stripLastDone = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.ctrFilesandPics = New System.Windows.Forms.SplitContainer()
        Me.ctrTreeandFiles = New System.Windows.Forms.SplitContainer()
        Me.tvMain2 = New MasaSam.Forms.Controls.FileSystemTree()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lbxFiles = New System.Windows.Forms.ListBox()
        Me.lbxShowList = New System.Windows.Forms.ListBox()
        Me.MainWMP = New AxWMPLib.AxWindowsMediaPlayer()
        Me.pbxBlanker = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.tsbAddMovies = New System.Windows.Forms.ToolStripButton()
        Me.TSBRemoveMovies = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton8 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbClear = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton10 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton11 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton12 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton13 = New System.Windows.Forms.ToolStripComboBox()
        Me.ToolStripButton15 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator10 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripComboBox1 = New System.Windows.Forms.ToolStripComboBox()
        Me.ToolStripButton16 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripTextBox1 = New System.Windows.Forms.ToolStripTextBox()
        Me.ToolStripButton9 = New System.Windows.Forms.ToolStripButton()
        Me.btnChooseRandom = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton5 = New System.Windows.Forms.ToolStripButton()
        Me.tsbFullscreen = New System.Windows.Forms.ToolStripButton()
        Me.toolStripSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.toolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.showButtons = New System.Windows.Forms.ToolStripButton()
        Me.tmrUpdateFileList = New System.Windows.Forms.Timer(Me.components)
        Me.tmrPicLoad = New System.Windows.Forms.Timer(Me.components)
        Me.tmrJumpVideo = New System.Windows.Forms.Timer(Me.components)
        Me.tmrInitialise = New System.Windows.Forms.Timer(Me.components)
        Me.tmrCheckFolders = New System.Windows.Forms.Timer(Me.components)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.tmrSlideShow = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.FileSystemWatcher1 = New System.IO.FileSystemWatcher()
        Me.tmrLoadLastFolder = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        CType(Me.ctrFilesandPics, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ctrFilesandPics.Panel1.SuspendLayout()
        Me.ctrFilesandPics.Panel2.SuspendLayout()
        Me.ctrFilesandPics.SuspendLayout()
        CType(Me.ctrTreeandFiles, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ctrTreeandFiles.Panel1.SuspendLayout()
        Me.ctrTreeandFiles.Panel2.SuspendLayout()
        Me.ctrTreeandFiles.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.MainWMP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbxBlanker, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(28, 28)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem1, Me.ListsToolStripMenuItem, Me.SlideshowToolStripMenuItem1, Me.VideoToolStripMenuItem1})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(2603, 38)
        Me.MenuStrip1.TabIndex = 15
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem1
        '
        Me.FileToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LoadToolStripMenuItem, Me.SaveToolStripMenuItem1})
        Me.FileToolStripMenuItem1.Name = "FileToolStripMenuItem1"
        Me.FileToolStripMenuItem1.Size = New System.Drawing.Size(56, 34)
        Me.FileToolStripMenuItem1.Text = "File"
        '
        'LoadToolStripMenuItem
        '
        Me.LoadToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ShowListToolStripMenuItem})
        Me.LoadToolStripMenuItem.Name = "LoadToolStripMenuItem"
        Me.LoadToolStripMenuItem.Size = New System.Drawing.Size(149, 34)
        Me.LoadToolStripMenuItem.Text = "Load"
        '
        'ShowListToolStripMenuItem
        '
        Me.ShowListToolStripMenuItem.Name = "ShowListToolStripMenuItem"
        Me.ShowListToolStripMenuItem.Size = New System.Drawing.Size(180, 34)
        Me.ShowListToolStripMenuItem.Text = "Showlist"
        '
        'SaveToolStripMenuItem1
        '
        Me.SaveToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SaveListToolStripMenuItem})
        Me.SaveToolStripMenuItem1.Name = "SaveToolStripMenuItem1"
        Me.SaveToolStripMenuItem1.Size = New System.Drawing.Size(149, 34)
        Me.SaveToolStripMenuItem1.Text = "Save"
        '
        'SaveListToolStripMenuItem
        '
        Me.SaveListToolStripMenuItem.Name = "SaveListToolStripMenuItem"
        Me.SaveListToolStripMenuItem.Size = New System.Drawing.Size(180, 34)
        Me.SaveListToolStripMenuItem.Text = "Showlist"
        '
        'ListsToolStripMenuItem
        '
        Me.ListsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddAllPicturesToolStripMenuItem, Me.AddAllVideosToolStripMenuItem, Me.AddAllFilesToolStripMenuItem, Me.AddPicturesAndVideosToolStripMenuItem})
        Me.ListsToolStripMenuItem.Name = "ListsToolStripMenuItem"
        Me.ListsToolStripMenuItem.Size = New System.Drawing.Size(65, 34)
        Me.ListsToolStripMenuItem.Text = "Lists"
        '
        'AddAllPicturesToolStripMenuItem
        '
        Me.AddAllPicturesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CurrentOnlyToolStripMenuItem, Me.IncludingSubfoldersToolStripMenuItem})
        Me.AddAllPicturesToolStripMenuItem.Name = "AddAllPicturesToolStripMenuItem"
        Me.AddAllPicturesToolStripMenuItem.Size = New System.Drawing.Size(327, 34)
        Me.AddAllPicturesToolStripMenuItem.Text = "Add all pictures"
        '
        'CurrentOnlyToolStripMenuItem
        '
        Me.CurrentOnlyToolStripMenuItem.Name = "CurrentOnlyToolStripMenuItem"
        Me.CurrentOnlyToolStripMenuItem.Size = New System.Drawing.Size(292, 34)
        Me.CurrentOnlyToolStripMenuItem.Text = "Current only"
        '
        'IncludingSubfoldersToolStripMenuItem
        '
        Me.IncludingSubfoldersToolStripMenuItem.Checked = True
        Me.IncludingSubfoldersToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.IncludingSubfoldersToolStripMenuItem.Name = "IncludingSubfoldersToolStripMenuItem"
        Me.IncludingSubfoldersToolStripMenuItem.Size = New System.Drawing.Size(292, 34)
        Me.IncludingSubfoldersToolStripMenuItem.Text = "Including subfolders"
        '
        'AddAllVideosToolStripMenuItem
        '
        Me.AddAllVideosToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CurrentOnlyToolStripMenuItem1, Me.IncludingSubsetsToolStripMenuItem})
        Me.AddAllVideosToolStripMenuItem.Name = "AddAllVideosToolStripMenuItem"
        Me.AddAllVideosToolStripMenuItem.Size = New System.Drawing.Size(327, 34)
        Me.AddAllVideosToolStripMenuItem.Text = "Add all videos"
        '
        'CurrentOnlyToolStripMenuItem1
        '
        Me.CurrentOnlyToolStripMenuItem1.Name = "CurrentOnlyToolStripMenuItem1"
        Me.CurrentOnlyToolStripMenuItem1.Size = New System.Drawing.Size(265, 34)
        Me.CurrentOnlyToolStripMenuItem1.Text = "Current only"
        '
        'IncludingSubsetsToolStripMenuItem
        '
        Me.IncludingSubsetsToolStripMenuItem.Name = "IncludingSubsetsToolStripMenuItem"
        Me.IncludingSubsetsToolStripMenuItem.Size = New System.Drawing.Size(265, 34)
        Me.IncludingSubsetsToolStripMenuItem.Text = "Including subsets"
        '
        'AddAllFilesToolStripMenuItem
        '
        Me.AddAllFilesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CurrentOnlyToolStripMenuItem2, Me.AllSubFoldersToolStripMenuItem})
        Me.AddAllFilesToolStripMenuItem.Name = "AddAllFilesToolStripMenuItem"
        Me.AddAllFilesToolStripMenuItem.Size = New System.Drawing.Size(327, 34)
        Me.AddAllFilesToolStripMenuItem.Text = "Add all files"
        '
        'CurrentOnlyToolStripMenuItem2
        '
        Me.CurrentOnlyToolStripMenuItem2.Name = "CurrentOnlyToolStripMenuItem2"
        Me.CurrentOnlyToolStripMenuItem2.Size = New System.Drawing.Size(236, 34)
        Me.CurrentOnlyToolStripMenuItem2.Text = "Current only"
        '
        'AllSubFoldersToolStripMenuItem
        '
        Me.AllSubFoldersToolStripMenuItem.Name = "AllSubFoldersToolStripMenuItem"
        Me.AllSubFoldersToolStripMenuItem.Size = New System.Drawing.Size(236, 34)
        Me.AllSubFoldersToolStripMenuItem.Text = "All sub folders"
        '
        'AddPicturesAndVideosToolStripMenuItem
        '
        Me.AddPicturesAndVideosToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CurrentOnlyToolStripMenuItem3, Me.AllSubfoldersToolStripMenuItem1})
        Me.AddPicturesAndVideosToolStripMenuItem.Name = "AddPicturesAndVideosToolStripMenuItem"
        Me.AddPicturesAndVideosToolStripMenuItem.Size = New System.Drawing.Size(327, 34)
        Me.AddPicturesAndVideosToolStripMenuItem.Text = "Add pictures and videos"
        '
        'CurrentOnlyToolStripMenuItem3
        '
        Me.CurrentOnlyToolStripMenuItem3.Name = "CurrentOnlyToolStripMenuItem3"
        Me.CurrentOnlyToolStripMenuItem3.Size = New System.Drawing.Size(230, 34)
        Me.CurrentOnlyToolStripMenuItem3.Text = "Current only"
        '
        'AllSubfoldersToolStripMenuItem1
        '
        Me.AllSubfoldersToolStripMenuItem1.Name = "AllSubfoldersToolStripMenuItem1"
        Me.AllSubfoldersToolStripMenuItem1.Size = New System.Drawing.Size(230, 34)
        Me.AllSubfoldersToolStripMenuItem1.Text = "All subfolders"
        '
        'SlideshowToolStripMenuItem1
        '
        Me.SlideshowToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SlowToolStripMenuItem, Me.NormalToolStripMenuItem, Me.FastToolStripMenuItem})
        Me.SlideshowToolStripMenuItem1.Name = "SlideshowToolStripMenuItem1"
        Me.SlideshowToolStripMenuItem1.Size = New System.Drawing.Size(117, 34)
        Me.SlideshowToolStripMenuItem1.Text = "Slideshow"
        '
        'SlowToolStripMenuItem
        '
        Me.SlowToolStripMenuItem.Name = "SlowToolStripMenuItem"
        Me.SlowToolStripMenuItem.Size = New System.Drawing.Size(173, 34)
        Me.SlowToolStripMenuItem.Text = "Slow"
        '
        'NormalToolStripMenuItem
        '
        Me.NormalToolStripMenuItem.Name = "NormalToolStripMenuItem"
        Me.NormalToolStripMenuItem.Size = New System.Drawing.Size(173, 34)
        Me.NormalToolStripMenuItem.Text = "Normal"
        '
        'FastToolStripMenuItem
        '
        Me.FastToolStripMenuItem.Name = "FastToolStripMenuItem"
        Me.FastToolStripMenuItem.Size = New System.Drawing.Size(173, 34)
        Me.FastToolStripMenuItem.Text = "Fast"
        '
        'VideoToolStripMenuItem1
        '
        Me.VideoToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SpeedToolStripMenuItem, Me.RandomStartToolStripMenuItem})
        Me.VideoToolStripMenuItem1.Name = "VideoToolStripMenuItem1"
        Me.VideoToolStripMenuItem1.Size = New System.Drawing.Size(78, 34)
        Me.VideoToolStripMenuItem1.Text = "Video"
        '
        'SpeedToolStripMenuItem
        '
        Me.SpeedToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SlowToolStripMenuItem1, Me.NormalToolStripMenuItem1, Me.FastToolStripMenuItem1})
        Me.SpeedToolStripMenuItem.Name = "SpeedToolStripMenuItem"
        Me.SpeedToolStripMenuItem.Size = New System.Drawing.Size(230, 34)
        Me.SpeedToolStripMenuItem.Text = "Speed"
        '
        'SlowToolStripMenuItem1
        '
        Me.SlowToolStripMenuItem1.Name = "SlowToolStripMenuItem1"
        Me.SlowToolStripMenuItem1.Size = New System.Drawing.Size(173, 34)
        Me.SlowToolStripMenuItem1.Text = "Slow"
        '
        'NormalToolStripMenuItem1
        '
        Me.NormalToolStripMenuItem1.Name = "NormalToolStripMenuItem1"
        Me.NormalToolStripMenuItem1.Size = New System.Drawing.Size(173, 34)
        Me.NormalToolStripMenuItem1.Text = "Normal"
        '
        'FastToolStripMenuItem1
        '
        Me.FastToolStripMenuItem1.Name = "FastToolStripMenuItem1"
        Me.FastToolStripMenuItem1.Size = New System.Drawing.Size(173, 34)
        Me.FastToolStripMenuItem1.Text = "Fast"
        '
        'RandomStartToolStripMenuItem
        '
        Me.RandomStartToolStripMenuItem.Name = "RandomStartToolStripMenuItem"
        Me.RandomStartToolStripMenuItem.Size = New System.Drawing.Size(230, 34)
        Me.RandomStartToolStripMenuItem.Text = "Random Start"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.OpenToolStripMenuItem, Me.ToolStripSeparator1, Me.SaveAsToolStripMenuItem, Me.SaveToolStripMenuItem, Me.ToolStripSeparator2, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(56, 34)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(176, 34)
        Me.NewToolStripMenuItem.Text = "&New"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(176, 34)
        Me.OpenToolStripMenuItem.Text = "&Open"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(173, 6)
        '
        'SaveAsToolStripMenuItem
        '
        Me.SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
        Me.SaveAsToolStripMenuItem.Size = New System.Drawing.Size(176, 34)
        Me.SaveAsToolStripMenuItem.Text = "Save &As"
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(176, 34)
        Me.SaveToolStripMenuItem.Text = "&Save"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(173, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(176, 34)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'ViewToolStripMenuItem
        '
        Me.ViewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FullScreenToolStripMenuItem, Me.TrueSizeToolStripMenuItem})
        Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
        Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(69, 34)
        Me.ViewToolStripMenuItem.Text = "&View"
        '
        'FullScreenToolStripMenuItem
        '
        Me.FullScreenToolStripMenuItem.Name = "FullScreenToolStripMenuItem"
        Me.FullScreenToolStripMenuItem.Size = New System.Drawing.Size(204, 34)
        Me.FullScreenToolStripMenuItem.Text = "&Full Screen"
        '
        'TrueSizeToolStripMenuItem
        '
        Me.TrueSizeToolStripMenuItem.Name = "TrueSizeToolStripMenuItem"
        Me.TrueSizeToolStripMenuItem.Size = New System.Drawing.Size(204, 34)
        Me.TrueSizeToolStripMenuItem.Text = "&True Size"
        '
        'ButtonsToolStripMenuItem1
        '
        Me.ButtonsToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SwitchModesToolStripMenuItem, Me.AssignNextToolStripMenuItem})
        Me.ButtonsToolStripMenuItem1.Name = "ButtonsToolStripMenuItem1"
        Me.ButtonsToolStripMenuItem1.Size = New System.Drawing.Size(96, 34)
        Me.ButtonsToolStripMenuItem1.Text = "Buttons"
        '
        'SwitchModesToolStripMenuItem
        '
        Me.SwitchModesToolStripMenuItem.Name = "SwitchModesToolStripMenuItem"
        Me.SwitchModesToolStripMenuItem.Size = New System.Drawing.Size(233, 34)
        Me.SwitchModesToolStripMenuItem.Text = "Switch &Modes"
        '
        'AssignNextToolStripMenuItem
        '
        Me.AssignNextToolStripMenuItem.Name = "AssignNextToolStripMenuItem"
        Me.AssignNextToolStripMenuItem.Size = New System.Drawing.Size(233, 34)
        Me.AssignNextToolStripMenuItem.Text = "Assign &Next"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(72, 34)
        Me.ToolsToolStripMenuItem.Text = "Tools"
        '
        'SlideShowToolStripMenuItem
        '
        Me.SlideShowToolStripMenuItem.Name = "SlideShowToolStripMenuItem"
        Me.SlideShowToolStripMenuItem.Size = New System.Drawing.Size(125, 34)
        Me.SlideShowToolStripMenuItem.Text = "Slide Show"
        '
        'VideoToolStripMenuItem
        '
        Me.VideoToolStripMenuItem.Name = "VideoToolStripMenuItem"
        Me.VideoToolStripMenuItem.Size = New System.Drawing.Size(78, 34)
        Me.VideoToolStripMenuItem.Text = "Video"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(68, 34)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(28, 28)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsslblFiles, Me.tsslblFilter, Me.tsslblRandom, Me.tsslblSPEED, Me.tsslblSTART, Me.tsslblZOOM, Me.tsslblShowfile, Me.tsslblLastfile, Me.DateSSL})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 1136)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(2603, 35)
        Me.StatusStrip1.TabIndex = 16
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tsslblFiles
        '
        Me.tsslblFiles.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.tsslblFiles.Name = "tsslblFiles"
        Me.tsslblFiles.Size = New System.Drawing.Size(61, 30)
        Me.tsslblFiles.Text = "FILES"
        '
        'tsslblFilter
        '
        Me.tsslblFilter.BorderStyle = System.Windows.Forms.Border3DStyle.RaisedOuter
        Me.tsslblFilter.Name = "tsslblFilter"
        Me.tsslblFilter.Size = New System.Drawing.Size(73, 30)
        Me.tsslblFilter.Text = "FILTER"
        '
        'tsslblRandom
        '
        Me.tsslblRandom.Name = "tsslblRandom"
        Me.tsslblRandom.Size = New System.Drawing.Size(106, 30)
        Me.tsslblRandom.Text = "RANDOM"
        '
        'tsslblSPEED
        '
        Me.tsslblSPEED.Name = "tsslblSPEED"
        Me.tsslblSPEED.Size = New System.Drawing.Size(73, 30)
        Me.tsslblSPEED.Text = "SPEED"
        '
        'tsslblSTART
        '
        Me.tsslblSTART.Name = "tsslblSTART"
        Me.tsslblSTART.Size = New System.Drawing.Size(132, 30)
        Me.tsslblSTART.Text = "STARTPOINT"
        '
        'tsslblZOOM
        '
        Me.tsslblZOOM.Name = "tsslblZOOM"
        Me.tsslblZOOM.Size = New System.Drawing.Size(76, 30)
        Me.tsslblZOOM.Text = "ZOOM"
        '
        'tsslblShowfile
        '
        Me.tsslblShowfile.BorderStyle = System.Windows.Forms.Border3DStyle.RaisedOuter
        Me.tsslblShowfile.Name = "tsslblShowfile"
        Me.tsslblShowfile.Size = New System.Drawing.Size(111, 30)
        Me.tsslblShowfile.Text = "SHOWFILE"
        '
        'tsslblLastfile
        '
        Me.tsslblLastfile.Name = "tsslblLastfile"
        Me.tsslblLastfile.Size = New System.Drawing.Size(97, 30)
        Me.tsslblLastfile.Text = "LASTFILE"
        '
        'DateSSL
        '
        Me.DateSSL.Name = "DateSSL"
        Me.DateSSL.Size = New System.Drawing.Size(61, 30)
        Me.DateSSL.Text = "DATE"
        '
        'tssFolderInfo
        '
        Me.tssFolderInfo.Name = "tssFolderInfo"
        Me.tssFolderInfo.Size = New System.Drawing.Size(107, 30)
        Me.tssFolderInfo.Text = "FolderInfo"
        '
        'tssMoveCopy
        '
        Me.tssMoveCopy.Name = "tssMoveCopy"
        Me.tssMoveCopy.Size = New System.Drawing.Size(71, 30)
        Me.tssMoveCopy.Text = "MOVE"
        '
        'tssFilter
        '
        Me.tssFilter.Name = "tssFilter"
        Me.tssFilter.Size = New System.Drawing.Size(58, 30)
        Me.tssFilter.Text = "Filter"
        '
        'ToolStripStatusLabel4
        '
        Me.ToolStripStatusLabel4.Name = "ToolStripStatusLabel4"
        Me.ToolStripStatusLabel4.Size = New System.Drawing.Size(98, 30)
        Me.ToolStripStatusLabel4.Text = "tssLabel4"
        '
        'stripLastDone
        '
        Me.stripLastDone.Name = "stripLastDone"
        Me.stripLastDone.Size = New System.Drawing.Size(207, 30)
        Me.stripLastDone.Text = "No action performed"
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 29)
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 1
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.ctrFilesandPics, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.ToolStrip1, 0, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(0, 38)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 2
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 8.276937!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 91.72307!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(2603, 1098)
        Me.TableLayoutPanel2.TabIndex = 22
        '
        'ctrFilesandPics
        '
        Me.ctrFilesandPics.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ctrFilesandPics.Location = New System.Drawing.Point(3, 93)
        Me.ctrFilesandPics.Name = "ctrFilesandPics"
        '
        'ctrFilesandPics.Panel1
        '
        Me.ctrFilesandPics.Panel1.Controls.Add(Me.ctrTreeandFiles)
        '
        'ctrFilesandPics.Panel2
        '
        Me.ctrFilesandPics.Panel2.BackColor = System.Drawing.Color.Black
        Me.ctrFilesandPics.Panel2.Controls.Add(Me.MainWMP)
        Me.ctrFilesandPics.Panel2.Controls.Add(Me.pbxBlanker)
        Me.ctrFilesandPics.Panel2.Controls.Add(Me.PictureBox1)
        Me.ctrFilesandPics.Size = New System.Drawing.Size(2597, 1002)
        Me.ctrFilesandPics.SplitterDistance = 345
        Me.ctrFilesandPics.SplitterWidth = 30
        Me.ctrFilesandPics.TabIndex = 15
        Me.ctrFilesandPics.TabStop = False
        '
        'ctrTreeandFiles
        '
        Me.ctrTreeandFiles.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ctrTreeandFiles.Location = New System.Drawing.Point(0, 0)
        Me.ctrTreeandFiles.Name = "ctrTreeandFiles"
        Me.ctrTreeandFiles.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'ctrTreeandFiles.Panel1
        '
        Me.ctrTreeandFiles.Panel1.Controls.Add(Me.tvMain2)
        '
        'ctrTreeandFiles.Panel2
        '
        Me.ctrTreeandFiles.Panel2.Controls.Add(Me.SplitContainer1)
        Me.ctrTreeandFiles.Size = New System.Drawing.Size(345, 1002)
        Me.ctrTreeandFiles.SplitterDistance = 446
        Me.ctrTreeandFiles.SplitterWidth = 15
        Me.ctrTreeandFiles.TabIndex = 0
        Me.ctrTreeandFiles.TabStop = False
        '
        'tvMain2
        '
        Me.tvMain2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvMain2.FileExtensions = "*"
        Me.tvMain2.Location = New System.Drawing.Point(0, 0)
        Me.tvMain2.Margin = New System.Windows.Forms.Padding(6)
        Me.tvMain2.Name = "tvMain2"
        Me.tvMain2.RootDrive = Nothing
        Me.tvMain2.SelectedFolder = Nothing
        Me.tvMain2.Size = New System.Drawing.Size(345, 446)
        Me.tvMain2.TabIndex = 0
        Me.tvMain2.TrackDriveState = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.lbxFiles)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lbxShowList)
        Me.SplitContainer1.Size = New System.Drawing.Size(345, 541)
        Me.SplitContainer1.SplitterDistance = 213
        Me.SplitContainer1.SplitterWidth = 30
        Me.SplitContainer1.TabIndex = 1
        Me.SplitContainer1.TabStop = False
        '
        'lbxFiles
        '
        Me.lbxFiles.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbxFiles.FormattingEnabled = True
        Me.lbxFiles.ItemHeight = 24
        Me.lbxFiles.Location = New System.Drawing.Point(0, 0)
        Me.lbxFiles.Margin = New System.Windows.Forms.Padding(6)
        Me.lbxFiles.Name = "lbxFiles"
        Me.lbxFiles.Size = New System.Drawing.Size(345, 213)
        Me.lbxFiles.TabIndex = 0
        '
        'lbxShowList
        '
        Me.lbxShowList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbxShowList.FormattingEnabled = True
        Me.lbxShowList.ItemHeight = 24
        Me.lbxShowList.Location = New System.Drawing.Point(0, 0)
        Me.lbxShowList.Name = "lbxShowList"
        Me.lbxShowList.Size = New System.Drawing.Size(345, 298)
        Me.lbxShowList.TabIndex = 0
        Me.lbxShowList.TabStop = False
        '
        'MainWMP
        '
        Me.MainWMP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MainWMP.Enabled = True
        Me.MainWMP.Location = New System.Drawing.Point(0, 0)
        Me.MainWMP.Name = "MainWMP"
        Me.MainWMP.OcxState = CType(resources.GetObject("MainWMP.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MainWMP.Size = New System.Drawing.Size(2222, 1002)
        Me.MainWMP.TabIndex = 1
        Me.MainWMP.TabStop = False
        Me.MainWMP.UseWaitCursor = True
        '
        'pbxBlanker
        '
        Me.pbxBlanker.Location = New System.Drawing.Point(837, 234)
        Me.pbxBlanker.Name = "pbxBlanker"
        Me.pbxBlanker.Size = New System.Drawing.Size(563, 535)
        Me.pbxBlanker.TabIndex = 5
        Me.pbxBlanker.TabStop = False
        Me.pbxBlanker.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Black
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(2222, 1002)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'ToolStrip1
        '
        Me.ToolStrip1.AutoSize = False
        Me.ToolStrip1.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(28, 28)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2, Me.ToolStripSeparator3, Me.ToolStripSeparator4, Me.ToolStripLabel1, Me.tsbAddMovies, Me.TSBRemoveMovies, Me.ToolStripButton8, Me.ToolStripSeparator6, Me.tsbClear, Me.ToolStripSeparator5, Me.ToolStripSeparator7, Me.ToolStripButton10, Me.ToolStripSeparator8, Me.ToolStripButton11, Me.ToolStripButton12, Me.ToolStripSeparator9, Me.ToolStripButton13, Me.ToolStripButton15, Me.ToolStripSeparator10, Me.ToolStripComboBox1, Me.ToolStripButton16, Me.ToolStripButton3, Me.ToolStripTextBox1, Me.ToolStripButton9, Me.btnChooseRandom, Me.ToolStripButton5, Me.tsbFullscreen, Me.toolStripSeparator, Me.toolStripSeparator11, Me.showButtons})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(2603, 90)
        Me.ToolStrip1.Stretch = True
        Me.ToolStrip1.TabIndex = 16
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(32, 87)
        Me.ToolStripButton1.Text = "tsbFilter"
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(32, 87)
        Me.ToolStripButton2.Text = "tsbShuffle"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 90)
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 90)
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(45, 87)
        Me.ToolStripLabel1.Text = "tsl1"
        '
        'tsbAddMovies
        '
        Me.tsbAddMovies.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbAddMovies.Name = "tsbAddMovies"
        Me.tsbAddMovies.Size = New System.Drawing.Size(121, 87)
        Me.tsbAddMovies.Text = "AddMovies"
        Me.tsbAddMovies.ToolTipText = "AddAllMoviesToCollection"
        '
        'TSBRemoveMovies
        '
        Me.TSBRemoveMovies.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.TSBRemoveMovies.Image = CType(resources.GetObject("TSBRemoveMovies.Image"), System.Drawing.Image)
        Me.TSBRemoveMovies.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TSBRemoveMovies.Name = "TSBRemoveMovies"
        Me.TSBRemoveMovies.Size = New System.Drawing.Size(185, 87)
        Me.TSBRemoveMovies.Text = "tsbRemoveMovies"
        Me.TSBRemoveMovies.ToolTipText = "RemoveMovies"
        '
        'ToolStripButton8
        '
        Me.ToolStripButton8.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton8.Image = CType(resources.GetObject("ToolStripButton8.Image"), System.Drawing.Image)
        Me.ToolStripButton8.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton8.Name = "ToolStripButton8"
        Me.ToolStripButton8.Size = New System.Drawing.Size(152, 87)
        Me.ToolStripButton8.Text = "AddAllPictures"
        Me.ToolStripButton8.ToolTipText = "AddAllPictures"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(6, 90)
        '
        'tsbClear
        '
        Me.tsbClear.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbClear.Image = CType(resources.GetObject("tsbClear.Image"), System.Drawing.Image)
        Me.tsbClear.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbClear.Name = "tsbClear"
        Me.tsbClear.Size = New System.Drawing.Size(92, 87)
        Me.tsbClear.Text = "tsbClear"
        Me.tsbClear.ToolTipText = "Clear Collection"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(6, 90)
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        Me.ToolStripSeparator7.Size = New System.Drawing.Size(6, 90)
        '
        'ToolStripButton10
        '
        Me.ToolStripButton10.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton10.Image = CType(resources.GetObject("ToolStripButton10.Image"), System.Drawing.Image)
        Me.ToolStripButton10.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton10.Name = "ToolStripButton10"
        Me.ToolStripButton10.Size = New System.Drawing.Size(114, 87)
        Me.ToolStripButton10.Text = "Duplicates"
        Me.ToolStripButton10.ToolTipText = "Duplicates"
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        Me.ToolStripSeparator8.Size = New System.Drawing.Size(6, 90)
        '
        'ToolStripButton11
        '
        Me.ToolStripButton11.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton11.Image = CType(resources.GetObject("ToolStripButton11.Image"), System.Drawing.Image)
        Me.ToolStripButton11.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton11.Name = "ToolStripButton11"
        Me.ToolStripButton11.Size = New System.Drawing.Size(95, 87)
        Me.ToolStripButton11.Text = "StoreList"
        Me.ToolStripButton11.ToolTipText = "Store List"
        '
        'ToolStripButton12
        '
        Me.ToolStripButton12.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton12.Image = CType(resources.GetObject("ToolStripButton12.Image"), System.Drawing.Image)
        Me.ToolStripButton12.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton12.Name = "ToolStripButton12"
        Me.ToolStripButton12.Size = New System.Drawing.Size(93, 87)
        Me.ToolStripButton12.Text = "LoadList"
        Me.ToolStripButton12.ToolTipText = "Load list"
        '
        'ToolStripSeparator9
        '
        Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
        Me.ToolStripSeparator9.Size = New System.Drawing.Size(6, 90)
        '
        'ToolStripButton13
        '
        Me.ToolStripButton13.Items.AddRange(New Object() {"Original", "Random", "Name", "PathName", "Time", "Length", "Type"})
        Me.ToolStripButton13.Name = "ToolStripButton13"
        Me.ToolStripButton13.Size = New System.Drawing.Size(75, 90)
        Me.ToolStripButton13.Text = "Order"
        Me.ToolStripButton13.ToolTipText = "Change Order"
        '
        'ToolStripButton15
        '
        Me.ToolStripButton15.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton15.Image = CType(resources.GetObject("ToolStripButton15.Image"), System.Drawing.Image)
        Me.ToolStripButton15.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton15.Name = "ToolStripButton15"
        Me.ToolStripButton15.Size = New System.Drawing.Size(111, 87)
        Me.ToolStripButton15.Text = "SlideShow"
        Me.ToolStripButton15.ToolTipText = "ToggleSlideShow"
        '
        'ToolStripSeparator10
        '
        Me.ToolStripSeparator10.Name = "ToolStripSeparator10"
        Me.ToolStripSeparator10.Size = New System.Drawing.Size(6, 90)
        '
        'ToolStripComboBox1
        '
        Me.ToolStripComboBox1.Items.AddRange(New Object() {"50", "400", "1000", "2000", "9000"})
        Me.ToolStripComboBox1.Name = "ToolStripComboBox1"
        Me.ToolStripComboBox1.Size = New System.Drawing.Size(121, 90)
        '
        'ToolStripButton16
        '
        Me.ToolStripButton16.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton16.Image = CType(resources.GetObject("ToolStripButton16.Image"), System.Drawing.Image)
        Me.ToolStripButton16.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton16.Name = "ToolStripButton16"
        Me.ToolStripButton16.Size = New System.Drawing.Size(77, 87)
        Me.ToolStripButton16.Text = "Rotate"
        Me.ToolStripButton16.ToolTipText = "Rotate 90 CCW"
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton3.Image = CType(resources.GetObject("ToolStripButton3.Image"), System.Drawing.Image)
        Me.ToolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.Size = New System.Drawing.Size(169, 87)
        Me.ToolStripButton3.Text = "ToolStripButton3"
        '
        'ToolStripTextBox1
        '
        Me.ToolStripTextBox1.Name = "ToolStripTextBox1"
        Me.ToolStripTextBox1.Size = New System.Drawing.Size(100, 90)
        '
        'ToolStripButton9
        '
        Me.ToolStripButton9.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton9.Image = CType(resources.GetObject("ToolStripButton9.Image"), System.Drawing.Image)
        Me.ToolStripButton9.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton9.Name = "ToolStripButton9"
        Me.ToolStripButton9.Size = New System.Drawing.Size(196, 87)
        Me.ToolStripButton9.Text = "Random Start Point"
        Me.ToolStripButton9.ToolTipText = "Random Start Point Toggle"
        '
        'btnChooseRandom
        '
        Me.btnChooseRandom.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnChooseRandom.Image = CType(resources.GetObject("btnChooseRandom.Image"), System.Drawing.Image)
        Me.btnChooseRandom.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnChooseRandom.Name = "btnChooseRandom"
        Me.btnChooseRandom.Size = New System.Drawing.Size(170, 87)
        Me.btnChooseRandom.Text = "Choose Random"
        '
        'ToolStripButton5
        '
        Me.ToolStripButton5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButton5.Image = CType(resources.GetObject("ToolStripButton5.Image"), System.Drawing.Image)
        Me.ToolStripButton5.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton5.Name = "ToolStripButton5"
        Me.ToolStripButton5.Size = New System.Drawing.Size(79, 87)
        Me.ToolStripButton5.Text = "Search"
        Me.ToolStripButton5.ToolTipText = "Base Collection on Search"
        '
        'tsbFullscreen
        '
        Me.tsbFullscreen.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsbFullscreen.Image = CType(resources.GetObject("tsbFullscreen.Image"), System.Drawing.Image)
        Me.tsbFullscreen.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbFullscreen.Name = "tsbFullscreen"
        Me.tsbFullscreen.Size = New System.Drawing.Size(139, 87)
        Me.tsbFullscreen.Text = "tsbFullScreen"
        Me.tsbFullscreen.ToolTipText = "OpenFullScreen"
        '
        'toolStripSeparator
        '
        Me.toolStripSeparator.Name = "toolStripSeparator"
        Me.toolStripSeparator.Size = New System.Drawing.Size(6, 90)
        '
        'toolStripSeparator11
        '
        Me.toolStripSeparator11.Name = "toolStripSeparator11"
        Me.toolStripSeparator11.Size = New System.Drawing.Size(6, 90)
        '
        'showButtons
        '
        Me.showButtons.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.showButtons.Image = CType(resources.GetObject("showButtons.Image"), System.Drawing.Image)
        Me.showButtons.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.showButtons.Name = "showButtons"
        Me.showButtons.Size = New System.Drawing.Size(144, 87)
        Me.showButtons.Text = "Show Buttons"
        '
        'tmrUpdateFileList
        '
        Me.tmrUpdateFileList.Enabled = True
        Me.tmrUpdateFileList.Interval = 10
        '
        'tmrPicLoad
        '
        '
        'tmrJumpVideo
        '
        '
        'tmrInitialise
        '
        Me.tmrInitialise.Enabled = True
        Me.tmrInitialise.Interval = 500
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "msl"
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Filter = "MV files|*.msl|All files|*.*"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "msl"
        Me.SaveFileDialog1.Filter = "MV files|*.msl|All files|*.*"
        '
        'tmrSlideShow
        '
        Me.tmrSlideShow.Interval = 750
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'FileSystemWatcher1
        '
        Me.FileSystemWatcher1.EnableRaisingEvents = True
        Me.FileSystemWatcher1.SynchronizingObject = Me
        '
        'tmrLoadLastFolder
        '
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(2603, 1171)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(6)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Metavisua"
        Me.TransparencyKey = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.ctrFilesandPics.Panel1.ResumeLayout(False)
        Me.ctrFilesandPics.Panel2.ResumeLayout(False)
        CType(Me.ctrFilesandPics, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ctrFilesandPics.ResumeLayout(False)
        Me.ctrTreeandFiles.Panel1.ResumeLayout(False)
        Me.ctrTreeandFiles.Panel2.ResumeLayout(False)
        CType(Me.ctrTreeandFiles, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ctrTreeandFiles.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.MainWMP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbxBlanker, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents tssFolderInfo As ToolStripStatusLabel
    Friend WithEvents tssMoveCopy As ToolStripStatusLabel
    Friend WithEvents tssFilter As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel4 As ToolStripStatusLabel
    Friend WithEvents stripLastDone As ToolStripStatusLabel
    Friend WithEvents ToolStripProgressBar1 As ToolStripProgressBar
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ViewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ButtonsToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SlideShowToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents VideoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents ctrFilesandPics As SplitContainer
    Friend WithEvents ctrTreeandFiles As SplitContainer
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents ToolStripButton2 As ToolStripButton


    Friend WithEvents NewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents SaveAsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FullScreenToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TrueSizeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SwitchModesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AssignNextToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents tmrUpdateFileList As Timer
    Friend WithEvents tmrPicLoad As Timer
    Friend WithEvents tmrJumpVideo As Timer
    Friend WithEvents tmrInitialise As Timer
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents tsbFullscreen As ToolStripButton
    Friend WithEvents tsbClear As ToolStripButton
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents ToolStripLabel1 As ToolStripLabel
    Friend WithEvents tsbAddMovies As ToolStripButton
    Friend WithEvents ToolStripButton8 As ToolStripButton
    Friend WithEvents tmrCheckFolders As Timer
    Friend WithEvents ToolStripButton9 As ToolStripButton
    Friend WithEvents tsslblFiles As ToolStripStatusLabel
    Friend WithEvents ToolStripButton10 As ToolStripButton
    Friend WithEvents tsslblFilter As ToolStripStatusLabel
    Friend WithEvents FileToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ListsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AddAllPicturesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CurrentOnlyToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents IncludingSubfoldersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AddAllVideosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CurrentOnlyToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents IncludingSubsetsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents ToolStripButton11 As ToolStripButton
    Friend WithEvents ToolStripButton12 As ToolStripButton
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents lbxFiles As ListBox
    Friend WithEvents lbxShowList As ListBox
    Friend WithEvents LoadToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ShowListToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents SaveListToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripButton13 As ToolStripComboBox
    Friend WithEvents TSBRemoveMovies As ToolStripButton
    Friend WithEvents ToolStripButton15 As ToolStripButton
    Friend WithEvents ToolStripComboBox1 As ToolStripComboBox
    Friend WithEvents tmrSlideShow As Timer
    Friend WithEvents ToolStripButton16 As ToolStripButton
    Friend WithEvents ToolStripSeparator6 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator7 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator8 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator9 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator10 As ToolStripSeparator
    Friend WithEvents ToolStripButton3 As ToolStripButton
    Friend WithEvents ToolStripButton5 As ToolStripButton
    Friend WithEvents ToolStripTextBox1 As ToolStripTextBox
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents tsslblRandom As ToolStripStatusLabel
    Friend WithEvents tsslblLastfile As ToolStripStatusLabel
    Friend WithEvents tsslblShowfile As ToolStripStatusLabel
    Friend WithEvents tsslblSTART As ToolStripStatusLabel
    Friend WithEvents tsslblSPEED As ToolStripStatusLabel
    Friend WithEvents tsslblZOOM As ToolStripStatusLabel
    Friend WithEvents toolStripSeparator As ToolStripSeparator
    Friend WithEvents toolStripSeparator11 As ToolStripSeparator
    Friend WithEvents SlideshowToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents MainWMP As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents showButtons As ToolStripButton
    Friend WithEvents pbxBlanker As PictureBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents FileSystemWatcher1 As IO.FileSystemWatcher
    Friend WithEvents AddAllFilesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CurrentOnlyToolStripMenuItem2 As ToolStripMenuItem
    Friend WithEvents AllSubFoldersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AddPicturesAndVideosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CurrentOnlyToolStripMenuItem3 As ToolStripMenuItem
    Friend WithEvents AllSubfoldersToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents tmrLoadLastFolder As Timer
    Friend WithEvents SlowToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NormalToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FastToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents VideoToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents SpeedToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SlowToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents NormalToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents FastToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents RandomStartToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DateSSL As ToolStripStatusLabel
    Friend WithEvents tvMain2 As Controls.FileSystemTree
    Friend WithEvents btnChooseRandom As ToolStripButton
End Class
