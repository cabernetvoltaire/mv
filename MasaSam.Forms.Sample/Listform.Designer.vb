<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Listform
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
        Me.FileBox1 = New MasaSam.Forms.Sample.FileBox()
        Me.SuspendLayout()
        '
        'FileBox1
        '
        Me.FileBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FileBox1.FolderPath = Nothing
        Me.FileBox1.List = Nothing
        Me.FileBox1.Location = New System.Drawing.Point(0, 0)
        Me.FileBox1.Name = "FileBox1"
        Me.FileBox1.Size = New System.Drawing.Size(800, 450)
        Me.FileBox1.TabIndex = 0
        '
        'Listform
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.FileBox1)
        Me.KeyPreview = True
        Me.Name = "Listform"
        Me.Text = "ShowList"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents FileBox1 As FileBox
End Class
