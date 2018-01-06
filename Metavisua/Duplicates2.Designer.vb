<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Duplicates2
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
        Me.lbxsorted = New System.Windows.Forms.ListBox()
        Me.lbxunique = New System.Windows.Forms.ListBox()
        Me.lbxduplicates = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'lbxsorted
        '
        Me.lbxsorted.FormattingEnabled = True
        Me.lbxsorted.ItemHeight = 24
        Me.lbxsorted.Location = New System.Drawing.Point(63, 86)
        Me.lbxsorted.Name = "lbxsorted"
        Me.lbxsorted.Size = New System.Drawing.Size(575, 1060)
        Me.lbxsorted.TabIndex = 0
        '
        'lbxunique
        '
        Me.lbxunique.FormattingEnabled = True
        Me.lbxunique.ItemHeight = 24
        Me.lbxunique.Location = New System.Drawing.Point(686, 86)
        Me.lbxunique.Name = "lbxunique"
        Me.lbxunique.Size = New System.Drawing.Size(804, 1060)
        Me.lbxunique.TabIndex = 1
        '
        'lbxduplicates
        '
        Me.lbxduplicates.FormattingEnabled = True
        Me.lbxduplicates.ItemHeight = 24
        Me.lbxduplicates.Location = New System.Drawing.Point(1520, 86)
        Me.lbxduplicates.Name = "lbxduplicates"
        Me.lbxduplicates.Size = New System.Drawing.Size(1070, 1060)
        Me.lbxduplicates.TabIndex = 2
        '
        'Duplicates2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2602, 1218)
        Me.Controls.Add(Me.lbxduplicates)
        Me.Controls.Add(Me.lbxunique)
        Me.Controls.Add(Me.lbxsorted)
        Me.Name = "Duplicates2"
        Me.Text = "Duplicates2"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lbxsorted As ListBox
    Friend WithEvents lbxunique As ListBox
    Friend WithEvents lbxduplicates As ListBox
End Class
