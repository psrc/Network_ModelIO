<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPath
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
        Me.Text1 = New System.Windows.Forms.TextBox
        Me.cmdTable = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.filePath = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Text1
        '
        Me.Text1.Location = New System.Drawing.Point(35, 62)
        Me.Text1.Name = "Text1"
        Me.Text1.Size = New System.Drawing.Size(142, 20)
        Me.Text1.TabIndex = 0
        '
        'cmdTable
        '
        Me.cmdTable.Location = New System.Drawing.Point(35, 129)
        Me.cmdTable.Name = "cmdTable"
        Me.cmdTable.Size = New System.Drawing.Size(75, 23)
        Me.cmdTable.TabIndex = 1
        Me.cmdTable.Text = "Open Table"
        Me.cmdTable.UseVisualStyleBackColor = True
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(141, 129)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'filePath
        '
        Me.filePath.Location = New System.Drawing.Point(205, 60)
        Me.filePath.Name = "filePath"
        Me.filePath.Size = New System.Drawing.Size(75, 23)
        Me.filePath.TabIndex = 3
        Me.filePath.Text = "path"
        Me.filePath.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(32, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(149, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Output Filename and Location"
        '
        'frmPath
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.filePath)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdTable)
        Me.Controls.Add(Me.Text1)
        Me.Name = "frmPath"
        Me.Text = "frmPath"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Text1 As System.Windows.Forms.TextBox
    Friend WithEvents cmdTable As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents filePath As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
