<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOutNetSelect
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
        Me.lblStatus = New System.Windows.Forms.Label
        Me.lblFile = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(50, 58)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(39, 13)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Label1"
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Location = New System.Drawing.Point(50, 34)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(39, 13)
        Me.lblFile.TabIndex = 1
        Me.lblFile.Text = "Label1"
        '
        'frmOutNetSelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.lblStatus)
        Me.Name = "frmOutNetSelect"
        Me.Text = "frmOutNetSelect"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lblFile As System.Windows.Forms.Label
End Class
