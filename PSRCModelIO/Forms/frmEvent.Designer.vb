<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEvent
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
        Me.LabelprjR = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'LabelprjR
        '
        Me.LabelprjR.AutoSize = True
        Me.LabelprjR.Location = New System.Drawing.Point(0, 0)
        Me.LabelprjR.Name = "LabelprjR"
        Me.LabelprjR.Size = New System.Drawing.Size(39, 13)
        Me.LabelprjR.TabIndex = 0
        Me.LabelprjR.Text = "Label1"
        '
        'frmEvent
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.LabelprjR)
        Me.Name = "frmEvent"
        Me.Text = "frmEvent"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelprjR As System.Windows.Forms.Label
End Class
