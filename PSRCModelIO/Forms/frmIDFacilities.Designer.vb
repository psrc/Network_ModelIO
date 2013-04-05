<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmIDFacilities
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
        Me.chkboxEvents = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'chkboxEvents
        '
        Me.chkboxEvents.AutoSize = True
        Me.chkboxEvents.Location = New System.Drawing.Point(21, 21)
        Me.chkboxEvents.Name = "chkboxEvents"
        Me.chkboxEvents.Size = New System.Drawing.Size(81, 17)
        Me.chkboxEvents.TabIndex = 0
        Me.chkboxEvents.Text = "CheckBox1"
        Me.chkboxEvents.UseVisualStyleBackColor = True
        '
        'frmIDFacilities
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.chkboxEvents)
        Me.Name = "frmIDFacilities"
        Me.Text = "frmIDFacilities"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkboxEvents As System.Windows.Forms.CheckBox
End Class
