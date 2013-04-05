<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBuildFiles2
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
        Me.txtOffset = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'txtOffset
        '
        Me.txtOffset.Location = New System.Drawing.Point(0, 0)
        Me.txtOffset.Name = "txtOffset"
        Me.txtOffset.Size = New System.Drawing.Size(100, 20)
        Me.txtOffset.TabIndex = 0
        '
        'frmBuildFiles
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.txtOffset)
        Me.Name = "frmBuildFiles"
        Me.Text = "frmBuildFiles"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtOffset As System.Windows.Forms.TextBox
End Class
