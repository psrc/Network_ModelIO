<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMapped
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
        Me.lstMmtp = New System.Windows.Forms.ListBox
        Me.lstBoxM = New System.Windows.Forms.ListBox
        Me.lstMtip = New System.Windows.Forms.ListBox
        Me.ListBoxmtpE = New System.Windows.Forms.ListBox
        Me.ListBoxM = New System.Windows.Forms.ListBox
        Me.ListBoxtipE = New System.Windows.Forms.ListBox
        Me.SuspendLayout()
        '
        'lstMmtp
        '
        Me.lstMmtp.FormattingEnabled = True
        Me.lstMmtp.Location = New System.Drawing.Point(2, 25)
        Me.lstMmtp.Name = "lstMmtp"
        Me.lstMmtp.Size = New System.Drawing.Size(120, 95)
        Me.lstMmtp.TabIndex = 0
        '
        'lstBoxM
        '
        Me.lstBoxM.FormattingEnabled = True
        Me.lstBoxM.Location = New System.Drawing.Point(72, 25)
        Me.lstBoxM.Name = "lstBoxM"
        Me.lstBoxM.Size = New System.Drawing.Size(120, 95)
        Me.lstBoxM.TabIndex = 1
        '
        'lstMtip
        '
        Me.lstMtip.FormattingEnabled = True
        Me.lstMtip.Location = New System.Drawing.Point(160, 25)
        Me.lstMtip.Name = "lstMtip"
        Me.lstMtip.Size = New System.Drawing.Size(120, 95)
        Me.lstMtip.TabIndex = 2
        '
        'ListBoxmtpE
        '
        Me.ListBoxmtpE.FormattingEnabled = True
        Me.ListBoxmtpE.Location = New System.Drawing.Point(2, 138)
        Me.ListBoxmtpE.Name = "ListBoxmtpE"
        Me.ListBoxmtpE.Size = New System.Drawing.Size(120, 95)
        Me.ListBoxmtpE.TabIndex = 3
        '
        'ListBoxM
        '
        Me.ListBoxM.FormattingEnabled = True
        Me.ListBoxM.Location = New System.Drawing.Point(72, 138)
        Me.ListBoxM.Name = "ListBoxM"
        Me.ListBoxM.Size = New System.Drawing.Size(120, 95)
        Me.ListBoxM.TabIndex = 4
        '
        'ListBoxtipE
        '
        Me.ListBoxtipE.FormattingEnabled = True
        Me.ListBoxtipE.Location = New System.Drawing.Point(170, 138)
        Me.ListBoxtipE.Name = "ListBoxtipE"
        Me.ListBoxtipE.Size = New System.Drawing.Size(120, 95)
        Me.ListBoxtipE.TabIndex = 5
        '
        'frmMapped
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(395, 428)
        Me.Controls.Add(Me.ListBoxtipE)
        Me.Controls.Add(Me.ListBoxM)
        Me.Controls.Add(Me.ListBoxmtpE)
        Me.Controls.Add(Me.lstMtip)
        Me.Controls.Add(Me.lstBoxM)
        Me.Controls.Add(Me.lstMmtp)
        Me.Name = "frmMapped"
        Me.Text = "frmMapped"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lstMmtp As System.Windows.Forms.ListBox
    Friend WithEvents lstBoxM As System.Windows.Forms.ListBox
    Friend WithEvents lstMtip As System.Windows.Forms.ListBox
    Friend WithEvents ListBoxmtpE As System.Windows.Forms.ListBox
    Friend WithEvents ListBoxM As System.Windows.Forms.ListBox
    Friend WithEvents ListBoxtipE As System.Windows.Forms.ListBox
End Class
