<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUnMap_model
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
        Me.lstNewmtp = New System.Windows.Forms.ListBox
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.lstNewtip = New System.Windows.Forms.ListBox
        Me.lstUnmtp = New System.Windows.Forms.ListBox
        Me.ListBox2 = New System.Windows.Forms.ListBox
        Me.lstUntip = New System.Windows.Forms.ListBox
        Me.SuspendLayout()
        '
        'lstNewmtp
        '
        Me.lstNewmtp.FormattingEnabled = True
        Me.lstNewmtp.Location = New System.Drawing.Point(38, 12)
        Me.lstNewmtp.Name = "lstNewmtp"
        Me.lstNewmtp.Size = New System.Drawing.Size(120, 95)
        Me.lstNewmtp.TabIndex = 0
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(147, 12)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(120, 95)
        Me.ListBox1.TabIndex = 1
        '
        'lstNewtip
        '
        Me.lstNewtip.FormattingEnabled = True
        Me.lstNewtip.Location = New System.Drawing.Point(249, 12)
        Me.lstNewtip.Name = "lstNewtip"
        Me.lstNewtip.Size = New System.Drawing.Size(120, 95)
        Me.lstNewtip.TabIndex = 2
        '
        'lstUnmtp
        '
        Me.lstUnmtp.FormattingEnabled = True
        Me.lstUnmtp.Location = New System.Drawing.Point(58, 136)
        Me.lstUnmtp.Name = "lstUnmtp"
        Me.lstUnmtp.Size = New System.Drawing.Size(120, 95)
        Me.lstUnmtp.TabIndex = 3
        '
        'ListBox2
        '
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.Location = New System.Drawing.Point(113, 136)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(120, 95)
        Me.ListBox2.TabIndex = 4
        '
        'lstUntip
        '
        Me.lstUntip.FormattingEnabled = True
        Me.lstUntip.Location = New System.Drawing.Point(219, 136)
        Me.lstUntip.Name = "lstUntip"
        Me.lstUntip.Size = New System.Drawing.Size(120, 95)
        Me.lstUntip.TabIndex = 5
        '
        'frmUnMap_model
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(509, 266)
        Me.Controls.Add(Me.lstUntip)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.lstUnmtp)
        Me.Controls.Add(Me.lstNewtip)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.lstNewmtp)
        Me.Name = "frmUnMap_model"
        Me.Text = "frmUnMap_model"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lstNewmtp As System.Windows.Forms.ListBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents lstNewtip As System.Windows.Forms.ListBox
    Friend WithEvents lstUnmtp As System.Windows.Forms.ListBox
    Friend WithEvents ListBox2 As System.Windows.Forms.ListBox
    Friend WithEvents lstUntip As System.Windows.Forms.ListBox
End Class
