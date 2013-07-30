<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmModelQuery
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
        Me.components = New System.ComponentModel.Container()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtSchema = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnOpenFBD1 = New System.Windows.Forms.Button()
        Me.txtOutPutDir = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkFT = New System.Windows.Forms.CheckBox()
        Me.cmbFC = New System.Windows.Forms.CheckedListBox()
        Me.ckcFC = New System.Windows.Forms.CheckBox()
        Me.ckcActive = New System.Windows.Forms.CheckBox()
        Me.ckcPseudo = New System.Windows.Forms.CheckBox()
        Me.chkDisolveOnly = New System.Windows.Forms.CheckBox()
        Me.txtOffset = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label7 = New System.Windows.Forms.Label()
        Me.rdoNewTAZ = New System.Windows.Forms.RadioButton()
        Me.rdoOldTAZ = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtSchema)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.btnOpenFBD1)
        Me.GroupBox1.Controls.Add(Me.txtOutPutDir)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.txtOffset)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtDesc)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtTitle)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtYear)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(21, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(391, 596)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Enter Modeling Scenario Information"
        '
        'txtSchema
        '
        Me.txtSchema.Location = New System.Drawing.Point(122, 189)
        Me.txtSchema.Name = "txtSchema"
        Me.txtSchema.Size = New System.Drawing.Size(162, 24)
        Me.txtSchema.TabIndex = 13
        Me.txtSchema.Text = "sde.SDE."
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(28, 189)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(67, 18)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Schema:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(28, 260)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(71, 18)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Log Path:"
        '
        'btnOpenFBD1
        '
        Me.btnOpenFBD1.Location = New System.Drawing.Point(297, 261)
        Me.btnOpenFBD1.Name = "btnOpenFBD1"
        Me.btnOpenFBD1.Size = New System.Drawing.Size(75, 23)
        Me.btnOpenFBD1.TabIndex = 10
        Me.btnOpenFBD1.Text = "Open "
        Me.btnOpenFBD1.UseVisualStyleBackColor = True
        '
        'txtOutPutDir
        '
        Me.txtOutPutDir.Location = New System.Drawing.Point(122, 260)
        Me.txtOutPutDir.Name = "txtOutPutDir"
        Me.txtOutPutDir.Size = New System.Drawing.Size(162, 24)
        Me.txtOutPutDir.TabIndex = 9
        Me.txtOutPutDir.Text = "C:\"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkFT)
        Me.GroupBox2.Controls.Add(Me.cmbFC)
        Me.GroupBox2.Controls.Add(Me.ckcFC)
        Me.GroupBox2.Controls.Add(Me.ckcActive)
        Me.GroupBox2.Controls.Add(Me.ckcPseudo)
        Me.GroupBox2.Controls.Add(Me.chkDisolveOnly)
        Me.GroupBox2.Location = New System.Drawing.Point(21, 288)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(351, 302)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Thinning Options"
        '
        'chkFT
        '
        Me.chkFT.AutoSize = True
        Me.chkFT.Location = New System.Drawing.Point(25, 51)
        Me.chkFT.Name = "chkFT"
        Me.chkFT.Size = New System.Drawing.Size(212, 22)
        Me.chkFT.TabIndex = 3
        Me.chkFT.Text = "Thin Legacy By Facility Type"
        Me.chkFT.UseVisualStyleBackColor = True
        '
        'cmbFC
        '
        Me.cmbFC.FormattingEnabled = True
        Me.cmbFC.Items.AddRange(New Object() {"Rural Interstate", "Rural Principal Arterial", "Rural Minor Arterial", "Rural Major Collector"})
        Me.cmbFC.Location = New System.Drawing.Point(101, 235)
        Me.cmbFC.Name = "cmbFC"
        Me.cmbFC.Size = New System.Drawing.Size(208, 61)
        Me.cmbFC.TabIndex = 4
        '
        'ckcFC
        '
        Me.ckcFC.Location = New System.Drawing.Point(25, 183)
        Me.ckcFC.Name = "ckcFC"
        Me.ckcFC.Size = New System.Drawing.Size(296, 46)
        Me.ckcFC.TabIndex = 3
        Me.ckcFC.Text = "Thin network by using only the selected functional classes"
        Me.ckcFC.UseVisualStyleBackColor = True
        '
        'ckcActive
        '
        Me.ckcActive.Checked = True
        Me.ckcActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckcActive.Location = New System.Drawing.Point(25, 133)
        Me.ckcActive.Name = "ckcActive"
        Me.ckcActive.Size = New System.Drawing.Size(317, 42)
        Me.ckcActive.TabIndex = 2
        Me.ckcActive.Text = "Thin networkby links  participating in the legacy 2000 EME/2 network"
        Me.ckcActive.UseVisualStyleBackColor = True
        '
        'ckcPseudo
        '
        Me.ckcPseudo.AutoSize = True
        Me.ckcPseudo.Checked = True
        Me.ckcPseudo.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckcPseudo.Location = New System.Drawing.Point(25, 23)
        Me.ckcPseudo.Name = "ckcPseudo"
        Me.ckcPseudo.Size = New System.Drawing.Size(317, 22)
        Me.ckcPseudo.TabIndex = 1
        Me.ckcPseudo.Text = "Thin pseudonodes (regardless of Junct cnt.)"
        Me.ckcPseudo.UseVisualStyleBackColor = True
        '
        'chkDisolveOnly
        '
        Me.chkDisolveOnly.Location = New System.Drawing.Point(25, 71)
        Me.chkDisolveOnly.Name = "chkDisolveOnly"
        Me.chkDisolveOnly.Size = New System.Drawing.Size(296, 56)
        Me.chkDisolveOnly.TabIndex = 0
        Me.chkDisolveOnly.Text = "Run ONLY until pseudo thin is complete (intermediate shp thinned)"
        Me.chkDisolveOnly.UseVisualStyleBackColor = True
        '
        'txtOffset
        '
        Me.txtOffset.Location = New System.Drawing.Point(122, 230)
        Me.txtOffset.Name = "txtOffset"
        Me.txtOffset.Size = New System.Drawing.Size(162, 24)
        Me.txtOffset.TabIndex = 7
        Me.txtOffset.Text = "4000"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(28, 230)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(92, 18)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Node Offset:"
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(122, 114)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(251, 64)
        Me.txtDesc.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(28, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(87, 18)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Description:"
        '
        'txtTitle
        '
        Me.txtTitle.Location = New System.Drawing.Point(85, 77)
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(288, 24)
        Me.txtTitle.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(28, 75)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Title:"
        '
        'txtYear
        '
        Me.txtYear.Location = New System.Drawing.Point(85, 40)
        Me.txtYear.Name = "txtYear"
        Me.txtYear.Size = New System.Drawing.Size(162, 24)
        Me.txtYear.TabIndex = 1
        Me.txtYear.Text = "2006"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(28, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Year:"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(61, 689)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(259, 737)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(76, 23)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.SelectedPath = "C:\"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(51, 627)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(116, 13)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "Choose TAZ Structure:"
        '
        'rdoNewTAZ
        '
        Me.rdoNewTAZ.AutoSize = True
        Me.rdoNewTAZ.Checked = True
        Me.rdoNewTAZ.Location = New System.Drawing.Point(73, 666)
        Me.rdoNewTAZ.Name = "rdoNewTAZ"
        Me.rdoNewTAZ.Size = New System.Drawing.Size(71, 17)
        Me.rdoNewTAZ.TabIndex = 20
        Me.rdoNewTAZ.TabStop = True
        Me.rdoNewTAZ.Text = "New TAZ"
        Me.rdoNewTAZ.UseVisualStyleBackColor = True
        '
        'rdoOldTAZ
        '
        Me.rdoOldTAZ.AutoSize = True
        Me.rdoOldTAZ.Location = New System.Drawing.Point(73, 643)
        Me.rdoOldTAZ.Name = "rdoOldTAZ"
        Me.rdoOldTAZ.Size = New System.Drawing.Size(68, 17)
        Me.rdoOldTAZ.TabIndex = 19
        Me.rdoOldTAZ.Text = "Old TAZ "
        Me.rdoOldTAZ.UseVisualStyleBackColor = True
        '
        'frmModelQuery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(446, 742)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.rdoNewTAZ)
        Me.Controls.Add(Me.rdoOldTAZ)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmModelQuery"
        Me.Text = "frmCreateNetwork"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtTitle As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtOffset As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ckcFC As System.Windows.Forms.CheckBox
    Friend WithEvents ckcActive As System.Windows.Forms.CheckBox
    Friend WithEvents ckcPseudo As System.Windows.Forms.CheckBox
    Friend WithEvents chkDisolveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents cmbFC As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOpenFBD1 As System.Windows.Forms.Button
    Friend WithEvents txtOutPutDir As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtSchema As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents chkFT As System.Windows.Forms.CheckBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents rdoNewTAZ As System.Windows.Forms.RadioButton
    Friend WithEvents rdoOldTAZ As System.Windows.Forms.RadioButton
End Class
