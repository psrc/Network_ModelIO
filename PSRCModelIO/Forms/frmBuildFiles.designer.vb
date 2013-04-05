<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBuildFiles
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
        Me.txtYear = New System.Windows.Forms.TextBox
        Me.txtFile = New System.Windows.Forms.TextBox
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.btnOpenFile = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnContinue = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.lblStatus = New System.Windows.Forms.Label
        Me.txtSchema = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.chkNetwork = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkTransit = New System.Windows.Forms.CheckBox
        Me.chkTolls = New System.Windows.Forms.CheckBox
        Me.chkTurns = New System.Windows.Forms.CheckBox
        Me.chkParkRides = New System.Windows.Forms.CheckBox
        Me.rdoOldTAZ = New System.Windows.Forms.RadioButton
        Me.rdoNewTAZ = New System.Windows.Forms.RadioButton
        Me.Label5 = New System.Windows.Forms.Label
        Me.chkExportShapefiles = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtOffset
        '
        Me.txtOffset.Location = New System.Drawing.Point(99, 90)
        Me.txtOffset.Name = "txtOffset"
        Me.txtOffset.Size = New System.Drawing.Size(100, 20)
        Me.txtOffset.TabIndex = 0
        Me.txtOffset.Text = "1200"
        '
        'txtYear
        '
        Me.txtYear.Location = New System.Drawing.Point(99, 55)
        Me.txtYear.Name = "txtYear"
        Me.txtYear.ShortcutsEnabled = False
        Me.txtYear.Size = New System.Drawing.Size(100, 20)
        Me.txtYear.TabIndex = 1
        Me.txtYear.Text = "2010"
        '
        'txtFile
        '
        Me.txtFile.Location = New System.Drawing.Point(12, 21)
        Me.txtFile.Name = "txtFile"
        Me.txtFile.Size = New System.Drawing.Size(187, 20)
        Me.txtFile.TabIndex = 2
        '
        'btnOpenFile
        '
        Me.btnOpenFile.Location = New System.Drawing.Point(205, 21)
        Me.btnOpenFile.Name = "btnOpenFile"
        Me.btnOpenFile.Size = New System.Drawing.Size(75, 23)
        Me.btnOpenFile.TabIndex = 3
        Me.btnOpenFile.Text = "Browse"
        Me.btnOpenFile.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(278, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Please select the output location and naming convention:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Model Run Year"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 93)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Node Offset:"
        '
        'btnContinue
        '
        Me.btnContinue.Location = New System.Drawing.Point(36, 251)
        Me.btnContinue.Name = "btnContinue"
        Me.btnContinue.Size = New System.Drawing.Size(75, 23)
        Me.btnContinue.TabIndex = 8
        Me.btnContinue.Text = "Continue"
        Me.btnContinue.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(158, 251)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 9
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(12, 292)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(39, 13)
        Me.lblStatus.TabIndex = 10
        Me.lblStatus.Text = "Label4"
        '
        'txtSchema
        '
        Me.txtSchema.Location = New System.Drawing.Point(99, 116)
        Me.txtSchema.Name = "txtSchema"
        Me.txtSchema.Size = New System.Drawing.Size(100, 20)
        Me.txtSchema.TabIndex = 11
        Me.txtSchema.Text = "sde.SDE."
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(2, 119)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Schema"
        '
        'chkNetwork
        '
        Me.chkNetwork.AutoSize = True
        Me.chkNetwork.Location = New System.Drawing.Point(19, 19)
        Me.chkNetwork.Name = "chkNetwork"
        Me.chkNetwork.Size = New System.Drawing.Size(116, 17)
        Me.chkNetwork.TabIndex = 13
        Me.chkNetwork.Text = "Build Network Files"
        Me.chkNetwork.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkExportShapefiles)
        Me.GroupBox1.Controls.Add(Me.chkTransit)
        Me.GroupBox1.Controls.Add(Me.chkTolls)
        Me.GroupBox1.Controls.Add(Me.chkTurns)
        Me.GroupBox1.Controls.Add(Me.chkParkRides)
        Me.GroupBox1.Controls.Add(Me.chkNetwork)
        Me.GroupBox1.Location = New System.Drawing.Point(17, 145)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(271, 106)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "GroupBox1"
        '
        'chkTransit
        '
        Me.chkTransit.AutoSize = True
        Me.chkTransit.Location = New System.Drawing.Point(135, 42)
        Me.chkTransit.Name = "chkTransit"
        Me.chkTransit.Size = New System.Drawing.Size(108, 17)
        Me.chkTransit.TabIndex = 17
        Me.chkTransit.Text = "Build Transit Files"
        Me.chkTransit.UseVisualStyleBackColor = True
        '
        'chkTolls
        '
        Me.chkTolls.AutoSize = True
        Me.chkTolls.Location = New System.Drawing.Point(135, 19)
        Me.chkTolls.Name = "chkTolls"
        Me.chkTolls.Size = New System.Drawing.Size(93, 17)
        Me.chkTolls.TabIndex = 16
        Me.chkTolls.Text = "Build Toll Files"
        Me.chkTolls.UseVisualStyleBackColor = True
        '
        'chkTurns
        '
        Me.chkTurns.AutoSize = True
        Me.chkTurns.Location = New System.Drawing.Point(19, 65)
        Me.chkTurns.Name = "chkTurns"
        Me.chkTurns.Size = New System.Drawing.Size(98, 17)
        Me.chkTurns.TabIndex = 15
        Me.chkTurns.Text = "Build Turn Files"
        Me.chkTurns.UseVisualStyleBackColor = True
        '
        'chkParkRides
        '
        Me.chkParkRides.AutoSize = True
        Me.chkParkRides.Location = New System.Drawing.Point(19, 42)
        Me.chkParkRides.Name = "chkParkRides"
        Me.chkParkRides.Size = New System.Drawing.Size(144, 17)
        Me.chkParkRides.TabIndex = 14
        Me.chkParkRides.Text = "Build Park and Ride Files"
        Me.chkParkRides.UseVisualStyleBackColor = True
        '
        'rdoOldTAZ
        '
        Me.rdoOldTAZ.AutoSize = True
        Me.rdoOldTAZ.Checked = True
        Me.rdoOldTAZ.Location = New System.Drawing.Point(36, 343)
        Me.rdoOldTAZ.Name = "rdoOldTAZ"
        Me.rdoOldTAZ.Size = New System.Drawing.Size(68, 17)
        Me.rdoOldTAZ.TabIndex = 16
        Me.rdoOldTAZ.TabStop = True
        Me.rdoOldTAZ.Text = "Old TAZ "
        Me.rdoOldTAZ.UseVisualStyleBackColor = True
        '
        'rdoNewTAZ
        '
        Me.rdoNewTAZ.AutoSize = True
        Me.rdoNewTAZ.Location = New System.Drawing.Point(36, 366)
        Me.rdoNewTAZ.Name = "rdoNewTAZ"
        Me.rdoNewTAZ.Size = New System.Drawing.Size(71, 17)
        Me.rdoNewTAZ.TabIndex = 17
        Me.rdoNewTAZ.Text = "New TAZ"
        Me.rdoNewTAZ.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(14, 327)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(116, 13)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Choose TAZ Structure:"
        '
        'chkExportShapefiles
        '
        Me.chkExportShapefiles.AutoSize = True
        Me.chkExportShapefiles.Location = New System.Drawing.Point(135, 65)
        Me.chkExportShapefiles.Name = "chkExportShapefiles"
        Me.chkExportShapefiles.Size = New System.Drawing.Size(108, 17)
        Me.chkExportShapefiles.TabIndex = 18
        Me.chkExportShapefiles.Text = "Export Shapefiles"
        Me.chkExportShapefiles.UseVisualStyleBackColor = True
        '
        'frmBuildFiles
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(365, 427)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.rdoNewTAZ)
        Me.Controls.Add(Me.rdoOldTAZ)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtSchema)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnContinue)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOpenFile)
        Me.Controls.Add(Me.txtFile)
        Me.Controls.Add(Me.txtYear)
        Me.Controls.Add(Me.txtOffset)
        Me.Name = "frmBuildFiles"
        Me.Text = "frmBuildFiles"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtOffset As System.Windows.Forms.TextBox
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnOpenFile As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnContinue As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents txtSchema As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents chkNetwork As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkTransit As System.Windows.Forms.CheckBox
    Friend WithEvents chkTolls As System.Windows.Forms.CheckBox
    Friend WithEvents chkTurns As System.Windows.Forms.CheckBox
    Friend WithEvents chkParkRides As System.Windows.Forms.CheckBox
    Friend WithEvents rdoOldTAZ As System.Windows.Forms.RadioButton
    Friend WithEvents rdoNewTAZ As System.Windows.Forms.RadioButton
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents chkExportShapefiles As System.Windows.Forms.CheckBox
End Class
