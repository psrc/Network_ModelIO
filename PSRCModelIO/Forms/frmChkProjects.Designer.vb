<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChkProjects
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
        Me.btnSkipProjects = New System.Windows.Forms.Button
        Me.btnOpenProj = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.rdoProjectScenario = New System.Windows.Forms.RadioButton
        Me.SuspendLayout()
        '
        'btnSkipProjects
        '
        Me.btnSkipProjects.Location = New System.Drawing.Point(110, 208)
        Me.btnSkipProjects.Name = "btnSkipProjects"
        Me.btnSkipProjects.Size = New System.Drawing.Size(92, 28)
        Me.btnSkipProjects.TabIndex = 0
        Me.btnSkipProjects.Text = "Skip Projects"
        Me.btnSkipProjects.UseVisualStyleBackColor = True
        '
        'btnOpenProj
        '
        Me.btnOpenProj.Location = New System.Drawing.Point(2, 208)
        Me.btnOpenProj.Name = "btnOpenProj"
        Me.btnOpenProj.Size = New System.Drawing.Size(92, 28)
        Me.btnOpenProj.TabIndex = 1
        Me.btnOpenProj.Text = "Open"
        Me.btnOpenProj.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(218, 208)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(92, 28)
        Me.btnExit.TabIndex = 2
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(37, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(254, 55)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Please Select the Project Database(s) to open:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(52, 81)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(219, 17)
        Me.RadioButton1.TabIndex = 4
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Transportation Improvement Project (TIP)"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(52, 104)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(210, 17)
        Me.RadioButton2.TabIndex = 5
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Metropolitan Transportation Plan (MTP)"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(52, 127)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(143, 17)
        Me.RadioButton3.TabIndex = 6
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "Open Both TIP and MTP"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'rdoProjectScenario
        '
        Me.rdoProjectScenario.AutoSize = True
        Me.rdoProjectScenario.Location = New System.Drawing.Point(52, 150)
        Me.rdoProjectScenario.Name = "rdoProjectScenario"
        Me.rdoProjectScenario.Size = New System.Drawing.Size(141, 17)
        Me.rdoProjectScenario.TabIndex = 7
        Me.rdoProjectScenario.TabStop = True
        Me.rdoProjectScenario.Text = "Open a Project Scenario"
        Me.rdoProjectScenario.UseVisualStyleBackColor = True
        '
        'frmChkProjects
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(319, 275)
        Me.Controls.Add(Me.rdoProjectScenario)
        Me.Controls.Add(Me.RadioButton3)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnOpenProj)
        Me.Controls.Add(Me.btnSkipProjects)
        Me.Name = "frmChkProjects"
        Me.Text = "frmChkProjects"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSkipProjects As System.Windows.Forms.Button
    Friend WithEvents btnOpenProj As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents rdoProjectScenario As System.Windows.Forms.RadioButton
End Class
