<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LocateFeatureForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.uxFind = New System.Windows.Forms.Button
        Me.uxHelp = New System.Windows.Forms.Button
        Me.MapnumberLabel = New System.Windows.Forms.Label
        Me.TaxlotLabel = New System.Windows.Forms.Label
        Me.uxTaxlot = New System.Windows.Forms.TextBox
        Me.uxMapNumber = New System.Windows.Forms.TextBox
        Me.uxSelectFeatures = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.uxCurrentlyAttNum = New System.Windows.Forms.Label
        Me.uxSetAttributeMode = New System.Windows.Forms.Button
        Me.uxAttributeMode = New System.Windows.Forms.Label
        Me.uxCurrentlyAttLbl = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'uxFind
        '
        Me.uxFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.uxFind.Location = New System.Drawing.Point(154, 43)
        Me.uxFind.Margin = New System.Windows.Forms.Padding(2)
        Me.uxFind.Name = "uxFind"
        Me.uxFind.Size = New System.Drawing.Size(75, 23)
        Me.uxFind.TabIndex = 5
        Me.uxFind.Text = "&Find"
        Me.uxFind.UseVisualStyleBackColor = True
        '
        'uxHelp
        '
        Me.uxHelp.Location = New System.Drawing.Point(169, 202)
        Me.uxHelp.Margin = New System.Windows.Forms.Padding(2)
        Me.uxHelp.Name = "uxHelp"
        Me.uxHelp.Size = New System.Drawing.Size(75, 23)
        Me.uxHelp.TabIndex = 6
        Me.uxHelp.Text = "&Help"
        Me.uxHelp.UseVisualStyleBackColor = True
        '
        'MapnumberLabel
        '
        Me.MapnumberLabel.AutoSize = True
        Me.MapnumberLabel.Location = New System.Drawing.Point(6, 9)
        Me.MapnumberLabel.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.MapnumberLabel.Name = "MapnumberLabel"
        Me.MapnumberLabel.Size = New System.Drawing.Size(71, 13)
        Me.MapnumberLabel.TabIndex = 0
        Me.MapnumberLabel.Text = "Map Number:"
        '
        'TaxlotLabel
        '
        Me.TaxlotLabel.AutoSize = True
        Me.TaxlotLabel.Location = New System.Drawing.Point(148, 9)
        Me.TaxlotLabel.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.TaxlotLabel.Name = "TaxlotLabel"
        Me.TaxlotLabel.Size = New System.Drawing.Size(39, 13)
        Me.TaxlotLabel.TabIndex = 2
        Me.TaxlotLabel.Text = "Taxlot:"
        '
        'uxTaxlot
        '
        Me.uxTaxlot.AllowDrop = True
        Me.uxTaxlot.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.uxTaxlot.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.uxTaxlot.Location = New System.Drawing.Point(151, 24)
        Me.uxTaxlot.Margin = New System.Windows.Forms.Padding(2)
        Me.uxTaxlot.Name = "uxTaxlot"
        Me.uxTaxlot.Size = New System.Drawing.Size(93, 20)
        Me.uxTaxlot.TabIndex = 3
        '
        'uxMapNumber
        '
        Me.uxMapNumber.AllowDrop = True
        Me.uxMapNumber.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.uxMapNumber.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.uxMapNumber.Location = New System.Drawing.Point(9, 24)
        Me.uxMapNumber.Margin = New System.Windows.Forms.Padding(2)
        Me.uxMapNumber.MaxLength = 12
        Me.uxMapNumber.Name = "uxMapNumber"
        Me.uxMapNumber.Size = New System.Drawing.Size(134, 20)
        Me.uxMapNumber.TabIndex = 1
        '
        'uxSelectFeatures
        '
        Me.uxSelectFeatures.AutoSize = True
        Me.uxSelectFeatures.Location = New System.Drawing.Point(5, 19)
        Me.uxSelectFeatures.Name = "uxSelectFeatures"
        Me.uxSelectFeatures.Size = New System.Drawing.Size(166, 17)
        Me.uxSelectFeatures.TabIndex = 4
        Me.uxSelectFeatures.Text = "Select features when locating"
        Me.uxSelectFeatures.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.uxCurrentlyAttNum)
        Me.GroupBox2.Controls.Add(Me.uxSetAttributeMode)
        Me.GroupBox2.Controls.Add(Me.uxAttributeMode)
        Me.GroupBox2.Controls.Add(Me.uxCurrentlyAttLbl)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 125)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(235, 72)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Editing:"
        '
        'uxCurrentlyAttNum
        '
        Me.uxCurrentlyAttNum.AutoSize = True
        Me.uxCurrentlyAttNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxCurrentlyAttNum.Location = New System.Drawing.Point(6, 52)
        Me.uxCurrentlyAttNum.Name = "uxCurrentlyAttNum"
        Me.uxCurrentlyAttNum.Size = New System.Drawing.Size(85, 13)
        Me.uxCurrentlyAttNum.TabIndex = 4
        Me.uxCurrentlyAttNum.Text = "<mapnumber>"
        Me.uxCurrentlyAttNum.Visible = False
        '
        'uxSetAttributeMode
        '
        Me.uxSetAttributeMode.Location = New System.Drawing.Point(154, 42)
        Me.uxSetAttributeMode.Name = "uxSetAttributeMode"
        Me.uxSetAttributeMode.Size = New System.Drawing.Size(75, 23)
        Me.uxSetAttributeMode.TabIndex = 2
        Me.uxSetAttributeMode.Text = "Set Manual"
        Me.uxSetAttributeMode.UseVisualStyleBackColor = True
        '
        'uxAttributeMode
        '
        Me.uxAttributeMode.AutoSize = True
        Me.uxAttributeMode.Location = New System.Drawing.Point(86, 16)
        Me.uxAttributeMode.Name = "uxAttributeMode"
        Me.uxAttributeMode.Size = New System.Drawing.Size(29, 13)
        Me.uxAttributeMode.TabIndex = 1
        Me.uxAttributeMode.Text = "Auto"
        '
        'uxCurrentlyAttLbl
        '
        Me.uxCurrentlyAttLbl.AutoSize = True
        Me.uxCurrentlyAttLbl.Location = New System.Drawing.Point(6, 37)
        Me.uxCurrentlyAttLbl.Name = "uxCurrentlyAttLbl"
        Me.uxCurrentlyAttLbl.Size = New System.Drawing.Size(101, 13)
        Me.uxCurrentlyAttLbl.TabIndex = 3
        Me.uxCurrentlyAttLbl.Text = "Currently Attributing:"
        Me.uxCurrentlyAttLbl.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Attribute Mode:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.uxSelectFeatures)
        Me.GroupBox1.Controls.Add(Me.uxFind)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 49)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(235, 72)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Locate:"
        '
        'LocateFeatureForm
        '
        Me.AcceptButton = Me.uxFind
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(253, 233)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.uxMapNumber)
        Me.Controls.Add(Me.uxHelp)
        Me.Controls.Add(Me.uxTaxlot)
        Me.Controls.Add(Me.TaxlotLabel)
        Me.Controls.Add(Me.MapnumberLabel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "LocateFeatureForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Locate Feature"
        Me.TopMost = True
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents uxFind As System.Windows.Forms.Button
    Friend WithEvents uxHelp As System.Windows.Forms.Button
    Friend WithEvents MapnumberLabel As System.Windows.Forms.Label
    Friend WithEvents TaxlotLabel As System.Windows.Forms.Label
    Friend WithEvents uxTaxlot As System.Windows.Forms.TextBox
    Friend WithEvents uxMapNumber As System.Windows.Forms.TextBox
    Friend WithEvents uxSelectFeatures As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents uxSetAttributeMode As System.Windows.Forms.Button
    Friend WithEvents uxAttributeMode As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents uxCurrentlyAttNum As System.Windows.Forms.Label
    Friend WithEvents uxCurrentlyAttLbl As System.Windows.Forms.Label
End Class
