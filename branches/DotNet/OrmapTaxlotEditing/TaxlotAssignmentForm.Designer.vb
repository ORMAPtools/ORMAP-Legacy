<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaxlotAssignmentForm
    Inherits System.Windows.Forms.Form

    'UserControl overrides dispose to clean up the component list.
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
        Me.uxTypeLabel = New System.Windows.Forms.Label
        Me.uxType = New System.Windows.Forms.ComboBox
        Me.uxTaxlotNumberingOptions = New System.Windows.Forms.GroupBox
        Me.uxIncrementByNone = New System.Windows.Forms.RadioButton
        Me.uxIncrementBy1 = New System.Windows.Forms.RadioButton
        Me.uxIncrementBy10 = New System.Windows.Forms.RadioButton
        Me.uxIncrementBy100 = New System.Windows.Forms.RadioButton
        Me.uxIncrementBy1000 = New System.Windows.Forms.RadioButton
        Me.uxIncrementByLabel = New System.Windows.Forms.Label
        Me.uxStartingFrom = New System.Windows.Forms.TextBox
        Me.uxStartingFromLabel = New System.Windows.Forms.Label
        Me.uxHelp = New System.Windows.Forms.Button
        Me.uxTaxlotNumberingOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'uxTypeLabel
        '
        Me.uxTypeLabel.BackColor = System.Drawing.SystemColors.Control
        Me.uxTypeLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.uxTypeLabel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxTypeLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.uxTypeLabel.Location = New System.Drawing.Point(16, 12)
        Me.uxTypeLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.uxTypeLabel.Name = "uxTypeLabel"
        Me.uxTypeLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uxTypeLabel.Size = New System.Drawing.Size(53, 25)
        Me.uxTypeLabel.TabIndex = 0
        Me.uxTypeLabel.Text = "Type:"
        Me.uxTypeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'uxType
        '
        Me.uxType.BackColor = System.Drawing.SystemColors.Window
        Me.uxType.Cursor = System.Windows.Forms.Cursors.Default
        Me.uxType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.uxType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.uxType.Location = New System.Drawing.Point(77, 12)
        Me.uxType.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxType.Name = "uxType"
        Me.uxType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uxType.Size = New System.Drawing.Size(169, 24)
        Me.uxType.TabIndex = 1
        '
        'uxTaxlotNumberingOptions
        '
        Me.uxTaxlotNumberingOptions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.uxTaxlotNumberingOptions.Controls.Add(Me.uxIncrementByNone)
        Me.uxTaxlotNumberingOptions.Controls.Add(Me.uxIncrementBy1)
        Me.uxTaxlotNumberingOptions.Controls.Add(Me.uxIncrementBy10)
        Me.uxTaxlotNumberingOptions.Controls.Add(Me.uxIncrementBy100)
        Me.uxTaxlotNumberingOptions.Controls.Add(Me.uxIncrementBy1000)
        Me.uxTaxlotNumberingOptions.Controls.Add(Me.uxIncrementByLabel)
        Me.uxTaxlotNumberingOptions.Controls.Add(Me.uxStartingFrom)
        Me.uxTaxlotNumberingOptions.Controls.Add(Me.uxStartingFromLabel)
        Me.uxTaxlotNumberingOptions.Location = New System.Drawing.Point(20, 44)
        Me.uxTaxlotNumberingOptions.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxTaxlotNumberingOptions.Name = "uxTaxlotNumberingOptions"
        Me.uxTaxlotNumberingOptions.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxTaxlotNumberingOptions.Size = New System.Drawing.Size(345, 129)
        Me.uxTaxlotNumberingOptions.TabIndex = 2
        Me.uxTaxlotNumberingOptions.TabStop = False
        Me.uxTaxlotNumberingOptions.Text = "Taxlot numbering options"
        '
        'uxIncrementByNone
        '
        Me.uxIncrementByNone.Appearance = System.Windows.Forms.Appearance.Button
        Me.uxIncrementByNone.Checked = True
        Me.uxIncrementByNone.Location = New System.Drawing.Point(11, 85)
        Me.uxIncrementByNone.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxIncrementByNone.Name = "uxIncrementByNone"
        Me.uxIncrementByNone.Size = New System.Drawing.Size(59, 25)
        Me.uxIncrementByNone.TabIndex = 14
        Me.uxIncrementByNone.TabStop = True
        Me.uxIncrementByNone.Text = "None"
        Me.uxIncrementByNone.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.uxIncrementByNone.UseVisualStyleBackColor = True
        '
        'uxIncrementBy1
        '
        Me.uxIncrementBy1.Appearance = System.Windows.Forms.Appearance.Button
        Me.uxIncrementBy1.Location = New System.Drawing.Point(77, 85)
        Me.uxIncrementBy1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxIncrementBy1.Name = "uxIncrementBy1"
        Me.uxIncrementBy1.Size = New System.Drawing.Size(59, 25)
        Me.uxIncrementBy1.TabIndex = 15
        Me.uxIncrementBy1.Text = "1"
        Me.uxIncrementBy1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.uxIncrementBy1.UseVisualStyleBackColor = True
        '
        'uxIncrementBy10
        '
        Me.uxIncrementBy10.Appearance = System.Windows.Forms.Appearance.Button
        Me.uxIncrementBy10.Location = New System.Drawing.Point(144, 85)
        Me.uxIncrementBy10.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxIncrementBy10.Name = "uxIncrementBy10"
        Me.uxIncrementBy10.Size = New System.Drawing.Size(59, 25)
        Me.uxIncrementBy10.TabIndex = 16
        Me.uxIncrementBy10.Text = "10"
        Me.uxIncrementBy10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.uxIncrementBy10.UseVisualStyleBackColor = True
        '
        'uxIncrementBy100
        '
        Me.uxIncrementBy100.Appearance = System.Windows.Forms.Appearance.Button
        Me.uxIncrementBy100.Location = New System.Drawing.Point(211, 85)
        Me.uxIncrementBy100.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxIncrementBy100.Name = "uxIncrementBy100"
        Me.uxIncrementBy100.Size = New System.Drawing.Size(59, 25)
        Me.uxIncrementBy100.TabIndex = 17
        Me.uxIncrementBy100.Text = "100"
        Me.uxIncrementBy100.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.uxIncrementBy100.UseVisualStyleBackColor = True
        '
        'uxIncrementBy1000
        '
        Me.uxIncrementBy1000.AccessibleDescription = ""
        Me.uxIncrementBy1000.Appearance = System.Windows.Forms.Appearance.Button
        Me.uxIncrementBy1000.Location = New System.Drawing.Point(277, 85)
        Me.uxIncrementBy1000.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxIncrementBy1000.Name = "uxIncrementBy1000"
        Me.uxIncrementBy1000.Size = New System.Drawing.Size(59, 25)
        Me.uxIncrementBy1000.TabIndex = 18
        Me.uxIncrementBy1000.Text = "1000"
        Me.uxIncrementBy1000.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.uxIncrementBy1000.UseVisualStyleBackColor = True
        '
        'uxIncrementByLabel
        '
        Me.uxIncrementByLabel.BackColor = System.Drawing.SystemColors.Control
        Me.uxIncrementByLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.uxIncrementByLabel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxIncrementByLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.uxIncrementByLabel.Location = New System.Drawing.Point(8, 60)
        Me.uxIncrementByLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.uxIncrementByLabel.Name = "uxIncrementByLabel"
        Me.uxIncrementByLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uxIncrementByLabel.Size = New System.Drawing.Size(100, 21)
        Me.uxIncrementByLabel.TabIndex = 3
        Me.uxIncrementByLabel.Text = "Increment by:"
        Me.uxIncrementByLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'uxStartingFrom
        '
        Me.uxStartingFrom.AcceptsReturn = True
        Me.uxStartingFrom.BackColor = System.Drawing.SystemColors.Window
        Me.uxStartingFrom.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.uxStartingFrom.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxStartingFrom.ForeColor = System.Drawing.SystemColors.WindowText
        Me.uxStartingFrom.Location = New System.Drawing.Point(116, 28)
        Me.uxStartingFrom.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxStartingFrom.MaxLength = 5
        Me.uxStartingFrom.Name = "uxStartingFrom"
        Me.uxStartingFrom.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uxStartingFrom.Size = New System.Drawing.Size(76, 23)
        Me.uxStartingFrom.TabIndex = 2
        '
        'uxStartingFromLabel
        '
        Me.uxStartingFromLabel.BackColor = System.Drawing.SystemColors.Control
        Me.uxStartingFromLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.uxStartingFromLabel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxStartingFromLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.uxStartingFromLabel.Location = New System.Drawing.Point(8, 31)
        Me.uxStartingFromLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.uxStartingFromLabel.Name = "uxStartingFromLabel"
        Me.uxStartingFromLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uxStartingFromLabel.Size = New System.Drawing.Size(100, 21)
        Me.uxStartingFromLabel.TabIndex = 1
        Me.uxStartingFromLabel.Text = "Starting from:"
        Me.uxStartingFromLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'uxHelp
        '
        Me.uxHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.uxHelp.Location = New System.Drawing.Point(265, 181)
        Me.uxHelp.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxHelp.Name = "uxHelp"
        Me.uxHelp.Size = New System.Drawing.Size(100, 28)
        Me.uxHelp.TabIndex = 4
        Me.uxHelp.Text = "Help"
        Me.uxHelp.UseVisualStyleBackColor = True
        '
        'TaxlotAssignmentForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(376, 218)
        Me.Controls.Add(Me.uxHelp)
        Me.Controls.Add(Me.uxTypeLabel)
        Me.Controls.Add(Me.uxType)
        Me.Controls.Add(Me.uxTaxlotNumberingOptions)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "TaxlotAssignmentForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Taxlot Assignment"
        Me.TopMost = True
        Me.uxTaxlotNumberingOptions.ResumeLayout(False)
        Me.uxTaxlotNumberingOptions.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents uxTaxlotNumberingOptions As System.Windows.Forms.GroupBox
    Friend WithEvents uxIncrementByNone As System.Windows.Forms.RadioButton
    Friend WithEvents uxIncrementBy1 As System.Windows.Forms.RadioButton
    Friend WithEvents uxIncrementBy10 As System.Windows.Forms.RadioButton
    Friend WithEvents uxIncrementBy100 As System.Windows.Forms.RadioButton
    Friend WithEvents uxIncrementBy1000 As System.Windows.Forms.RadioButton
    Friend WithEvents uxTypeLabel As System.Windows.Forms.Label
    Friend WithEvents uxType As System.Windows.Forms.ComboBox
    Friend WithEvents uxStartingFrom As System.Windows.Forms.TextBox
    Friend WithEvents uxStartingFromLabel As System.Windows.Forms.Label
    Friend WithEvents uxIncrementByLabel As System.Windows.Forms.Label
    Friend WithEvents uxHelp As System.Windows.Forms.Button

End Class
