<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaxlotAssignmentForm
    Inherits System.Windows.Forms.UserControl

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
        Me.uxIncrementBy100 = New System.Windows.Forms.RadioButton
        Me.uxIncrementBy10 = New System.Windows.Forms.RadioButton
        Me.uxIncrementBy1 = New System.Windows.Forms.RadioButton
        Me.uxTypeLabel = New System.Windows.Forms.Label
        Me.uxType = New System.Windows.Forms.ComboBox
        Me.uxIncrementByNone = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.uxIncrementBy1000 = New System.Windows.Forms.RadioButton
        Me.uxIncrementByLabel = New System.Windows.Forms.Label
        Me.uxStartingFrom = New System.Windows.Forms.TextBox
        Me.uxStartingFromLabel = New System.Windows.Forms.Label
        Me.uxHelp = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'uxIncrementBy100
        '
        Me.uxIncrementBy100.Appearance = System.Windows.Forms.Appearance.Button
        Me.uxIncrementBy100.Location = New System.Drawing.Point(174, 68)
        Me.uxIncrementBy100.Name = "uxIncrementBy100"
        Me.uxIncrementBy100.Size = New System.Drawing.Size(50, 24)
        Me.uxIncrementBy100.TabIndex = 7
        Me.uxIncrementBy100.TabStop = True
        Me.uxIncrementBy100.Text = "100"
        Me.uxIncrementBy100.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.uxIncrementBy100.UseVisualStyleBackColor = True
        '
        'uxIncrementBy10
        '
        Me.uxIncrementBy10.Appearance = System.Windows.Forms.Appearance.Button
        Me.uxIncrementBy10.Location = New System.Drawing.Point(118, 68)
        Me.uxIncrementBy10.Name = "uxIncrementBy10"
        Me.uxIncrementBy10.Size = New System.Drawing.Size(50, 24)
        Me.uxIncrementBy10.TabIndex = 6
        Me.uxIncrementBy10.TabStop = True
        Me.uxIncrementBy10.Text = "10"
        Me.uxIncrementBy10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.uxIncrementBy10.UseVisualStyleBackColor = True
        '
        'uxIncrementBy1
        '
        Me.uxIncrementBy1.Appearance = System.Windows.Forms.Appearance.Button
        Me.uxIncrementBy1.Location = New System.Drawing.Point(62, 68)
        Me.uxIncrementBy1.Name = "uxIncrementBy1"
        Me.uxIncrementBy1.Size = New System.Drawing.Size(50, 24)
        Me.uxIncrementBy1.TabIndex = 5
        Me.uxIncrementBy1.TabStop = True
        Me.uxIncrementBy1.Text = "1"
        Me.uxIncrementBy1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.uxIncrementBy1.UseVisualStyleBackColor = True
        '
        'uxTypeLabel
        '
        Me.uxTypeLabel.BackColor = System.Drawing.SystemColors.Control
        Me.uxTypeLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.uxTypeLabel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxTypeLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.uxTypeLabel.Location = New System.Drawing.Point(12, 10)
        Me.uxTypeLabel.Name = "uxTypeLabel"
        Me.uxTypeLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uxTypeLabel.Size = New System.Drawing.Size(40, 20)
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
        Me.uxType.Location = New System.Drawing.Point(58, 10)
        Me.uxType.Name = "uxType"
        Me.uxType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uxType.Size = New System.Drawing.Size(128, 22)
        Me.uxType.TabIndex = 1
        '
        'uxIncrementByNone
        '
        Me.uxIncrementByNone.Appearance = System.Windows.Forms.Appearance.Button
        Me.uxIncrementByNone.Location = New System.Drawing.Point(6, 68)
        Me.uxIncrementByNone.Name = "uxIncrementByNone"
        Me.uxIncrementByNone.Size = New System.Drawing.Size(50, 24)
        Me.uxIncrementByNone.TabIndex = 4
        Me.uxIncrementByNone.TabStop = True
        Me.uxIncrementByNone.Text = "None"
        Me.uxIncrementByNone.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.uxIncrementByNone.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.uxIncrementBy1000)
        Me.GroupBox1.Controls.Add(Me.uxIncrementByLabel)
        Me.GroupBox1.Controls.Add(Me.uxIncrementBy100)
        Me.GroupBox1.Controls.Add(Me.uxIncrementBy10)
        Me.GroupBox1.Controls.Add(Me.uxIncrementBy1)
        Me.GroupBox1.Controls.Add(Me.uxIncrementByNone)
        Me.GroupBox1.Controls.Add(Me.uxStartingFrom)
        Me.GroupBox1.Controls.Add(Me.uxStartingFromLabel)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 36)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(292, 109)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Taxlot numbering options"
        '
        'uxIncrementBy1000
        '
        Me.uxIncrementBy1000.AccessibleDescription = ""
        Me.uxIncrementBy1000.Appearance = System.Windows.Forms.Appearance.Button
        Me.uxIncrementBy1000.Location = New System.Drawing.Point(230, 68)
        Me.uxIncrementBy1000.Name = "uxIncrementBy1000"
        Me.uxIncrementBy1000.Size = New System.Drawing.Size(50, 24)
        Me.uxIncrementBy1000.TabIndex = 8
        Me.uxIncrementBy1000.TabStop = True
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
        Me.uxIncrementByLabel.Location = New System.Drawing.Point(6, 49)
        Me.uxIncrementByLabel.Name = "uxIncrementByLabel"
        Me.uxIncrementByLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uxIncrementByLabel.Size = New System.Drawing.Size(75, 17)
        Me.uxIncrementByLabel.TabIndex = 3
        Me.uxIncrementByLabel.Text = "Increment by:"
        Me.uxIncrementByLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'uxStartingFrom
        '
        Me.uxStartingFrom.AcceptsReturn = True
        Me.uxStartingFrom.BackColor = System.Drawing.Color.White
        Me.uxStartingFrom.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.uxStartingFrom.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxStartingFrom.ForeColor = System.Drawing.SystemColors.WindowText
        Me.uxStartingFrom.Location = New System.Drawing.Point(87, 23)
        Me.uxStartingFrom.MaxLength = 5
        Me.uxStartingFrom.Name = "uxStartingFrom"
        Me.uxStartingFrom.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uxStartingFrom.Size = New System.Drawing.Size(58, 20)
        Me.uxStartingFrom.TabIndex = 2
        '
        'uxStartingFromLabel
        '
        Me.uxStartingFromLabel.BackColor = System.Drawing.SystemColors.Control
        Me.uxStartingFromLabel.Cursor = System.Windows.Forms.Cursors.Default
        Me.uxStartingFromLabel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxStartingFromLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.uxStartingFromLabel.Location = New System.Drawing.Point(6, 25)
        Me.uxStartingFromLabel.Name = "uxStartingFromLabel"
        Me.uxStartingFromLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.uxStartingFromLabel.Size = New System.Drawing.Size(75, 17)
        Me.uxStartingFromLabel.TabIndex = 1
        Me.uxStartingFromLabel.Text = "Starting from:"
        Me.uxStartingFromLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'uxHelp
        '
        Me.uxHelp.Location = New System.Drawing.Point(232, 151)
        Me.uxHelp.Name = "uxHelp"
        Me.uxHelp.Size = New System.Drawing.Size(75, 23)
        Me.uxHelp.TabIndex = 3
        Me.uxHelp.Text = "Help"
        Me.uxHelp.UseVisualStyleBackColor = True
        '
        'TaxlotAssignmentForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.uxHelp)
        Me.Controls.Add(Me.uxTypeLabel)
        Me.Controls.Add(Me.uxType)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "TaxlotAssignmentForm"
        Me.Size = New System.Drawing.Size(321, 183)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents uxIncrementBy100 As System.Windows.Forms.RadioButton
    Friend WithEvents uxIncrementBy10 As System.Windows.Forms.RadioButton
    Friend WithEvents uxIncrementBy1 As System.Windows.Forms.RadioButton
    Public WithEvents uxTypeLabel As System.Windows.Forms.Label
    Public WithEvents uxType As System.Windows.Forms.ComboBox
    Friend WithEvents uxIncrementByNone As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents uxStartingFrom As System.Windows.Forms.TextBox
    Public WithEvents uxStartingFromLabel As System.Windows.Forms.Label
    Public WithEvents uxIncrementByLabel As System.Windows.Forms.Label
    Friend WithEvents uxIncrementBy1000 As System.Windows.Forms.RadioButton
    Friend WithEvents uxHelp As System.Windows.Forms.Button

End Class
