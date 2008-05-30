<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddArrowsForm
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
        Me.uxArrowLineStyle = New System.Windows.Forms.ComboBox
        Me.uxAddStandard = New System.Windows.Forms.Button
        Me.uxStandardGroup = New System.Windows.Forms.GroupBox
        Me.uxDimensionGroup = New System.Windows.Forms.GroupBox
        Me.uxAddDimension = New System.Windows.Forms.Button
        Me.uxQuit = New System.Windows.Forms.Button
        Me.uxHelp = New System.Windows.Forms.Button
        Me.uxNote = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.uxStandardGroup.SuspendLayout()
        Me.uxDimensionGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'uxArrowLineStyle
        '
        Me.uxArrowLineStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.uxArrowLineStyle.FormattingEnabled = True
        Me.uxArrowLineStyle.Location = New System.Drawing.Point(6, 20)
        Me.uxArrowLineStyle.Name = "uxArrowLineStyle"
        Me.uxArrowLineStyle.Size = New System.Drawing.Size(175, 21)
        Me.uxArrowLineStyle.TabIndex = 0
        '
        'uxAddStandard
        '
        Me.uxAddStandard.Location = New System.Drawing.Point(49, 47)
        Me.uxAddStandard.Name = "uxAddStandard"
        Me.uxAddStandard.Size = New System.Drawing.Size(132, 23)
        Me.uxAddStandard.TabIndex = 1
        Me.uxAddStandard.Text = "Add Standard Arrow"
        Me.uxAddStandard.UseVisualStyleBackColor = True
        '
        'uxStandardGroup
        '
        Me.uxStandardGroup.Controls.Add(Me.uxAddStandard)
        Me.uxStandardGroup.Controls.Add(Me.uxArrowLineStyle)
        Me.uxStandardGroup.Location = New System.Drawing.Point(7, 10)
        Me.uxStandardGroup.Name = "uxStandardGroup"
        Me.uxStandardGroup.Size = New System.Drawing.Size(195, 84)
        Me.uxStandardGroup.TabIndex = 2
        Me.uxStandardGroup.TabStop = False
        Me.uxStandardGroup.Text = "Standard Arrows"
        '
        'uxDimensionGroup
        '
        Me.uxDimensionGroup.Controls.Add(Me.TextBox1)
        Me.uxDimensionGroup.Controls.Add(Me.uxAddDimension)
        Me.uxDimensionGroup.Location = New System.Drawing.Point(7, 100)
        Me.uxDimensionGroup.Name = "uxDimensionGroup"
        Me.uxDimensionGroup.Size = New System.Drawing.Size(194, 112)
        Me.uxDimensionGroup.TabIndex = 3
        Me.uxDimensionGroup.TabStop = False
        Me.uxDimensionGroup.Text = "Dimension Arrows"
        '
        'uxAddDimension
        '
        Me.uxAddDimension.Location = New System.Drawing.Point(49, 19)
        Me.uxAddDimension.Name = "uxAddDimension"
        Me.uxAddDimension.Size = New System.Drawing.Size(132, 23)
        Me.uxAddDimension.TabIndex = 0
        Me.uxAddDimension.Text = "Add Dimension Arrow"
        Me.uxAddDimension.UseVisualStyleBackColor = True
        '
        'uxQuit
        '
        Me.uxQuit.Location = New System.Drawing.Point(45, 240)
        Me.uxQuit.Name = "uxQuit"
        Me.uxQuit.Size = New System.Drawing.Size(75, 23)
        Me.uxQuit.TabIndex = 4
        Me.uxQuit.Text = "Quit"
        Me.uxQuit.UseVisualStyleBackColor = True
        '
        'uxHelp
        '
        Me.uxHelp.Location = New System.Drawing.Point(126, 240)
        Me.uxHelp.Name = "uxHelp"
        Me.uxHelp.Size = New System.Drawing.Size(75, 23)
        Me.uxHelp.TabIndex = 5
        Me.uxHelp.Text = "Help"
        Me.uxHelp.UseVisualStyleBackColor = True
        '
        'uxNote
        '
        Me.uxNote.AutoSize = True
        Me.uxNote.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxNote.Location = New System.Drawing.Point(20, 218)
        Me.uxNote.Name = "uxNote"
        Me.uxNote.Size = New System.Drawing.Size(172, 13)
        Me.uxNote.TabIndex = 7
        Me.uxNote.Text = "Press ""q"" to exit Add Arrows."
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Location = New System.Drawing.Point(6, 48)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(175, 58)
        Me.TextBox1.TabIndex = 8
        Me.TextBox1.Text = "Note: Dimension Properties can be altered by pressing ""d"" when in ""Add Dimension " & _
            "Arrow"" mode.  See the Help section for more information."
        '
        'AddArrowsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(212, 275)
        Me.Controls.Add(Me.uxNote)
        Me.Controls.Add(Me.uxHelp)
        Me.Controls.Add(Me.uxQuit)
        Me.Controls.Add(Me.uxDimensionGroup)
        Me.Controls.Add(Me.uxStandardGroup)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AddArrowsForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add Arrows"
        Me.TopMost = True
        Me.uxStandardGroup.ResumeLayout(False)
        Me.uxDimensionGroup.ResumeLayout(False)
        Me.uxDimensionGroup.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents uxArrowLineStyle As System.Windows.Forms.ComboBox
    Friend WithEvents uxAddStandard As System.Windows.Forms.Button
    Friend WithEvents uxStandardGroup As System.Windows.Forms.GroupBox
    Friend WithEvents uxDimensionGroup As System.Windows.Forms.GroupBox
    Friend WithEvents uxAddDimension As System.Windows.Forms.Button
    Friend WithEvents uxQuit As System.Windows.Forms.Button
    Friend WithEvents uxHelp As System.Windows.Forms.Button
    Friend WithEvents uxNote As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
