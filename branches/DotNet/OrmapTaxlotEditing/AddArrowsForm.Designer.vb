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
        Me.uxCurrentTool = New System.Windows.Forms.Label
        Me.uxNote = New System.Windows.Forms.Label
        Me.uxStandardGroup.SuspendLayout()
        Me.uxDimensionGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'uxArrowLineStyle
        '
        Me.uxArrowLineStyle.FormattingEnabled = True
        Me.uxArrowLineStyle.Location = New System.Drawing.Point(8, 25)
        Me.uxArrowLineStyle.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxArrowLineStyle.Name = "uxArrowLineStyle"
        Me.uxArrowLineStyle.Size = New System.Drawing.Size(232, 24)
        Me.uxArrowLineStyle.TabIndex = 0
        '
        'uxAddStandard
        '
        Me.uxAddStandard.Location = New System.Drawing.Point(84, 58)
        Me.uxAddStandard.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxAddStandard.Name = "uxAddStandard"
        Me.uxAddStandard.Size = New System.Drawing.Size(157, 28)
        Me.uxAddStandard.TabIndex = 1
        Me.uxAddStandard.Text = "Add Standard Arrow"
        Me.uxAddStandard.UseVisualStyleBackColor = True
        '
        'uxStandardGroup
        '
        Me.uxStandardGroup.Controls.Add(Me.uxAddStandard)
        Me.uxStandardGroup.Controls.Add(Me.uxArrowLineStyle)
        Me.uxStandardGroup.Location = New System.Drawing.Point(9, 12)
        Me.uxStandardGroup.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxStandardGroup.Name = "uxStandardGroup"
        Me.uxStandardGroup.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxStandardGroup.Size = New System.Drawing.Size(260, 103)
        Me.uxStandardGroup.TabIndex = 2
        Me.uxStandardGroup.TabStop = False
        Me.uxStandardGroup.Text = "Standard Arrows"
        '
        'uxDimensionGroup
        '
        Me.uxDimensionGroup.Controls.Add(Me.uxAddDimension)
        Me.uxDimensionGroup.Location = New System.Drawing.Point(9, 123)
        Me.uxDimensionGroup.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxDimensionGroup.Name = "uxDimensionGroup"
        Me.uxDimensionGroup.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxDimensionGroup.Size = New System.Drawing.Size(260, 63)
        Me.uxDimensionGroup.TabIndex = 3
        Me.uxDimensionGroup.TabStop = False
        Me.uxDimensionGroup.Text = "Dimension Arrows"
        '
        'uxAddDimension
        '
        Me.uxAddDimension.Location = New System.Drawing.Point(84, 23)
        Me.uxAddDimension.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxAddDimension.Name = "uxAddDimension"
        Me.uxAddDimension.Size = New System.Drawing.Size(157, 28)
        Me.uxAddDimension.TabIndex = 0
        Me.uxAddDimension.Text = "Add Dimension Arrow"
        Me.uxAddDimension.UseVisualStyleBackColor = True
        '
        'uxQuit
        '
        Me.uxQuit.Location = New System.Drawing.Point(17, 229)
        Me.uxQuit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxQuit.Name = "uxQuit"
        Me.uxQuit.Size = New System.Drawing.Size(100, 28)
        Me.uxQuit.TabIndex = 4
        Me.uxQuit.Text = "Quit"
        Me.uxQuit.UseVisualStyleBackColor = True
        '
        'uxHelp
        '
        Me.uxHelp.Location = New System.Drawing.Point(151, 229)
        Me.uxHelp.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxHelp.Name = "uxHelp"
        Me.uxHelp.Size = New System.Drawing.Size(100, 28)
        Me.uxHelp.TabIndex = 5
        Me.uxHelp.Text = "Help"
        Me.uxHelp.UseVisualStyleBackColor = True
        '
        'uxCurrentTool
        '
        Me.uxCurrentTool.AutoSize = True
        Me.uxCurrentTool.Location = New System.Drawing.Point(16, 261)
        Me.uxCurrentTool.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.uxCurrentTool.Name = "uxCurrentTool"
        Me.uxCurrentTool.Size = New System.Drawing.Size(40, 17)
        Me.uxCurrentTool.TabIndex = 6
        Me.uxCurrentTool.Text = "none"
        Me.uxCurrentTool.Visible = False
        '
        'uxNote
        '
        Me.uxNote.AutoSize = True
        Me.uxNote.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxNote.Location = New System.Drawing.Point(13, 193)
        Me.uxNote.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.uxNote.Name = "uxNote"
        Me.uxNote.Size = New System.Drawing.Size(239, 17)
        Me.uxNote.TabIndex = 7
        Me.uxNote.Text = "Press ctrl+G to exit Add Arrows."
        '
        'AddArrowsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 266)
        Me.Controls.Add(Me.uxNote)
        Me.Controls.Add(Me.uxCurrentTool)
        Me.Controls.Add(Me.uxHelp)
        Me.Controls.Add(Me.uxQuit)
        Me.Controls.Add(Me.uxDimensionGroup)
        Me.Controls.Add(Me.uxStandardGroup)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AddArrowsForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add Arrows"
        Me.TopMost = True
        Me.uxStandardGroup.ResumeLayout(False)
        Me.uxDimensionGroup.ResumeLayout(False)
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
    Friend WithEvents uxCurrentTool As System.Windows.Forms.Label
    Friend WithEvents uxNote As System.Windows.Forms.Label
End Class
