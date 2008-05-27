<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DimensionArrowsForm
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
        Me.uxCurveLabel = New System.Windows.Forms.Label
        Me.uxLineLabel = New System.Windows.Forms.Label
        Me.uxSmoothLabel = New System.Windows.Forms.Label
        Me.uxReset = New System.Windows.Forms.Button
        Me.uxApply = New System.Windows.Forms.Button
        Me.uxRatioOfCurve = New System.Windows.Forms.TextBox
        Me.uxRatioOfLine = New System.Windows.Forms.TextBox
        Me.uxSmoothRatio = New System.Windows.Forms.TextBox
        Me.uxManuallyAddArrow = New System.Windows.Forms.CheckBox
        Me.uxDimensionPropertiesGroup = New System.Windows.Forms.GroupBox
        Me.uxDimensionPropertiesGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'uxCurveLabel
        '
        Me.uxCurveLabel.AutoSize = True
        Me.uxCurveLabel.Location = New System.Drawing.Point(8, 27)
        Me.uxCurveLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.uxCurveLabel.Name = "uxCurveLabel"
        Me.uxCurveLabel.Size = New System.Drawing.Size(126, 17)
        Me.uxCurveLabel.TabIndex = 0
        Me.uxCurveLabel.Text = "Ratio of the Curve:"
        '
        'uxLineLabel
        '
        Me.uxLineLabel.AutoSize = True
        Me.uxLineLabel.Location = New System.Drawing.Point(8, 59)
        Me.uxLineLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.uxLineLabel.Name = "uxLineLabel"
        Me.uxLineLabel.Size = New System.Drawing.Size(140, 17)
        Me.uxLineLabel.TabIndex = 1
        Me.uxLineLabel.Text = "Ration from the Line:"
        '
        'uxSmoothLabel
        '
        Me.uxSmoothLabel.AutoSize = True
        Me.uxSmoothLabel.Location = New System.Drawing.Point(8, 91)
        Me.uxSmoothLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.uxSmoothLabel.Name = "uxSmoothLabel"
        Me.uxSmoothLabel.Size = New System.Drawing.Size(97, 17)
        Me.uxSmoothLabel.TabIndex = 2
        Me.uxSmoothLabel.Text = "Smooth Ratio:"
        '
        'uxReset
        '
        Me.uxReset.Location = New System.Drawing.Point(52, 172)
        Me.uxReset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxReset.Name = "uxReset"
        Me.uxReset.Size = New System.Drawing.Size(100, 28)
        Me.uxReset.TabIndex = 4
        Me.uxReset.Text = "Reset"
        Me.uxReset.UseVisualStyleBackColor = True
        '
        'uxApply
        '
        Me.uxApply.Location = New System.Drawing.Point(164, 172)
        Me.uxApply.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxApply.Name = "uxApply"
        Me.uxApply.Size = New System.Drawing.Size(100, 28)
        Me.uxApply.TabIndex = 5
        Me.uxApply.Text = "Apply"
        Me.uxApply.UseVisualStyleBackColor = True
        '
        'uxRatioOfCurve
        '
        Me.uxRatioOfCurve.Location = New System.Drawing.Point(181, 18)
        Me.uxRatioOfCurve.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxRatioOfCurve.Name = "uxRatioOfCurve"
        Me.uxRatioOfCurve.Size = New System.Drawing.Size(57, 22)
        Me.uxRatioOfCurve.TabIndex = 6
        Me.uxRatioOfCurve.Text = "1.35"
        '
        'uxRatioOfLine
        '
        Me.uxRatioOfLine.Location = New System.Drawing.Point(181, 50)
        Me.uxRatioOfLine.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxRatioOfLine.Name = "uxRatioOfLine"
        Me.uxRatioOfLine.Size = New System.Drawing.Size(57, 22)
        Me.uxRatioOfLine.TabIndex = 7
        Me.uxRatioOfLine.Text = "1.75"
        '
        'uxSmoothRatio
        '
        Me.uxSmoothRatio.Location = New System.Drawing.Point(181, 82)
        Me.uxSmoothRatio.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxSmoothRatio.Name = "uxSmoothRatio"
        Me.uxSmoothRatio.Size = New System.Drawing.Size(57, 22)
        Me.uxSmoothRatio.TabIndex = 8
        Me.uxSmoothRatio.Text = "10"
        '
        'uxManuallyAddArrow
        '
        Me.uxManuallyAddArrow.AutoSize = True
        Me.uxManuallyAddArrow.Location = New System.Drawing.Point(16, 144)
        Me.uxManuallyAddArrow.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxManuallyAddArrow.Name = "uxManuallyAddArrow"
        Me.uxManuallyAddArrow.Size = New System.Drawing.Size(148, 21)
        Me.uxManuallyAddArrow.TabIndex = 9
        Me.uxManuallyAddArrow.Text = "Manuall Add Arrow"
        Me.uxManuallyAddArrow.UseVisualStyleBackColor = True
        '
        'uxDimensionPropertiesGroup
        '
        Me.uxDimensionPropertiesGroup.Controls.Add(Me.uxRatioOfCurve)
        Me.uxDimensionPropertiesGroup.Controls.Add(Me.uxCurveLabel)
        Me.uxDimensionPropertiesGroup.Controls.Add(Me.uxSmoothRatio)
        Me.uxDimensionPropertiesGroup.Controls.Add(Me.uxLineLabel)
        Me.uxDimensionPropertiesGroup.Controls.Add(Me.uxRatioOfLine)
        Me.uxDimensionPropertiesGroup.Controls.Add(Me.uxSmoothLabel)
        Me.uxDimensionPropertiesGroup.Location = New System.Drawing.Point(16, 15)
        Me.uxDimensionPropertiesGroup.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxDimensionPropertiesGroup.Name = "uxDimensionPropertiesGroup"
        Me.uxDimensionPropertiesGroup.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.uxDimensionPropertiesGroup.Size = New System.Drawing.Size(248, 114)
        Me.uxDimensionPropertiesGroup.TabIndex = 10
        Me.uxDimensionPropertiesGroup.TabStop = False
        Me.uxDimensionPropertiesGroup.Text = "Dimension Arrow Properties"
        '
        'DimensionArrowsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(281, 214)
        Me.Controls.Add(Me.uxDimensionPropertiesGroup)
        Me.Controls.Add(Me.uxManuallyAddArrow)
        Me.Controls.Add(Me.uxApply)
        Me.Controls.Add(Me.uxReset)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DimensionArrowsForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Dimension Arrows"
        Me.TopMost = True
        Me.uxDimensionPropertiesGroup.ResumeLayout(False)
        Me.uxDimensionPropertiesGroup.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents uxCurveLabel As System.Windows.Forms.Label
    Friend WithEvents uxLineLabel As System.Windows.Forms.Label
    Friend WithEvents uxSmoothLabel As System.Windows.Forms.Label
    Friend WithEvents uxReset As System.Windows.Forms.Button
    Friend WithEvents uxApply As System.Windows.Forms.Button
    Friend WithEvents uxRatioOfCurve As System.Windows.Forms.TextBox
    Friend WithEvents uxRatioOfLine As System.Windows.Forms.TextBox
    Friend WithEvents uxSmoothRatio As System.Windows.Forms.TextBox
    Friend WithEvents uxManuallyAddArrow As System.Windows.Forms.CheckBox
    Friend WithEvents uxDimensionPropertiesGroup As System.Windows.Forms.GroupBox
End Class
