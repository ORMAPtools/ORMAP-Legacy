<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MapDefinitionForm
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
        Me.EnterMapNumberLabel = New System.Windows.Forms.Label
        Me.uxMapNumberTextBox = New System.Windows.Forms.TextBox
        Me.uxSetMapDefinitionQuery = New System.Windows.Forms.Button
        Me.uxCancelSetDefinitionQuery = New System.Windows.Forms.Button
        Me.uxHelpDefinitionQuery = New System.Windows.Forms.Button
        Me.uxMapNumberOption = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.uxMapScaleOption = New System.Windows.Forms.ComboBox
        Me.uxMapScale = New System.Windows.Forms.TextBox
        Me.WarningLabel = New System.Windows.Forms.Label
        Me.SpecifyMapNumberLabel = New System.Windows.Forms.Label
        Me.FeaturesUsingMapScaleLabel = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'EnterMapNumberLabel
        '
        Me.EnterMapNumberLabel.AutoSize = True
        Me.EnterMapNumberLabel.Location = New System.Drawing.Point(13, 25)
        Me.EnterMapNumberLabel.Name = "EnterMapNumberLabel"
        Me.EnterMapNumberLabel.Size = New System.Drawing.Size(188, 13)
        Me.EnterMapNumberLabel.TabIndex = 0
        Me.EnterMapNumberLabel.Text = "Display features where Map Number is"
        '
        'uxMapNumberTextBox
        '
        Me.uxMapNumberTextBox.AllowDrop = True
        Me.uxMapNumberTextBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.uxMapNumberTextBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.uxMapNumberTextBox.Location = New System.Drawing.Point(273, 22)
        Me.uxMapNumberTextBox.Name = "uxMapNumberTextBox"
        Me.uxMapNumberTextBox.Size = New System.Drawing.Size(110, 20)
        Me.uxMapNumberTextBox.TabIndex = 0
        '
        'uxSetMapDefinitionQuery
        '
        Me.uxSetMapDefinitionQuery.Location = New System.Drawing.Point(167, 85)
        Me.uxSetMapDefinitionQuery.Name = "uxSetMapDefinitionQuery"
        Me.uxSetMapDefinitionQuery.Size = New System.Drawing.Size(75, 23)
        Me.uxSetMapDefinitionQuery.TabIndex = 1
        Me.uxSetMapDefinitionQuery.Text = "Set Query"
        Me.uxSetMapDefinitionQuery.UseVisualStyleBackColor = True
        '
        'uxCancelSetDefinitionQuery
        '
        Me.uxCancelSetDefinitionQuery.Location = New System.Drawing.Point(248, 85)
        Me.uxCancelSetDefinitionQuery.Name = "uxCancelSetDefinitionQuery"
        Me.uxCancelSetDefinitionQuery.Size = New System.Drawing.Size(75, 23)
        Me.uxCancelSetDefinitionQuery.TabIndex = 3
        Me.uxCancelSetDefinitionQuery.Text = "Cancel"
        Me.uxCancelSetDefinitionQuery.UseVisualStyleBackColor = True
        '
        'uxHelpDefinitionQuery
        '
        Me.uxHelpDefinitionQuery.Location = New System.Drawing.Point(330, 84)
        Me.uxHelpDefinitionQuery.Name = "uxHelpDefinitionQuery"
        Me.uxHelpDefinitionQuery.Size = New System.Drawing.Size(52, 23)
        Me.uxHelpDefinitionQuery.TabIndex = 4
        Me.uxHelpDefinitionQuery.Text = "Help"
        Me.uxHelpDefinitionQuery.UseVisualStyleBackColor = True
        '
        'uxMapNumberOption
        '
        Me.uxMapNumberOption.AllowDrop = True
        Me.uxMapNumberOption.FormattingEnabled = True
        Me.uxMapNumberOption.Location = New System.Drawing.Point(207, 21)
        Me.uxMapNumberOption.Name = "uxMapNumberOption"
        Me.uxMapNumberOption.Size = New System.Drawing.Size(44, 21)
        Me.uxMapNumberOption.TabIndex = 5
        Me.uxMapNumberOption.Text = "  =  "
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(175, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Display features where map scale is"
        '
        'uxMapScaleOption
        '
        Me.uxMapScaleOption.FormattingEnabled = True
        Me.uxMapScaleOption.Location = New System.Drawing.Point(207, 58)
        Me.uxMapScaleOption.Name = "uxMapScaleOption"
        Me.uxMapScaleOption.Size = New System.Drawing.Size(44, 21)
        Me.uxMapScaleOption.TabIndex = 7
        Me.uxMapScaleOption.Text = "  =  "
        '
        'uxMapScale
        '
        Me.uxMapScale.Location = New System.Drawing.Point(274, 58)
        Me.uxMapScale.Name = "uxMapScale"
        Me.uxMapScale.Size = New System.Drawing.Size(109, 20)
        Me.uxMapScale.TabIndex = 8
        '
        'WarningLabel
        '
        Me.WarningLabel.AutoSize = True
        Me.WarningLabel.ForeColor = System.Drawing.Color.Red
        Me.WarningLabel.Location = New System.Drawing.Point(16, 115)
        Me.WarningLabel.Name = "WarningLabel"
        Me.WarningLabel.Size = New System.Drawing.Size(307, 26)
        Me.WarningLabel.TabIndex = 9
        Me.WarningLabel.Text = "Warning: This tool will clear out any exisiting definition queries in" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "          " & _
            "       participaiting layers."
        '
        'SpecifyMapNumberLabel
        '
        Me.SpecifyMapNumberLabel.AutoSize = True
        Me.SpecifyMapNumberLabel.ForeColor = System.Drawing.Color.MediumBlue
        Me.SpecifyMapNumberLabel.Location = New System.Drawing.Point(15, 9)
        Me.SpecifyMapNumberLabel.Name = "SpecifyMapNumberLabel"
        Me.SpecifyMapNumberLabel.Size = New System.Drawing.Size(137, 13)
        Me.SpecifyMapNumberLabel.TabIndex = 10
        Me.SpecifyMapNumberLabel.Text = "Features using map number"
        '
        'FeaturesUsingMapScaleLabel
        '
        Me.FeaturesUsingMapScaleLabel.AutoSize = True
        Me.FeaturesUsingMapScaleLabel.ForeColor = System.Drawing.Color.MediumBlue
        Me.FeaturesUsingMapScaleLabel.Location = New System.Drawing.Point(12, 49)
        Me.FeaturesUsingMapScaleLabel.Name = "FeaturesUsingMapScaleLabel"
        Me.FeaturesUsingMapScaleLabel.Size = New System.Drawing.Size(127, 13)
        Me.FeaturesUsingMapScaleLabel.TabIndex = 11
        Me.FeaturesUsingMapScaleLabel.Text = "Features using map scale"
        '
        'MapDefinitionForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(398, 165)
        Me.Controls.Add(Me.FeaturesUsingMapScaleLabel)
        Me.Controls.Add(Me.SpecifyMapNumberLabel)
        Me.Controls.Add(Me.WarningLabel)
        Me.Controls.Add(Me.uxMapScale)
        Me.Controls.Add(Me.uxMapScaleOption)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.uxMapNumberOption)
        Me.Controls.Add(Me.uxHelpDefinitionQuery)
        Me.Controls.Add(Me.uxCancelSetDefinitionQuery)
        Me.Controls.Add(Me.uxSetMapDefinitionQuery)
        Me.Controls.Add(Me.uxMapNumberTextBox)
        Me.Controls.Add(Me.EnterMapNumberLabel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "MapDefinitionForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Definition Query"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents EnterMapNumberLabel As System.Windows.Forms.Label
    Friend WithEvents uxMapNumberTextBox As System.Windows.Forms.TextBox
    Friend WithEvents uxSetMapDefinitionQuery As System.Windows.Forms.Button
    Friend WithEvents uxCancelSetDefinitionQuery As System.Windows.Forms.Button
    Friend WithEvents uxHelpDefinitionQuery As System.Windows.Forms.Button
    Friend WithEvents uxMapNumberOption As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents uxMapScaleOption As System.Windows.Forms.ComboBox
    Friend WithEvents uxMapScale As System.Windows.Forms.TextBox
    Friend WithEvents WarningLabel As System.Windows.Forms.Label
    Friend WithEvents SpecifyMapNumberLabel As System.Windows.Forms.Label
    Friend WithEvents FeaturesUsingMapScaleLabel As System.Windows.Forms.Label
End Class
