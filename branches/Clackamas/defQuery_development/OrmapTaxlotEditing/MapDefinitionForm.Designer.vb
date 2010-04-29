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
        Me.uxEnterMapNumber = New System.Windows.Forms.Label
        Me.uxDefinitonQueryTextBox = New System.Windows.Forms.TextBox
        Me.uxSetMapDefinitionQuery = New System.Windows.Forms.Button
        Me.uxCancelSetDefinitionQuery = New System.Windows.Forms.Button
        Me.uxHelpDefinitionQuery = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'uxEnterMapNumber
        '
        Me.uxEnterMapNumber.AutoSize = True
        Me.uxEnterMapNumber.Location = New System.Drawing.Point(13, 13)
        Me.uxEnterMapNumber.Name = "uxEnterMapNumber"
        Me.uxEnterMapNumber.Size = New System.Drawing.Size(102, 13)
        Me.uxEnterMapNumber.TabIndex = 0
        Me.uxEnterMapNumber.Text = "Enter Map Number: "
        '
        'uxDefinitonQueryTextBox
        '
        Me.uxDefinitonQueryTextBox.Location = New System.Drawing.Point(122, 13)
        Me.uxDefinitonQueryTextBox.Name = "uxDefinitonQueryTextBox"
        Me.uxDefinitonQueryTextBox.Size = New System.Drawing.Size(100, 20)
        Me.uxDefinitonQueryTextBox.TabIndex = 0
        '
        'uxSetMapDefinitionQuery
        '
        Me.uxSetMapDefinitionQuery.Location = New System.Drawing.Point(16, 41)
        Me.uxSetMapDefinitionQuery.Name = "uxSetMapDefinitionQuery"
        Me.uxSetMapDefinitionQuery.Size = New System.Drawing.Size(75, 23)
        Me.uxSetMapDefinitionQuery.TabIndex = 1
        Me.uxSetMapDefinitionQuery.Text = "Set Query"
        Me.uxSetMapDefinitionQuery.UseVisualStyleBackColor = True
        '
        'uxCancelSetDefinitionQuery
        '
        Me.uxCancelSetDefinitionQuery.Location = New System.Drawing.Point(97, 41)
        Me.uxCancelSetDefinitionQuery.Name = "uxCancelSetDefinitionQuery"
        Me.uxCancelSetDefinitionQuery.Size = New System.Drawing.Size(75, 23)
        Me.uxCancelSetDefinitionQuery.TabIndex = 3
        Me.uxCancelSetDefinitionQuery.Text = "Cancel"
        Me.uxCancelSetDefinitionQuery.UseVisualStyleBackColor = True
        '
        'uxHelpDefinitionQuery
        '
        Me.uxHelpDefinitionQuery.Location = New System.Drawing.Point(179, 40)
        Me.uxHelpDefinitionQuery.Name = "uxHelpDefinitionQuery"
        Me.uxHelpDefinitionQuery.Size = New System.Drawing.Size(52, 23)
        Me.uxHelpDefinitionQuery.TabIndex = 4
        Me.uxHelpDefinitionQuery.Text = "Help"
        Me.uxHelpDefinitionQuery.UseVisualStyleBackColor = True
        '
        'MapDefinitionForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(236, 80)
        Me.Controls.Add(Me.uxHelpDefinitionQuery)
        Me.Controls.Add(Me.uxCancelSetDefinitionQuery)
        Me.Controls.Add(Me.uxSetMapDefinitionQuery)
        Me.Controls.Add(Me.uxDefinitonQueryTextBox)
        Me.Controls.Add(Me.uxEnterMapNumber)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "MapDefinitionForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Definition Query"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents uxEnterMapNumber As System.Windows.Forms.Label
    Friend WithEvents uxDefinitonQueryTextBox As System.Windows.Forms.TextBox
    Friend WithEvents uxSetMapDefinitionQuery As System.Windows.Forms.Button
    Friend WithEvents uxCancelSetDefinitionQuery As System.Windows.Forms.Button
    Friend WithEvents uxHelpDefinitionQuery As System.Windows.Forms.Button
End Class
