<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SaveCogoForm
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
        Me.uxCogoSave = New System.Windows.Forms.Button
        Me.uxCogoProportion = New System.Windows.Forms.Button
        Me.uxCogoHelp = New System.Windows.Forms.Button
        Me.uxCogoQuit = New System.Windows.Forms.Button
        Me.uxPanel = New System.Windows.Forms.Panel
        Me.uxStep = New System.Windows.Forms.TextBox
        Me.uxPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'uxCogoSave
        '
        Me.uxCogoSave.Location = New System.Drawing.Point(35, 21)
        Me.uxCogoSave.Name = "uxCogoSave"
        Me.uxCogoSave.Size = New System.Drawing.Size(149, 23)
        Me.uxCogoSave.TabIndex = 0
        Me.uxCogoSave.Text = "1. Save COGO Values"
        Me.uxCogoSave.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.uxCogoSave.UseVisualStyleBackColor = True
        '
        'uxCogoProportion
        '
        Me.uxCogoProportion.Enabled = False
        Me.uxCogoProportion.Location = New System.Drawing.Point(35, 95)
        Me.uxCogoProportion.Name = "uxCogoProportion"
        Me.uxCogoProportion.Size = New System.Drawing.Size(149, 23)
        Me.uxCogoProportion.TabIndex = 1
        Me.uxCogoProportion.Text = "3. Proportion Distance(s)"
        Me.uxCogoProportion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.uxCogoProportion.UseVisualStyleBackColor = True
        '
        'uxCogoHelp
        '
        Me.uxCogoHelp.Location = New System.Drawing.Point(22, 13)
        Me.uxCogoHelp.Name = "uxCogoHelp"
        Me.uxCogoHelp.Size = New System.Drawing.Size(55, 23)
        Me.uxCogoHelp.TabIndex = 2
        Me.uxCogoHelp.Text = "Help"
        Me.uxCogoHelp.UseVisualStyleBackColor = True
        '
        'uxCogoQuit
        '
        Me.uxCogoQuit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.uxCogoQuit.Location = New System.Drawing.Point(116, 13)
        Me.uxCogoQuit.Name = "uxCogoQuit"
        Me.uxCogoQuit.Size = New System.Drawing.Size(55, 23)
        Me.uxCogoQuit.TabIndex = 3
        Me.uxCogoQuit.Text = "Quit"
        Me.uxCogoQuit.UseVisualStyleBackColor = True
        '
        'uxPanel
        '
        Me.uxPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.uxPanel.Controls.Add(Me.uxCogoHelp)
        Me.uxPanel.Controls.Add(Me.uxCogoQuit)
        Me.uxPanel.Location = New System.Drawing.Point(12, 154)
        Me.uxPanel.Name = "uxPanel"
        Me.uxPanel.Size = New System.Drawing.Size(198, 51)
        Me.uxPanel.TabIndex = 4
        '
        'uxStep
        '
        Me.uxStep.BackColor = System.Drawing.SystemColors.Control
        Me.uxStep.Location = New System.Drawing.Point(35, 60)
        Me.uxStep.Name = "uxStep"
        Me.uxStep.Size = New System.Drawing.Size(149, 20)
        Me.uxStep.TabIndex = 5
        Me.uxStep.Text = " 2. Perform Operation(s)"
        '
        'SaveCogoForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange
        Me.CancelButton = Me.uxCogoQuit
        Me.ClientSize = New System.Drawing.Size(225, 216)
        Me.Controls.Add(Me.uxStep)
        Me.Controls.Add(Me.uxCogoProportion)
        Me.Controls.Add(Me.uxCogoSave)
        Me.Controls.Add(Me.uxPanel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SaveCogoForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Save Cogo"
        Me.TopMost = True
        Me.uxPanel.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents uxCogoSave As System.Windows.Forms.Button
    Friend WithEvents uxCogoProportion As System.Windows.Forms.Button
    Friend WithEvents uxCogoHelp As System.Windows.Forms.Button
    Friend WithEvents uxCogoQuit As System.Windows.Forms.Button
    Friend WithEvents uxPanel As System.Windows.Forms.Panel
    Friend WithEvents uxStep As System.Windows.Forms.TextBox
End Class
