<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class HelpForm
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
        Me.uxOK = New System.Windows.Forms.Button
        Me.uxContent = New System.Windows.Forms.RichTextBox
        Me.SuspendLayout()
        '
        'uxOK
        '
        Me.uxOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.uxOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.uxOK.Location = New System.Drawing.Point(509, 335)
        Me.uxOK.Name = "uxOK"
        Me.uxOK.Size = New System.Drawing.Size(75, 23)
        Me.uxOK.TabIndex = 2
        Me.uxOK.Text = "&OK"
        '
        'uxContent
        '
        Me.uxContent.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.uxContent.BackColor = System.Drawing.SystemColors.Window
        Me.uxContent.Location = New System.Drawing.Point(8, 8)
        Me.uxContent.Name = "uxContent"
        Me.uxContent.Size = New System.Drawing.Size(576, 321)
        Me.uxContent.TabIndex = 3
        Me.uxContent.Text = "(No help file loaded.)"
        '
        'HelpForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(592, 366)
        Me.Controls.Add(Me.uxOK)
        Me.Controls.Add(Me.uxContent)
        Me.Name = "HelpForm"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.ShowIcon = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.Text = "ORMAP Taxlot Editing Help - No Topic"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents uxOK As System.Windows.Forms.Button
    Friend WithEvents uxContent As System.Windows.Forms.RichTextBox
End Class
