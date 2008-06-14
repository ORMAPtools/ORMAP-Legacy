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
        Me.SuspendLayout()
        '
        'uxFind
        '
        Me.uxFind.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.uxFind.Location = New System.Drawing.Point(27, 123)
        Me.uxFind.Margin = New System.Windows.Forms.Padding(2)
        Me.uxFind.Name = "uxFind"
        Me.uxFind.Size = New System.Drawing.Size(75, 23)
        Me.uxFind.TabIndex = 5
        Me.uxFind.Text = "&Find"
        Me.uxFind.UseVisualStyleBackColor = True
        '
        'uxHelp
        '
        Me.uxHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.uxHelp.Location = New System.Drawing.Point(108, 123)
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
        Me.TaxlotLabel.Location = New System.Drawing.Point(6, 51)
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
        Me.uxTaxlot.Location = New System.Drawing.Point(9, 66)
        Me.uxTaxlot.Margin = New System.Windows.Forms.Padding(2)
        Me.uxTaxlot.Name = "uxTaxlot"
        Me.uxTaxlot.Size = New System.Drawing.Size(74, 20)
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
        Me.uxMapNumber.Size = New System.Drawing.Size(107, 20)
        Me.uxMapNumber.TabIndex = 1
        '
        'uxSelectFeatures
        '
        Me.uxSelectFeatures.AutoSize = True
        Me.uxSelectFeatures.Location = New System.Drawing.Point(9, 96)
        Me.uxSelectFeatures.Name = "uxSelectFeatures"
        Me.uxSelectFeatures.Size = New System.Drawing.Size(166, 17)
        Me.uxSelectFeatures.TabIndex = 4
        Me.uxSelectFeatures.Text = "Select features when locating"
        Me.uxSelectFeatures.UseVisualStyleBackColor = True
        '
        'LocateFeatureForm
        '
        Me.AcceptButton = Me.uxFind
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(192, 154)
        Me.Controls.Add(Me.uxSelectFeatures)
        Me.Controls.Add(Me.uxMapNumber)
        Me.Controls.Add(Me.uxFind)
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
End Class
