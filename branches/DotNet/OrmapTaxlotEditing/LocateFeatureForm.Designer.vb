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
        Me.uxMapnumber = New System.Windows.Forms.ComboBox
        Me.uxFind = New System.Windows.Forms.Button
        Me.uxHelp = New System.Windows.Forms.Button
        Me.MapnumberLabel = New System.Windows.Forms.Label
        Me.TaxlotLabel = New System.Windows.Forms.Label
        Me.uxTaxlot = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'uxMapnumber
        '
        Me.uxMapnumber.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.uxMapnumber.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.uxMapnumber.FormattingEnabled = True
        Me.uxMapnumber.Location = New System.Drawing.Point(12, 34)
        Me.uxMapnumber.Name = "uxMapnumber"
        Me.uxMapnumber.Size = New System.Drawing.Size(202, 24)
        Me.uxMapnumber.Sorted = True
        Me.uxMapnumber.TabIndex = 0
        '
        'uxFind
        '
        Me.uxFind.Location = New System.Drawing.Point(58, 125)
        Me.uxFind.Name = "uxFind"
        Me.uxFind.Size = New System.Drawing.Size(75, 27)
        Me.uxFind.TabIndex = 2
        Me.uxFind.Text = "Find"
        Me.uxFind.UseVisualStyleBackColor = True
        '
        'uxHelp
        '
        Me.uxHelp.Location = New System.Drawing.Point(139, 125)
        Me.uxHelp.Name = "uxHelp"
        Me.uxHelp.Size = New System.Drawing.Size(75, 27)
        Me.uxHelp.TabIndex = 3
        Me.uxHelp.Text = "Help"
        Me.uxHelp.UseVisualStyleBackColor = True
        '
        'MapnumberLabel
        '
        Me.MapnumberLabel.AutoSize = True
        Me.MapnumberLabel.Location = New System.Drawing.Point(9, 14)
        Me.MapnumberLabel.Name = "MapnumberLabel"
        Me.MapnumberLabel.Size = New System.Drawing.Size(93, 17)
        Me.MapnumberLabel.TabIndex = 4
        Me.MapnumberLabel.Text = "Map Number:"
        '
        'TaxlotLabel
        '
        Me.TaxlotLabel.AutoSize = True
        Me.TaxlotLabel.Location = New System.Drawing.Point(12, 63)
        Me.TaxlotLabel.Name = "TaxlotLabel"
        Me.TaxlotLabel.Size = New System.Drawing.Size(50, 17)
        Me.TaxlotLabel.TabIndex = 5
        Me.TaxlotLabel.Text = "Taxlot:"
        '
        'uxTaxlot
        '
        Me.uxTaxlot.Location = New System.Drawing.Point(12, 83)
        Me.uxTaxlot.Name = "uxTaxlot"
        Me.uxTaxlot.Size = New System.Drawing.Size(100, 22)
        Me.uxTaxlot.TabIndex = 6
        '
        'LocateFeatureForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(235, 163)
        Me.Controls.Add(Me.uxFind)
        Me.Controls.Add(Me.uxHelp)
        Me.Controls.Add(Me.uxMapnumber)
        Me.Controls.Add(Me.uxTaxlot)
        Me.Controls.Add(Me.TaxlotLabel)
        Me.Controls.Add(Me.MapnumberLabel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
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
    Friend WithEvents uxMapnumber As System.Windows.Forms.ComboBox
    Friend WithEvents uxFind As System.Windows.Forms.Button
    Friend WithEvents uxHelp As System.Windows.Forms.Button
    Friend WithEvents MapnumberLabel As System.Windows.Forms.Label
    Friend WithEvents TaxlotLabel As System.Windows.Forms.Label
    Friend WithEvents uxTaxlot As System.Windows.Forms.TextBox
End Class
