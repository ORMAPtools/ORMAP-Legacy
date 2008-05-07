<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EditMapIndexForm
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
        Me.uxTownship = New System.Windows.Forms.ComboBox
        Me.uxTownshipDirectional = New System.Windows.Forms.ComboBox
        Me.uxTownshipPartial = New System.Windows.Forms.ComboBox
        Me.uxTownshipLabel = New System.Windows.Forms.Label
        Me.uxTownshipDirectionalLabel = New System.Windows.Forms.Label
        Me.uxTownshipPartialLabel = New System.Windows.Forms.Label
        Me.uxTownshipGroupBox = New System.Windows.Forms.GroupBox
        Me.uxRangeGroupBox = New System.Windows.Forms.GroupBox
        Me.uxRangePartialLabel = New System.Windows.Forms.Label
        Me.uxRangeDirectionalLabel = New System.Windows.Forms.Label
        Me.uxRangeLabel = New System.Windows.Forms.Label
        Me.uxRangePartial = New System.Windows.Forms.ComboBox
        Me.uxRangeDirectional = New System.Windows.Forms.ComboBox
        Me.uxRange = New System.Windows.Forms.ComboBox
        Me.uxSectionGroupBox = New System.Windows.Forms.GroupBox
        Me.uxSectionQtrQtrLabel = New System.Windows.Forms.Label
        Me.uxSectionQtrLabel = New System.Windows.Forms.Label
        Me.uxSectionLabel = New System.Windows.Forms.Label
        Me.uxSectionQtrQtr = New System.Windows.Forms.ComboBox
        Me.uxSectionQtr = New System.Windows.Forms.ComboBox
        Me.uxSection = New System.Windows.Forms.ComboBox
        Me.uxMapInfoGroupBox = New System.Windows.Forms.GroupBox
        Me.uxAnomalyLabel = New System.Windows.Forms.Label
        Me.uxPageLabel = New System.Windows.Forms.Label
        Me.uxScaleLabel = New System.Windows.Forms.Label
        Me.uxReliabilityLabel = New System.Windows.Forms.Label
        Me.uxSuffixTypeLabel = New System.Windows.Forms.Label
        Me.uxSuffixNumberLabel = New System.Windows.Forms.Label
        Me.uxMapNumberLabel = New System.Windows.Forms.Label
        Me.uxAnomaly = New System.Windows.Forms.TextBox
        Me.uxPage = New System.Windows.Forms.TextBox
        Me.uxScale = New System.Windows.Forms.ComboBox
        Me.uxReliability = New System.Windows.Forms.ComboBox
        Me.uxCountyLabel = New System.Windows.Forms.Label
        Me.uxSuffixType = New System.Windows.Forms.ComboBox
        Me.uxSuffixNumber = New System.Windows.Forms.TextBox
        Me.uxMapNumber = New System.Windows.Forms.TextBox
        Me.uxCounty = New System.Windows.Forms.ComboBox
        Me.uxORMAPNumberGroupBox = New System.Windows.Forms.GroupBox
        Me.uxORMAPNumberLabel = New System.Windows.Forms.Label
        Me.uxHelp = New System.Windows.Forms.Button
        Me.uxEdit = New System.Windows.Forms.Button
        Me.uxQuit = New System.Windows.Forms.Button
        Me.uxTownshipGroupBox.SuspendLayout()
        Me.uxRangeGroupBox.SuspendLayout()
        Me.uxSectionGroupBox.SuspendLayout()
        Me.uxMapInfoGroupBox.SuspendLayout()
        Me.uxORMAPNumberGroupBox.SuspendLayout()
        Me.SuspendLayout()
        '
        'uxTownship
        '
        Me.uxTownship.FormattingEnabled = True
        Me.uxTownship.Location = New System.Drawing.Point(101, 14)
        Me.uxTownship.Name = "uxTownship"
        Me.uxTownship.Size = New System.Drawing.Size(70, 21)
        Me.uxTownship.TabIndex = 0
        '
        'uxTownshipDirectional
        '
        Me.uxTownshipDirectional.FormattingEnabled = True
        Me.uxTownshipDirectional.Location = New System.Drawing.Point(101, 39)
        Me.uxTownshipDirectional.Name = "uxTownshipDirectional"
        Me.uxTownshipDirectional.Size = New System.Drawing.Size(70, 21)
        Me.uxTownshipDirectional.TabIndex = 1
        '
        'uxTownshipPartial
        '
        Me.uxTownshipPartial.FormattingEnabled = True
        Me.uxTownshipPartial.Location = New System.Drawing.Point(101, 64)
        Me.uxTownshipPartial.Name = "uxTownshipPartial"
        Me.uxTownshipPartial.Size = New System.Drawing.Size(70, 21)
        Me.uxTownshipPartial.TabIndex = 2
        '
        'uxTownshipLabel
        '
        Me.uxTownshipLabel.AutoSize = True
        Me.uxTownshipLabel.Location = New System.Drawing.Point(6, 17)
        Me.uxTownshipLabel.Name = "uxTownshipLabel"
        Me.uxTownshipLabel.Size = New System.Drawing.Size(47, 13)
        Me.uxTownshipLabel.TabIndex = 3
        Me.uxTownshipLabel.Text = "Number:"
        '
        'uxTownshipDirectionalLabel
        '
        Me.uxTownshipDirectionalLabel.AutoSize = True
        Me.uxTownshipDirectionalLabel.Location = New System.Drawing.Point(6, 42)
        Me.uxTownshipDirectionalLabel.Name = "uxTownshipDirectionalLabel"
        Me.uxTownshipDirectionalLabel.Size = New System.Drawing.Size(60, 13)
        Me.uxTownshipDirectionalLabel.TabIndex = 4
        Me.uxTownshipDirectionalLabel.Text = "Directional:"
        '
        'uxTownshipPartialLabel
        '
        Me.uxTownshipPartialLabel.AutoSize = True
        Me.uxTownshipPartialLabel.Location = New System.Drawing.Point(6, 67)
        Me.uxTownshipPartialLabel.Name = "uxTownshipPartialLabel"
        Me.uxTownshipPartialLabel.Size = New System.Drawing.Size(67, 13)
        Me.uxTownshipPartialLabel.TabIndex = 5
        Me.uxTownshipPartialLabel.Text = "Partial Code:"
        '
        'uxTownshipGroupBox
        '
        Me.uxTownshipGroupBox.Controls.Add(Me.uxTownshipPartialLabel)
        Me.uxTownshipGroupBox.Controls.Add(Me.uxTownshipPartial)
        Me.uxTownshipGroupBox.Controls.Add(Me.uxTownshipDirectional)
        Me.uxTownshipGroupBox.Controls.Add(Me.uxTownshipDirectionalLabel)
        Me.uxTownshipGroupBox.Controls.Add(Me.uxTownship)
        Me.uxTownshipGroupBox.Controls.Add(Me.uxTownshipLabel)
        Me.uxTownshipGroupBox.Location = New System.Drawing.Point(10, 6)
        Me.uxTownshipGroupBox.Name = "uxTownshipGroupBox"
        Me.uxTownshipGroupBox.Size = New System.Drawing.Size(181, 92)
        Me.uxTownshipGroupBox.TabIndex = 5
        Me.uxTownshipGroupBox.TabStop = False
        Me.uxTownshipGroupBox.Text = "Township"
        '
        'uxRangeGroupBox
        '
        Me.uxRangeGroupBox.Controls.Add(Me.uxRangePartialLabel)
        Me.uxRangeGroupBox.Controls.Add(Me.uxRangeDirectionalLabel)
        Me.uxRangeGroupBox.Controls.Add(Me.uxRangeLabel)
        Me.uxRangeGroupBox.Controls.Add(Me.uxRangePartial)
        Me.uxRangeGroupBox.Controls.Add(Me.uxRangeDirectional)
        Me.uxRangeGroupBox.Controls.Add(Me.uxRange)
        Me.uxRangeGroupBox.Location = New System.Drawing.Point(10, 104)
        Me.uxRangeGroupBox.Name = "uxRangeGroupBox"
        Me.uxRangeGroupBox.Size = New System.Drawing.Size(181, 92)
        Me.uxRangeGroupBox.TabIndex = 6
        Me.uxRangeGroupBox.TabStop = False
        Me.uxRangeGroupBox.Text = "Range"
        '
        'uxRangePartialLabel
        '
        Me.uxRangePartialLabel.AutoSize = True
        Me.uxRangePartialLabel.Location = New System.Drawing.Point(6, 67)
        Me.uxRangePartialLabel.Name = "uxRangePartialLabel"
        Me.uxRangePartialLabel.Size = New System.Drawing.Size(67, 13)
        Me.uxRangePartialLabel.TabIndex = 5
        Me.uxRangePartialLabel.Text = "Partial Code:"
        '
        'uxRangeDirectionalLabel
        '
        Me.uxRangeDirectionalLabel.AutoSize = True
        Me.uxRangeDirectionalLabel.Location = New System.Drawing.Point(6, 42)
        Me.uxRangeDirectionalLabel.Name = "uxRangeDirectionalLabel"
        Me.uxRangeDirectionalLabel.Size = New System.Drawing.Size(60, 13)
        Me.uxRangeDirectionalLabel.TabIndex = 4
        Me.uxRangeDirectionalLabel.Text = "Directional:"
        '
        'uxRangeLabel
        '
        Me.uxRangeLabel.AutoSize = True
        Me.uxRangeLabel.Location = New System.Drawing.Point(6, 17)
        Me.uxRangeLabel.Name = "uxRangeLabel"
        Me.uxRangeLabel.Size = New System.Drawing.Size(47, 13)
        Me.uxRangeLabel.TabIndex = 3
        Me.uxRangeLabel.Text = "Number:"
        '
        'uxRangePartial
        '
        Me.uxRangePartial.FormattingEnabled = True
        Me.uxRangePartial.Location = New System.Drawing.Point(101, 64)
        Me.uxRangePartial.Name = "uxRangePartial"
        Me.uxRangePartial.Size = New System.Drawing.Size(70, 21)
        Me.uxRangePartial.TabIndex = 2
        '
        'uxRangeDirectional
        '
        Me.uxRangeDirectional.FormattingEnabled = True
        Me.uxRangeDirectional.Location = New System.Drawing.Point(101, 39)
        Me.uxRangeDirectional.Name = "uxRangeDirectional"
        Me.uxRangeDirectional.Size = New System.Drawing.Size(70, 21)
        Me.uxRangeDirectional.TabIndex = 1
        '
        'uxRange
        '
        Me.uxRange.FormattingEnabled = True
        Me.uxRange.Location = New System.Drawing.Point(101, 14)
        Me.uxRange.Name = "uxRange"
        Me.uxRange.Size = New System.Drawing.Size(70, 21)
        Me.uxRange.TabIndex = 0
        '
        'uxSectionGroupBox
        '
        Me.uxSectionGroupBox.Controls.Add(Me.uxSectionQtrQtrLabel)
        Me.uxSectionGroupBox.Controls.Add(Me.uxSectionQtrLabel)
        Me.uxSectionGroupBox.Controls.Add(Me.uxSectionLabel)
        Me.uxSectionGroupBox.Controls.Add(Me.uxSectionQtrQtr)
        Me.uxSectionGroupBox.Controls.Add(Me.uxSectionQtr)
        Me.uxSectionGroupBox.Controls.Add(Me.uxSection)
        Me.uxSectionGroupBox.Location = New System.Drawing.Point(10, 203)
        Me.uxSectionGroupBox.Name = "uxSectionGroupBox"
        Me.uxSectionGroupBox.Size = New System.Drawing.Size(181, 92)
        Me.uxSectionGroupBox.TabIndex = 7
        Me.uxSectionGroupBox.TabStop = False
        Me.uxSectionGroupBox.Text = "Section"
        '
        'uxSectionQtrQtrLabel
        '
        Me.uxSectionQtrQtrLabel.AutoSize = True
        Me.uxSectionQtrQtrLabel.Location = New System.Drawing.Point(2, 67)
        Me.uxSectionQtrQtrLabel.Name = "uxSectionQtrQtrLabel"
        Me.uxSectionQtrQtrLabel.Size = New System.Drawing.Size(95, 13)
        Me.uxSectionQtrQtrLabel.TabIndex = 5
        Me.uxSectionQtrQtrLabel.Text = "Quarter of Quarter:"
        '
        'uxSectionQtrLabel
        '
        Me.uxSectionQtrLabel.AutoSize = True
        Me.uxSectionQtrLabel.Location = New System.Drawing.Point(6, 42)
        Me.uxSectionQtrLabel.Name = "uxSectionQtrLabel"
        Me.uxSectionQtrLabel.Size = New System.Drawing.Size(45, 13)
        Me.uxSectionQtrLabel.TabIndex = 4
        Me.uxSectionQtrLabel.Text = "Quarter:"
        '
        'uxSectionLabel
        '
        Me.uxSectionLabel.AutoSize = True
        Me.uxSectionLabel.Location = New System.Drawing.Point(6, 17)
        Me.uxSectionLabel.Name = "uxSectionLabel"
        Me.uxSectionLabel.Size = New System.Drawing.Size(47, 13)
        Me.uxSectionLabel.TabIndex = 3
        Me.uxSectionLabel.Text = "Number:"
        '
        'uxSectionQtrQtr
        '
        Me.uxSectionQtrQtr.FormattingEnabled = True
        Me.uxSectionQtrQtr.Location = New System.Drawing.Point(101, 64)
        Me.uxSectionQtrQtr.Name = "uxSectionQtrQtr"
        Me.uxSectionQtrQtr.Size = New System.Drawing.Size(70, 21)
        Me.uxSectionQtrQtr.TabIndex = 2
        '
        'uxSectionQtr
        '
        Me.uxSectionQtr.FormattingEnabled = True
        Me.uxSectionQtr.Location = New System.Drawing.Point(101, 39)
        Me.uxSectionQtr.Name = "uxSectionQtr"
        Me.uxSectionQtr.Size = New System.Drawing.Size(70, 21)
        Me.uxSectionQtr.TabIndex = 1
        '
        'uxSection
        '
        Me.uxSection.FormattingEnabled = True
        Me.uxSection.Location = New System.Drawing.Point(101, 14)
        Me.uxSection.Name = "uxSection"
        Me.uxSection.Size = New System.Drawing.Size(70, 21)
        Me.uxSection.TabIndex = 0
        '
        'uxMapInfoGroupBox
        '
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxAnomalyLabel)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxPageLabel)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxScaleLabel)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxReliabilityLabel)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxSuffixTypeLabel)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxSuffixNumberLabel)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxMapNumberLabel)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxAnomaly)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxPage)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxScale)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxReliability)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxCountyLabel)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxSuffixType)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxSuffixNumber)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxMapNumber)
        Me.uxMapInfoGroupBox.Controls.Add(Me.uxCounty)
        Me.uxMapInfoGroupBox.Location = New System.Drawing.Point(203, 6)
        Me.uxMapInfoGroupBox.Name = "uxMapInfoGroupBox"
        Me.uxMapInfoGroupBox.Size = New System.Drawing.Size(253, 214)
        Me.uxMapInfoGroupBox.TabIndex = 8
        Me.uxMapInfoGroupBox.TabStop = False
        Me.uxMapInfoGroupBox.Text = "Map Information"
        '
        'uxAnomalyLabel
        '
        Me.uxAnomalyLabel.AutoSize = True
        Me.uxAnomalyLabel.Location = New System.Drawing.Point(6, 186)
        Me.uxAnomalyLabel.Name = "uxAnomalyLabel"
        Me.uxAnomalyLabel.Size = New System.Drawing.Size(50, 13)
        Me.uxAnomalyLabel.TabIndex = 16
        Me.uxAnomalyLabel.Text = "Anomaly:"
        '
        'uxPageLabel
        '
        Me.uxPageLabel.AutoSize = True
        Me.uxPageLabel.Location = New System.Drawing.Point(6, 165)
        Me.uxPageLabel.Name = "uxPageLabel"
        Me.uxPageLabel.Size = New System.Drawing.Size(35, 13)
        Me.uxPageLabel.TabIndex = 15
        Me.uxPageLabel.Text = "Page:"
        '
        'uxScaleLabel
        '
        Me.uxScaleLabel.AutoSize = True
        Me.uxScaleLabel.Location = New System.Drawing.Point(6, 140)
        Me.uxScaleLabel.Name = "uxScaleLabel"
        Me.uxScaleLabel.Size = New System.Drawing.Size(37, 13)
        Me.uxScaleLabel.TabIndex = 14
        Me.uxScaleLabel.Text = "Scale:"
        '
        'uxReliabilityLabel
        '
        Me.uxReliabilityLabel.AutoSize = True
        Me.uxReliabilityLabel.Location = New System.Drawing.Point(6, 115)
        Me.uxReliabilityLabel.Name = "uxReliabilityLabel"
        Me.uxReliabilityLabel.Size = New System.Drawing.Size(54, 13)
        Me.uxReliabilityLabel.TabIndex = 13
        Me.uxReliabilityLabel.Text = "Reliability:"
        '
        'uxSuffixTypeLabel
        '
        Me.uxSuffixTypeLabel.AutoSize = True
        Me.uxSuffixTypeLabel.Location = New System.Drawing.Point(6, 90)
        Me.uxSuffixTypeLabel.Name = "uxSuffixTypeLabel"
        Me.uxSuffixTypeLabel.Size = New System.Drawing.Size(87, 13)
        Me.uxSuffixTypeLabel.TabIndex = 12
        Me.uxSuffixTypeLabel.Text = "Map Suffix Type:"
        '
        'uxSuffixNumberLabel
        '
        Me.uxSuffixNumberLabel.AutoSize = True
        Me.uxSuffixNumberLabel.Location = New System.Drawing.Point(6, 66)
        Me.uxSuffixNumberLabel.Name = "uxSuffixNumberLabel"
        Me.uxSuffixNumberLabel.Size = New System.Drawing.Size(100, 13)
        Me.uxSuffixNumberLabel.TabIndex = 11
        Me.uxSuffixNumberLabel.Text = "Map Suffix Number:"
        '
        'uxMapNumberLabel
        '
        Me.uxMapNumberLabel.AutoSize = True
        Me.uxMapNumberLabel.Location = New System.Drawing.Point(6, 42)
        Me.uxMapNumberLabel.Name = "uxMapNumberLabel"
        Me.uxMapNumberLabel.Size = New System.Drawing.Size(71, 13)
        Me.uxMapNumberLabel.TabIndex = 9
        Me.uxMapNumberLabel.Text = "Map Number:"
        '
        'uxAnomaly
        '
        Me.uxAnomaly.Location = New System.Drawing.Point(113, 186)
        Me.uxAnomaly.Name = "uxAnomaly"
        Me.uxAnomaly.Size = New System.Drawing.Size(55, 20)
        Me.uxAnomaly.TabIndex = 8
        '
        'uxPage
        '
        Me.uxPage.Location = New System.Drawing.Point(113, 162)
        Me.uxPage.Name = "uxPage"
        Me.uxPage.Size = New System.Drawing.Size(55, 20)
        Me.uxPage.TabIndex = 7
        '
        'uxScale
        '
        Me.uxScale.FormattingEnabled = True
        Me.uxScale.Location = New System.Drawing.Point(113, 137)
        Me.uxScale.Name = "uxScale"
        Me.uxScale.Size = New System.Drawing.Size(130, 21)
        Me.uxScale.TabIndex = 6
        '
        'uxReliability
        '
        Me.uxReliability.FormattingEnabled = True
        Me.uxReliability.Location = New System.Drawing.Point(113, 112)
        Me.uxReliability.Name = "uxReliability"
        Me.uxReliability.Size = New System.Drawing.Size(130, 21)
        Me.uxReliability.TabIndex = 5
        '
        'uxCountyLabel
        '
        Me.uxCountyLabel.AutoSize = True
        Me.uxCountyLabel.Location = New System.Drawing.Point(6, 17)
        Me.uxCountyLabel.Name = "uxCountyLabel"
        Me.uxCountyLabel.Size = New System.Drawing.Size(43, 13)
        Me.uxCountyLabel.TabIndex = 4
        Me.uxCountyLabel.Text = "County:"
        '
        'uxSuffixType
        '
        Me.uxSuffixType.FormattingEnabled = True
        Me.uxSuffixType.Location = New System.Drawing.Point(113, 87)
        Me.uxSuffixType.Name = "uxSuffixType"
        Me.uxSuffixType.Size = New System.Drawing.Size(130, 21)
        Me.uxSuffixType.TabIndex = 3
        '
        'uxSuffixNumber
        '
        Me.uxSuffixNumber.Location = New System.Drawing.Point(113, 63)
        Me.uxSuffixNumber.Name = "uxSuffixNumber"
        Me.uxSuffixNumber.Size = New System.Drawing.Size(130, 20)
        Me.uxSuffixNumber.TabIndex = 2
        '
        'uxMapNumber
        '
        Me.uxMapNumber.Location = New System.Drawing.Point(113, 39)
        Me.uxMapNumber.Name = "uxMapNumber"
        Me.uxMapNumber.Size = New System.Drawing.Size(130, 20)
        Me.uxMapNumber.TabIndex = 1
        '
        'uxCounty
        '
        Me.uxCounty.FormattingEnabled = True
        Me.uxCounty.Location = New System.Drawing.Point(113, 14)
        Me.uxCounty.Name = "uxCounty"
        Me.uxCounty.Size = New System.Drawing.Size(70, 21)
        Me.uxCounty.TabIndex = 0
        '
        'uxORMAPNumberGroupBox
        '
        Me.uxORMAPNumberGroupBox.Controls.Add(Me.uxORMAPNumberLabel)
        Me.uxORMAPNumberGroupBox.Location = New System.Drawing.Point(203, 227)
        Me.uxORMAPNumberGroupBox.Name = "uxORMAPNumberGroupBox"
        Me.uxORMAPNumberGroupBox.Size = New System.Drawing.Size(253, 68)
        Me.uxORMAPNumberGroupBox.TabIndex = 9
        Me.uxORMAPNumberGroupBox.TabStop = False
        Me.uxORMAPNumberGroupBox.Text = "Preview"
        '
        'uxORMAPNumberLabel
        '
        Me.uxORMAPNumberLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.uxORMAPNumberLabel.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.uxORMAPNumberLabel.Location = New System.Drawing.Point(9, 25)
        Me.uxORMAPNumberLabel.Name = "uxORMAPNumberLabel"
        Me.uxORMAPNumberLabel.Size = New System.Drawing.Size(234, 25)
        Me.uxORMAPNumberLabel.TabIndex = 0
        Me.uxORMAPNumberLabel.Text = "2015.00S05.00W3600--T000"
        Me.uxORMAPNumberLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'uxHelp
        '
        Me.uxHelp.Location = New System.Drawing.Point(381, 301)
        Me.uxHelp.Name = "uxHelp"
        Me.uxHelp.Size = New System.Drawing.Size(75, 23)
        Me.uxHelp.TabIndex = 10
        Me.uxHelp.Text = "&Help"
        Me.uxHelp.UseVisualStyleBackColor = True
        '
        'uxEdit
        '
        Me.uxEdit.Location = New System.Drawing.Point(221, 301)
        Me.uxEdit.Name = "uxEdit"
        Me.uxEdit.Size = New System.Drawing.Size(75, 23)
        Me.uxEdit.TabIndex = 11
        Me.uxEdit.Text = "&Edit"
        Me.uxEdit.UseVisualStyleBackColor = True
        '
        'uxQuit
        '
        Me.uxQuit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.uxQuit.Location = New System.Drawing.Point(301, 301)
        Me.uxQuit.Name = "uxQuit"
        Me.uxQuit.Size = New System.Drawing.Size(75, 23)
        Me.uxQuit.TabIndex = 12
        Me.uxQuit.Text = "&Quit"
        Me.uxQuit.UseVisualStyleBackColor = True
        '
        'EditMapIndexForm
        '
        Me.AcceptButton = Me.uxEdit
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.uxQuit
        Me.ClientSize = New System.Drawing.Size(467, 333)
        Me.Controls.Add(Me.uxQuit)
        Me.Controls.Add(Me.uxEdit)
        Me.Controls.Add(Me.uxHelp)
        Me.Controls.Add(Me.uxORMAPNumberGroupBox)
        Me.Controls.Add(Me.uxMapInfoGroupBox)
        Me.Controls.Add(Me.uxSectionGroupBox)
        Me.Controls.Add(Me.uxRangeGroupBox)
        Me.Controls.Add(Me.uxTownshipGroupBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EditMapIndexForm"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Edit Map Index"
        Me.TopMost = True
        Me.uxTownshipGroupBox.ResumeLayout(False)
        Me.uxTownshipGroupBox.PerformLayout()
        Me.uxRangeGroupBox.ResumeLayout(False)
        Me.uxRangeGroupBox.PerformLayout()
        Me.uxSectionGroupBox.ResumeLayout(False)
        Me.uxSectionGroupBox.PerformLayout()
        Me.uxMapInfoGroupBox.ResumeLayout(False)
        Me.uxMapInfoGroupBox.PerformLayout()
        Me.uxORMAPNumberGroupBox.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents uxTownshipDirectionalLabel As System.Windows.Forms.Label
    Friend WithEvents uxTownshipLabel As System.Windows.Forms.Label
    Friend WithEvents uxTownshipPartial As System.Windows.Forms.ComboBox
    Friend WithEvents uxTownshipDirectional As System.Windows.Forms.ComboBox
    Friend WithEvents uxTownship As System.Windows.Forms.ComboBox
    Friend WithEvents uxTownshipPartialLabel As System.Windows.Forms.Label
    Friend WithEvents uxTownshipGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents uxRangeGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents uxSectionGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents uxRangeDirectional As System.Windows.Forms.ComboBox
    Friend WithEvents uxRange As System.Windows.Forms.ComboBox
    Friend WithEvents uxMapInfoGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents uxORMAPNumberGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents uxRangeLabel As System.Windows.Forms.Label
    Friend WithEvents uxRangePartial As System.Windows.Forms.ComboBox
    Friend WithEvents uxRangeDirectionalLabel As System.Windows.Forms.Label
    Friend WithEvents uxRangePartialLabel As System.Windows.Forms.Label
    Friend WithEvents uxSectionLabel As System.Windows.Forms.Label
    Friend WithEvents uxSectionQtrQtr As System.Windows.Forms.ComboBox
    Friend WithEvents uxSectionQtr As System.Windows.Forms.ComboBox
    Friend WithEvents uxSection As System.Windows.Forms.ComboBox
    Friend WithEvents uxSectionQtrQtrLabel As System.Windows.Forms.Label
    Friend WithEvents uxSectionQtrLabel As System.Windows.Forms.Label
    Friend WithEvents uxSuffixNumber As System.Windows.Forms.TextBox
    Friend WithEvents uxMapNumber As System.Windows.Forms.TextBox
    Friend WithEvents uxCounty As System.Windows.Forms.ComboBox
    Friend WithEvents uxSuffixType As System.Windows.Forms.ComboBox
    Friend WithEvents uxAnomaly As System.Windows.Forms.TextBox
    Friend WithEvents uxPage As System.Windows.Forms.TextBox
    Friend WithEvents uxScale As System.Windows.Forms.ComboBox
    Friend WithEvents uxReliability As System.Windows.Forms.ComboBox
    Friend WithEvents uxCountyLabel As System.Windows.Forms.Label
    Friend WithEvents uxORMAPNumberLabel As System.Windows.Forms.Label
    Friend WithEvents uxPageLabel As System.Windows.Forms.Label
    Friend WithEvents uxScaleLabel As System.Windows.Forms.Label
    Friend WithEvents uxReliabilityLabel As System.Windows.Forms.Label
    Friend WithEvents uxSuffixTypeLabel As System.Windows.Forms.Label
    Friend WithEvents uxSuffixNumberLabel As System.Windows.Forms.Label
    Friend WithEvents uxMapNumberLabel As System.Windows.Forms.Label
    Friend WithEvents uxAnomalyLabel As System.Windows.Forms.Label
    Friend WithEvents uxHelp As System.Windows.Forms.Button
    Friend WithEvents uxEdit As System.Windows.Forms.Button
    Friend WithEvents uxQuit As System.Windows.Forms.Button
End Class
