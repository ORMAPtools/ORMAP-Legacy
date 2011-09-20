<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class uxSpiralCurve
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.uxTemplate = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.uxBeginRadius = New System.Windows.Forms.TextBox()
        Me.uxEndRadius = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.uxByAngle = New System.Windows.Forms.RadioButton()
        Me.uxByArcLength = New System.Windows.Forms.RadioButton()
        Me.uxByAngleValue = New System.Windows.Forms.TextBox()
        Me.uxArcLength = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.uxFromPointX = New System.Windows.Forms.TextBox()
        Me.uxFromPointY = New System.Windows.Forms.TextBox()
        Me.uxGetFromPoint = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.uxGetTangentPoint = New System.Windows.Forms.Button()
        Me.uxApply = New System.Windows.Forms.Button()
        Me.uxCancel = New System.Windows.Forms.Button()
        Me.uxHelp = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Template:"
        '
        'uxTemplate
        '
        Me.uxTemplate.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.uxTemplate.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.uxTemplate.Location = New System.Drawing.Point(74, 10)
        Me.uxTemplate.Name = "uxTemplate"
        Me.uxTemplate.Size = New System.Drawing.Size(198, 20)
        Me.uxTemplate.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Begin Radius:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(24, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "End Radius:"
        '
        'uxBeginRadius
        '
        Me.uxBeginRadius.Location = New System.Drawing.Point(96, 43)
        Me.uxBeginRadius.Name = "uxBeginRadius"
        Me.uxBeginRadius.Size = New System.Drawing.Size(100, 20)
        Me.uxBeginRadius.TabIndex = 4
        '
        'uxEndRadius
        '
        Me.uxEndRadius.Location = New System.Drawing.Point(96, 70)
        Me.uxEndRadius.Name = "uxEndRadius"
        Me.uxEndRadius.Size = New System.Drawing.Size(100, 20)
        Me.uxEndRadius.TabIndex = 5
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.uxArcLength)
        Me.GroupBox1.Controls.Add(Me.uxByAngleValue)
        Me.GroupBox1.Controls.Add(Me.uxByArcLength)
        Me.GroupBox1.Controls.Add(Me.uxByAngle)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 115)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(256, 74)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Create Spiral by"
        '
        'uxByAngle
        '
        Me.uxByAngle.AutoSize = True
        Me.uxByAngle.Checked = True
        Me.uxByAngle.Location = New System.Drawing.Point(11, 20)
        Me.uxByAngle.Name = "uxByAngle"
        Me.uxByAngle.Size = New System.Drawing.Size(83, 17)
        Me.uxByAngle.TabIndex = 0
        Me.uxByAngle.TabStop = True
        Me.uxByAngle.Text = "Delta Angle:"
        Me.uxByAngle.UseVisualStyleBackColor = True
        '
        'uxByArcLength
        '
        Me.uxByArcLength.AutoSize = True
        Me.uxByArcLength.Location = New System.Drawing.Point(11, 44)
        Me.uxByArcLength.Name = "uxByArcLength"
        Me.uxByArcLength.Size = New System.Drawing.Size(80, 17)
        Me.uxByArcLength.TabIndex = 1
        Me.uxByArcLength.Text = "Arc Length:"
        Me.uxByArcLength.UseVisualStyleBackColor = True
        '
        'uxByAngleValue
        '
        Me.uxByAngleValue.Location = New System.Drawing.Point(101, 20)
        Me.uxByAngleValue.Name = "uxByAngleValue"
        Me.uxByAngleValue.Size = New System.Drawing.Size(100, 20)
        Me.uxByAngleValue.TabIndex = 2
        '
        'uxArcLength
        '
        Me.uxArcLength.Location = New System.Drawing.Point(101, 44)
        Me.uxArcLength.Name = "uxArcLength"
        Me.uxArcLength.Size = New System.Drawing.Size(100, 20)
        Me.uxArcLength.TabIndex = 3
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.uxGetFromPoint)
        Me.GroupBox2.Controls.Add(Me.uxFromPointY)
        Me.GroupBox2.Controls.Add(Me.uxFromPointX)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Location = New System.Drawing.Point(19, 196)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(253, 54)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "From Point"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(17, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "X:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 35)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Y:"
        '
        'uxFromPointX
        '
        Me.uxFromPointX.Location = New System.Drawing.Point(32, 12)
        Me.uxFromPointX.Name = "uxFromPointX"
        Me.uxFromPointX.Size = New System.Drawing.Size(166, 20)
        Me.uxFromPointX.TabIndex = 2
        '
        'uxFromPointY
        '
        Me.uxFromPointY.Location = New System.Drawing.Point(32, 32)
        Me.uxFromPointY.Name = "uxFromPointY"
        Me.uxFromPointY.Size = New System.Drawing.Size(166, 20)
        Me.uxFromPointY.TabIndex = 3
        '
        'uxGetFromPoint
        '
        Me.uxGetFromPoint.Location = New System.Drawing.Point(221, 20)
        Me.uxGetFromPoint.Name = "uxGetFromPoint"
        Me.uxGetFromPoint.Size = New System.Drawing.Size(26, 23)
        Me.uxGetFromPoint.TabIndex = 4
        Me.uxGetFromPoint.Text = "uxGetFromPoint"
        Me.uxGetFromPoint.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.uxGetTangentPoint)
        Me.GroupBox3.Controls.Add(Me.TextBox1)
        Me.GroupBox3.Controls.Add(Me.TextBox2)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Location = New System.Drawing.Point(19, 256)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(253, 74)
        Me.GroupBox3.TabIndex = 8
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Tangent Point"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(17, 13)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Y:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 29)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(17, 13)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "X:"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(32, 45)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(166, 20)
        Me.TextBox1.TabIndex = 10
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(32, 25)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(166, 20)
        Me.TextBox2.TabIndex = 9
        '
        'uxGetTangentPoint
        '
        Me.uxGetTangentPoint.Location = New System.Drawing.Point(221, 29)
        Me.uxGetTangentPoint.Name = "uxGetTangentPoint"
        Me.uxGetTangentPoint.Size = New System.Drawing.Size(26, 23)
        Me.uxGetTangentPoint.TabIndex = 11
        Me.uxGetTangentPoint.Text = "Button2"
        Me.uxGetTangentPoint.UseVisualStyleBackColor = True
        '
        'uxApply
        '
        Me.uxApply.Location = New System.Drawing.Point(19, 359)
        Me.uxApply.Name = "uxApply"
        Me.uxApply.Size = New System.Drawing.Size(75, 23)
        Me.uxApply.TabIndex = 9
        Me.uxApply.Text = "Apply"
        Me.uxApply.UseVisualStyleBackColor = True
        '
        'uxCancel
        '
        Me.uxCancel.Location = New System.Drawing.Point(101, 358)
        Me.uxCancel.Name = "uxCancel"
        Me.uxCancel.Size = New System.Drawing.Size(75, 23)
        Me.uxCancel.TabIndex = 10
        Me.uxCancel.Text = "Cancel"
        Me.uxCancel.UseVisualStyleBackColor = True
        '
        'uxHelp
        '
        Me.uxHelp.Location = New System.Drawing.Point(183, 358)
        Me.uxHelp.Name = "uxHelp"
        Me.uxHelp.Size = New System.Drawing.Size(75, 23)
        Me.uxHelp.TabIndex = 11
        Me.uxHelp.Text = "Help"
        Me.uxHelp.UseVisualStyleBackColor = True
        '
        'uxSpiralCurve
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 400)
        Me.Controls.Add(Me.uxHelp)
        Me.Controls.Add(Me.uxCancel)
        Me.Controls.Add(Me.uxApply)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.uxEndRadius)
        Me.Controls.Add(Me.uxBeginRadius)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.uxTemplate)
        Me.Controls.Add(Me.Label1)
        Me.Name = "uxSpiralCurve"
        Me.Text = "Spiral Curve"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents uxTemplate As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents uxBeginRadius As System.Windows.Forms.TextBox
    Friend WithEvents uxEndRadius As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents uxByArcLength As System.Windows.Forms.RadioButton
    Friend WithEvents uxByAngle As System.Windows.Forms.RadioButton
    Friend WithEvents uxArcLength As System.Windows.Forms.TextBox
    Friend WithEvents uxByAngleValue As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents uxFromPointY As System.Windows.Forms.TextBox
    Friend WithEvents uxFromPointX As System.Windows.Forms.TextBox
    Friend WithEvents uxGetFromPoint As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents uxGetTangentPoint As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents uxApply As System.Windows.Forms.Button
    Friend WithEvents uxCancel As System.Windows.Forms.Button
    Friend WithEvents uxHelp As System.Windows.Forms.Button
End Class
