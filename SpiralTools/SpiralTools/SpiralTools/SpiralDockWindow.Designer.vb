﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SpiralDockWindow
    Inherits System.Windows.Forms.UserControl

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
        Me.uxCreate = New System.Windows.Forms.Button()
        Me.uxHelp = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.uxTemplateValue = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.uxCurvetotheLeft = New System.Windows.Forms.RadioButton()
        Me.uxCurvetotheRight = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.uxDeltaAngle = New System.Windows.Forms.TextBox()
        Me.uxArcLenghtValue = New System.Windows.Forms.TextBox()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.uxByArcLength = New System.Windows.Forms.RadioButton()
        Me.uxBeginRadiusValue = New System.Windows.Forms.TextBox()
        Me.uxEndRadiusValue = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.uxGetFromPoint = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.usGetTangentPoint = New System.Windows.Forms.Button()
        Me.uxTangentPointYValue = New System.Windows.Forms.TextBox()
        Me.uxTangentPointXValue = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'uxCreate
        '
        Me.uxCreate.Location = New System.Drawing.Point(4, 398)
        Me.uxCreate.Name = "uxCreate"
        Me.uxCreate.Size = New System.Drawing.Size(75, 23)
        Me.uxCreate.TabIndex = 0
        Me.uxCreate.Text = "Create"
        Me.uxCreate.UseVisualStyleBackColor = True
        '
        'uxHelp
        '
        Me.uxHelp.Location = New System.Drawing.Point(85, 398)
        Me.uxHelp.Name = "uxHelp"
        Me.uxHelp.Size = New System.Drawing.Size(75, 23)
        Me.uxHelp.TabIndex = 2
        Me.uxHelp.Text = "Help"
        Me.uxHelp.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Template"
        '
        'uxTemplateValue
        '
        Me.uxTemplateValue.AllowDrop = True
        Me.uxTemplateValue.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.uxTemplateValue.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.uxTemplateValue.Location = New System.Drawing.Point(4, 21)
        Me.uxTemplateValue.Name = "uxTemplateValue"
        Me.uxTemplateValue.Size = New System.Drawing.Size(293, 20)
        Me.uxTemplateValue.TabIndex = 4
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.uxCurvetotheLeft)
        Me.GroupBox1.Controls.Add(Me.uxCurvetotheRight)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 50)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(287, 45)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Curve to the"
        '
        'uxCurvetotheLeft
        '
        Me.uxCurvetotheLeft.AutoSize = True
        Me.uxCurvetotheLeft.Location = New System.Drawing.Point(65, 19)
        Me.uxCurvetotheLeft.Name = "uxCurvetotheLeft"
        Me.uxCurvetotheLeft.Size = New System.Drawing.Size(43, 17)
        Me.uxCurvetotheLeft.TabIndex = 1
        Me.uxCurvetotheLeft.Text = "Left"
        Me.uxCurvetotheLeft.UseVisualStyleBackColor = True
        '
        'uxCurvetotheRight
        '
        Me.uxCurvetotheRight.AutoSize = True
        Me.uxCurvetotheRight.Checked = True
        Me.uxCurvetotheRight.Location = New System.Drawing.Point(7, 20)
        Me.uxCurvetotheRight.Name = "uxCurvetotheRight"
        Me.uxCurvetotheRight.Size = New System.Drawing.Size(50, 17)
        Me.uxCurvetotheRight.TabIndex = 0
        Me.uxCurvetotheRight.TabStop = True
        Me.uxCurvetotheRight.Text = "Right"
        Me.uxCurvetotheRight.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 106)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Begin Radius:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(29, 130)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "End Radius:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.uxDeltaAngle)
        Me.GroupBox2.Controls.Add(Me.uxArcLenghtValue)
        Me.GroupBox2.Controls.Add(Me.RadioButton2)
        Me.GroupBox2.Controls.Add(Me.uxByArcLength)
        Me.GroupBox2.Location = New System.Drawing.Point(13, 155)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(280, 80)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Construct Spiral using"
        '
        'uxDeltaAngle
        '
        Me.uxDeltaAngle.Location = New System.Drawing.Point(89, 41)
        Me.uxDeltaAngle.Name = "uxDeltaAngle"
        Me.uxDeltaAngle.Size = New System.Drawing.Size(179, 20)
        Me.uxDeltaAngle.TabIndex = 3
        '
        'uxArcLenghtValue
        '
        Me.uxArcLenghtValue.Location = New System.Drawing.Point(89, 19)
        Me.uxArcLenghtValue.Name = "uxArcLenghtValue"
        Me.uxArcLenghtValue.Size = New System.Drawing.Size(179, 20)
        Me.uxArcLenghtValue.TabIndex = 2
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(7, 44)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(80, 17)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "Delta Angle"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'uxByArcLength
        '
        Me.uxByArcLength.AutoSize = True
        Me.uxByArcLength.Checked = True
        Me.uxByArcLength.Location = New System.Drawing.Point(7, 20)
        Me.uxByArcLength.Name = "uxByArcLength"
        Me.uxByArcLength.Size = New System.Drawing.Size(76, 17)
        Me.uxByArcLength.TabIndex = 0
        Me.uxByArcLength.TabStop = True
        Me.uxByArcLength.Text = "Arc length:"
        Me.uxByArcLength.UseVisualStyleBackColor = True
        '
        'uxBeginRadiusValue
        '
        Me.uxBeginRadiusValue.Location = New System.Drawing.Point(98, 102)
        Me.uxBeginRadiusValue.Name = "uxBeginRadiusValue"
        Me.uxBeginRadiusValue.Size = New System.Drawing.Size(183, 20)
        Me.uxBeginRadiusValue.TabIndex = 9
        Me.uxBeginRadiusValue.Text = "Infinity"
        '
        'uxEndRadiusValue
        '
        Me.uxEndRadiusValue.Location = New System.Drawing.Point(98, 127)
        Me.uxEndRadiusValue.Name = "uxEndRadiusValue"
        Me.uxEndRadiusValue.Size = New System.Drawing.Size(183, 20)
        Me.uxEndRadiusValue.TabIndex = 10
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.uxGetFromPoint)
        Me.GroupBox3.Controls.Add(Me.TextBox2)
        Me.GroupBox3.Controls.Add(Me.TextBox1)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Location = New System.Drawing.Point(14, 242)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(279, 69)
        Me.GroupBox3.TabIndex = 11
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "From Point"
        '
        'uxGetFromPoint
        '
        Me.uxGetFromPoint.Location = New System.Drawing.Point(198, 26)
        Me.uxGetFromPoint.Name = "uxGetFromPoint"
        Me.uxGetFromPoint.Size = New System.Drawing.Size(75, 23)
        Me.uxGetFromPoint.TabIndex = 4
        Me.uxGetFromPoint.Text = "Get Point"
        Me.uxGetFromPoint.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(34, 41)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(157, 20)
        Me.TextBox2.TabIndex = 3
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(34, 20)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(157, 20)
        Me.TextBox1.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(10, 44)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Y:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(17, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "X:"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.usGetTangentPoint)
        Me.GroupBox4.Controls.Add(Me.uxTangentPointYValue)
        Me.GroupBox4.Controls.Add(Me.uxTangentPointXValue)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Controls.Add(Me.Label7)
        Me.GroupBox4.Location = New System.Drawing.Point(13, 317)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(279, 69)
        Me.GroupBox4.TabIndex = 12
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Tangent Point"
        '
        'usGetTangentPoint
        '
        Me.usGetTangentPoint.Location = New System.Drawing.Point(199, 28)
        Me.usGetTangentPoint.Name = "usGetTangentPoint"
        Me.usGetTangentPoint.Size = New System.Drawing.Size(75, 23)
        Me.usGetTangentPoint.TabIndex = 5
        Me.usGetTangentPoint.Text = "Get Point"
        Me.usGetTangentPoint.UseVisualStyleBackColor = True
        '
        'uxTangentPointYValue
        '
        Me.uxTangentPointYValue.Location = New System.Drawing.Point(34, 41)
        Me.uxTangentPointYValue.Name = "uxTangentPointYValue"
        Me.uxTangentPointYValue.Size = New System.Drawing.Size(157, 20)
        Me.uxTangentPointYValue.TabIndex = 3
        '
        'uxTangentPointXValue
        '
        Me.uxTangentPointXValue.Location = New System.Drawing.Point(34, 20)
        Me.uxTangentPointXValue.Name = "uxTangentPointXValue"
        Me.uxTangentPointXValue.Size = New System.Drawing.Size(157, 20)
        Me.uxTangentPointXValue.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(10, 44)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(17, 13)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Y:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(10, 20)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(17, 13)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "X:"
        '
        'SpiralDockWindow
        '
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.uxEndRadiusValue)
        Me.Controls.Add(Me.uxBeginRadiusValue)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.uxTemplateValue)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.uxHelp)
        Me.Controls.Add(Me.uxCreate)
        Me.Name = "SpiralDockWindow"
        Me.Size = New System.Drawing.Size(300, 436)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents uxCreate As System.Windows.Forms.Button
    Friend WithEvents uxHelp As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents uxTemplateValue As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents uxCurvetotheLeft As System.Windows.Forms.RadioButton
    Friend WithEvents uxCurvetotheRight As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents uxByArcLength As System.Windows.Forms.RadioButton
    Friend WithEvents uxBeginRadiusValue As System.Windows.Forms.TextBox
    Friend WithEvents uxEndRadiusValue As System.Windows.Forms.TextBox
    Friend WithEvents uxDeltaAngle As System.Windows.Forms.TextBox
    Friend WithEvents uxArcLenghtValue As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents uxTangentPointYValue As System.Windows.Forms.TextBox
    Friend WithEvents uxTangentPointXValue As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents uxGetFromPoint As System.Windows.Forms.Button
    Friend WithEvents usGetTangentPoint As System.Windows.Forms.Button

End Class
