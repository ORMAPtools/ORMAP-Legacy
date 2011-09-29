<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SpiralCurveSpiralDockWindow
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.uxGettoPoint = New System.Windows.Forms.Button()
        Me.uxToPointYValue = New System.Windows.Forms.TextBox()
        Me.uxToPointXValue = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.uxGetFromPoint = New System.Windows.Forms.Button()
        Me.uxFromPointYValue = New System.Windows.Forms.TextBox()
        Me.uxFromPointXValue = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.uxGetTangentPoint = New System.Windows.Forms.Button()
        Me.uxTangentPointYValue = New System.Windows.Forms.TextBox()
        Me.uxTangentPointXValue = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.uxCurveDegreeValue = New System.Windows.Forms.TextBox()
        Me.uxCurveByRadiusValue = New System.Windows.Forms.TextBox()
        Me.uxCurvebyDegree = New System.Windows.Forms.RadioButton()
        Me.uxCurveByRadius = New System.Windows.Forms.RadioButton()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.uxCurvetotheLeft = New System.Windows.Forms.RadioButton()
        Me.uxCurvetotheRight = New System.Windows.Forms.RadioButton()
        Me.uxArcLengthValue = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.SuspendLayout()
        '
        'uxCreate
        '
        Me.uxCreate.Location = New System.Drawing.Point(2, 383)
        Me.uxCreate.Name = "uxCreate"
        Me.uxCreate.Size = New System.Drawing.Size(75, 23)
        Me.uxCreate.TabIndex = 0
        Me.uxCreate.Text = "Create"
        Me.uxCreate.UseVisualStyleBackColor = True
        '
        'uxHelp
        '
        Me.uxHelp.Location = New System.Drawing.Point(83, 383)
        Me.uxHelp.Name = "uxHelp"
        Me.uxHelp.Size = New System.Drawing.Size(75, 23)
        Me.uxHelp.TabIndex = 2
        Me.uxHelp.Text = "Help"
        Me.uxHelp.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.uxGettoPoint)
        Me.GroupBox1.Controls.Add(Me.uxToPointYValue)
        Me.GroupBox1.Controls.Add(Me.uxToPointXValue)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(0, 308)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(267, 69)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "To Point"
        '
        'uxGettoPoint
        '
        Me.uxGettoPoint.Image = Global.SpiralTools.My.Resources.Resources._3DCenterOnTarget32
        Me.uxGettoPoint.Location = New System.Drawing.Point(200, 25)
        Me.uxGettoPoint.Name = "uxGettoPoint"
        Me.uxGettoPoint.Size = New System.Drawing.Size(35, 32)
        Me.uxGettoPoint.TabIndex = 4
        Me.uxGettoPoint.UseVisualStyleBackColor = True
        '
        'uxToPointYValue
        '
        Me.uxToPointYValue.Location = New System.Drawing.Point(31, 44)
        Me.uxToPointYValue.Name = "uxToPointYValue"
        Me.uxToPointYValue.Size = New System.Drawing.Size(124, 20)
        Me.uxToPointYValue.TabIndex = 3
        '
        'uxToPointXValue
        '
        Me.uxToPointXValue.Location = New System.Drawing.Point(31, 20)
        Me.uxToPointXValue.Name = "uxToPointXValue"
        Me.uxToPointXValue.Size = New System.Drawing.Size(124, 20)
        Me.uxToPointXValue.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(17, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Y:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(17, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "X:"
        '
        'uxGetFromPoint
        '
        Me.uxGetFromPoint.Image = Global.SpiralTools.My.Resources.Resources._3DCenterOnTarget32
        Me.uxGetFromPoint.Location = New System.Drawing.Point(197, 22)
        Me.uxGetFromPoint.Name = "uxGetFromPoint"
        Me.uxGetFromPoint.Size = New System.Drawing.Size(35, 35)
        Me.uxGetFromPoint.TabIndex = 4
        Me.uxGetFromPoint.UseVisualStyleBackColor = True
        '
        'uxFromPointYValue
        '
        Me.uxFromPointYValue.Location = New System.Drawing.Point(31, 44)
        Me.uxFromPointYValue.Name = "uxFromPointYValue"
        Me.uxFromPointYValue.Size = New System.Drawing.Size(124, 20)
        Me.uxFromPointYValue.TabIndex = 3
        '
        'uxFromPointXValue
        '
        Me.uxFromPointXValue.Location = New System.Drawing.Point(31, 20)
        Me.uxFromPointXValue.Name = "uxFromPointXValue"
        Me.uxFromPointXValue.Size = New System.Drawing.Size(124, 20)
        Me.uxFromPointXValue.TabIndex = 2
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.uxGetFromPoint)
        Me.GroupBox2.Controls.Add(Me.uxFromPointYValue)
        Me.GroupBox2.Controls.Add(Me.uxFromPointXValue)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 158)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(267, 69)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "From Point"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(17, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Y:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 23)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(17, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "X:"
        '
        'uxGetTangentPoint
        '
        Me.uxGetTangentPoint.Image = Global.SpiralTools.My.Resources.Resources._3DCenterOnTarget32
        Me.uxGetTangentPoint.Location = New System.Drawing.Point(197, 25)
        Me.uxGetTangentPoint.Name = "uxGetTangentPoint"
        Me.uxGetTangentPoint.Size = New System.Drawing.Size(35, 31)
        Me.uxGetTangentPoint.TabIndex = 4
        Me.uxGetTangentPoint.UseVisualStyleBackColor = True
        '
        'uxTangentPointYValue
        '
        Me.uxTangentPointYValue.Location = New System.Drawing.Point(31, 44)
        Me.uxTangentPointYValue.Name = "uxTangentPointYValue"
        Me.uxTangentPointYValue.Size = New System.Drawing.Size(124, 20)
        Me.uxTangentPointYValue.TabIndex = 3
        '
        'uxTangentPointXValue
        '
        Me.uxTangentPointXValue.Location = New System.Drawing.Point(31, 20)
        Me.uxTangentPointXValue.Name = "uxTangentPointXValue"
        Me.uxTangentPointXValue.Size = New System.Drawing.Size(124, 20)
        Me.uxTangentPointXValue.TabIndex = 2
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.uxGetTangentPoint)
        Me.GroupBox3.Controls.Add(Me.uxTangentPointYValue)
        Me.GroupBox3.Controls.Add(Me.uxTangentPointXValue)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Location = New System.Drawing.Point(3, 233)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(267, 69)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Tangent Point"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Y:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(7, 23)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(17, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "X:"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.uxCurveDegreeValue)
        Me.GroupBox4.Controls.Add(Me.uxCurveByRadiusValue)
        Me.GroupBox4.Controls.Add(Me.uxCurvebyDegree)
        Me.GroupBox4.Controls.Add(Me.uxCurveByRadius)
        Me.GroupBox4.Location = New System.Drawing.Point(6, 75)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(257, 77)
        Me.GroupBox4.TabIndex = 6
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Define Circular Curve using:"
        '
        'uxCurveDegreeValue
        '
        Me.uxCurveDegreeValue.Enabled = False
        Me.uxCurveDegreeValue.Location = New System.Drawing.Point(120, 49)
        Me.uxCurveDegreeValue.Name = "uxCurveDegreeValue"
        Me.uxCurveDegreeValue.Size = New System.Drawing.Size(124, 20)
        Me.uxCurveDegreeValue.TabIndex = 3
        '
        'uxCurveByRadiusValue
        '
        Me.uxCurveByRadiusValue.Location = New System.Drawing.Point(71, 22)
        Me.uxCurveByRadiusValue.Name = "uxCurveByRadiusValue"
        Me.uxCurveByRadiusValue.Size = New System.Drawing.Size(173, 20)
        Me.uxCurveByRadiusValue.TabIndex = 2
        '
        'uxCurvebyDegree
        '
        Me.uxCurvebyDegree.AutoSize = True
        Me.uxCurvebyDegree.Location = New System.Drawing.Point(7, 49)
        Me.uxCurvebyDegree.Name = "uxCurvebyDegree"
        Me.uxCurvebyDegree.Size = New System.Drawing.Size(106, 17)
        Me.uxCurvebyDegree.TabIndex = 1
        Me.uxCurvebyDegree.Text = "Degree of Curve:"
        Me.uxCurvebyDegree.UseVisualStyleBackColor = True
        '
        'uxCurveByRadius
        '
        Me.uxCurveByRadius.AutoSize = True
        Me.uxCurveByRadius.Checked = True
        Me.uxCurveByRadius.Location = New System.Drawing.Point(7, 22)
        Me.uxCurveByRadius.Name = "uxCurveByRadius"
        Me.uxCurveByRadius.Size = New System.Drawing.Size(61, 17)
        Me.uxCurveByRadius.TabIndex = 0
        Me.uxCurveByRadius.TabStop = True
        Me.uxCurveByRadius.Text = "Radius:"
        Me.uxCurveByRadius.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 52)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(94, 13)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Lengths of Spirals:"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.uxCurvetotheLeft)
        Me.GroupBox6.Controls.Add(Me.uxCurvetotheRight)
        Me.GroupBox6.Location = New System.Drawing.Point(3, 9)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(247, 34)
        Me.GroupBox6.TabIndex = 10
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Curve to the"
        '
        'uxCurvetotheLeft
        '
        Me.uxCurvetotheLeft.AutoSize = True
        Me.uxCurvetotheLeft.Location = New System.Drawing.Point(63, 12)
        Me.uxCurvetotheLeft.Name = "uxCurvetotheLeft"
        Me.uxCurvetotheLeft.Size = New System.Drawing.Size(43, 17)
        Me.uxCurvetotheLeft.TabIndex = 1
        Me.uxCurvetotheLeft.TabStop = True
        Me.uxCurvetotheLeft.Text = "Left"
        Me.uxCurvetotheLeft.UseVisualStyleBackColor = True
        '
        'uxCurvetotheRight
        '
        Me.uxCurvetotheRight.AutoSize = True
        Me.uxCurvetotheRight.Checked = True
        Me.uxCurvetotheRight.Location = New System.Drawing.Point(7, 12)
        Me.uxCurvetotheRight.Name = "uxCurvetotheRight"
        Me.uxCurvetotheRight.Size = New System.Drawing.Size(50, 17)
        Me.uxCurvetotheRight.TabIndex = 0
        Me.uxCurvetotheRight.TabStop = True
        Me.uxCurvetotheRight.Text = "Right"
        Me.uxCurvetotheRight.UseVisualStyleBackColor = True
        '
        'uxArcLengthValue
        '
        Me.uxArcLengthValue.Location = New System.Drawing.Point(109, 49)
        Me.uxArcLengthValue.Name = "uxArcLengthValue"
        Me.uxArcLengthValue.Size = New System.Drawing.Size(141, 20)
        Me.uxArcLengthValue.TabIndex = 2
        '
        'SpiralCurveSpiralDockWindow
        '
        Me.Controls.Add(Me.uxArcLengthValue)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.uxHelp)
        Me.Controls.Add(Me.uxCreate)
        Me.Name = "SpiralCurveSpiralDockWindow"
        Me.Size = New System.Drawing.Size(272, 414)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents uxCreate As System.Windows.Forms.Button
    Friend WithEvents uxHelp As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents uxGettoPoint As System.Windows.Forms.Button
    Friend WithEvents uxToPointYValue As System.Windows.Forms.TextBox
    Friend WithEvents uxToPointXValue As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents uxGetFromPoint As System.Windows.Forms.Button
    Friend WithEvents uxFromPointYValue As System.Windows.Forms.TextBox
    Friend WithEvents uxFromPointXValue As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents uxGetTangentPoint As System.Windows.Forms.Button
    Friend WithEvents uxTangentPointYValue As System.Windows.Forms.TextBox
    Friend WithEvents uxTangentPointXValue As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents uxCurvebyDegree As System.Windows.Forms.RadioButton
    Friend WithEvents uxCurveByRadius As System.Windows.Forms.RadioButton
    Friend WithEvents uxCurveDegreeValue As System.Windows.Forms.TextBox
    Friend WithEvents uxCurveByRadiusValue As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents uxCurvetotheRight As System.Windows.Forms.RadioButton
    Friend WithEvents uxCurvetotheLeft As System.Windows.Forms.RadioButton
    Friend WithEvents uxArcLengthValue As System.Windows.Forms.TextBox

End Class
