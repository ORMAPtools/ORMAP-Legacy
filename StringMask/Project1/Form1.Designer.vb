<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form1
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents txtMask As System.Windows.Forms.TextBox
	Public WithEvents cmdApplyMask As System.Windows.Forms.Button
	Public WithEvents txtOut As System.Windows.Forms.TextBox
	Public WithEvents txtIn As System.Windows.Forms.TextBox
	Public WithEvents lblInfo As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMask = New System.Windows.Forms.TextBox
        Me.cmdApplyMask = New System.Windows.Forms.Button
        Me.txtOut = New System.Windows.Forms.TextBox
        Me.txtIn = New System.Windows.Forms.TextBox
        Me.lblInfo = New System.Windows.Forms.Label
        Me._Label1_2 = New System.Windows.Forms.Label
        Me._Label1_1 = New System.Windows.Forms.Label
        Me._Label1_0 = New System.Windows.Forms.Label
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtNewOutput = New System.Windows.Forms.TextBox
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtMask
        '
        Me.txtMask.AcceptsReturn = True
        Me.txtMask.BackColor = System.Drawing.SystemColors.Window
        Me.txtMask.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMask.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMask.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMask.Location = New System.Drawing.Point(32, 96)
        Me.txtMask.MaxLength = 0
        Me.txtMask.Name = "txtMask"
        Me.txtMask.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMask.Size = New System.Drawing.Size(137, 20)
        Me.txtMask.TabIndex = 3
        Me.txtMask.Text = "TR^DSSQQ@@@@@"
        '
        'cmdApplyMask
        '
        Me.cmdApplyMask.BackColor = System.Drawing.SystemColors.Control
        Me.cmdApplyMask.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdApplyMask.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdApplyMask.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdApplyMask.Location = New System.Drawing.Point(32, 128)
        Me.cmdApplyMask.Name = "cmdApplyMask"
        Me.cmdApplyMask.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdApplyMask.Size = New System.Drawing.Size(153, 25)
        Me.cmdApplyMask.TabIndex = 2
        Me.cmdApplyMask.Text = "Apply &Mask"
        Me.cmdApplyMask.UseVisualStyleBackColor = False
        '
        'txtOut
        '
        Me.txtOut.AcceptsReturn = True
        Me.txtOut.BackColor = System.Drawing.SystemColors.Window
        Me.txtOut.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOut.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOut.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOut.Location = New System.Drawing.Point(280, 32)
        Me.txtOut.MaxLength = 32
        Me.txtOut.Name = "txtOut"
        Me.txtOut.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOut.Size = New System.Drawing.Size(161, 20)
        Me.txtOut.TabIndex = 1
        '
        'txtIn
        '
        Me.txtIn.AcceptsReturn = True
        Me.txtIn.BackColor = System.Drawing.SystemColors.Window
        Me.txtIn.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtIn.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIn.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtIn.Location = New System.Drawing.Point(32, 32)
        Me.txtIn.MaxLength = 0
        Me.txtIn.Name = "txtIn"
        Me.txtIn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtIn.Size = New System.Drawing.Size(193, 20)
        Me.txtIn.TabIndex = 0
        Me.txtIn.Text = "0313.00S13.00E23CC--000004900"
        '
        'lblInfo
        '
        Me.lblInfo.BackColor = System.Drawing.SystemColors.Control
        Me.lblInfo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblInfo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInfo.Location = New System.Drawing.Point(16, 184)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblInfo.Size = New System.Drawing.Size(473, 65)
        Me.lblInfo.TabIndex = 7
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_2, CType(2, Short))
        Me._Label1_2.Location = New System.Drawing.Point(32, 72)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(145, 17)
        Me._Label1_2.TabIndex = 6
        Me._Label1_2.Text = "Mask"
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(304, 8)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(65, 17)
        Me._Label1_1.TabIndex = 5
        Me._Label1_1.Text = "Old Output"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(40, 8)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(65, 17)
        Me._Label1_0.TabIndex = 4
        Me._Label1_0.Text = "Input"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(307, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 14)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "New output"
        '
        'txtNewOutput
        '
        Me.txtNewOutput.Location = New System.Drawing.Point(280, 101)
        Me.txtNewOutput.MaxLength = 32
        Me.txtNewOutput.Name = "txtNewOutput"
        Me.txtNewOutput.Size = New System.Drawing.Size(161, 20)
        Me.txtNewOutput.TabIndex = 9
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(516, 256)
        Me.Controls.Add(Me.txtNewOutput)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtMask)
        Me.Controls.Add(Me.cmdApplyMask)
        Me.Controls.Add(Me.txtOut)
        Me.Controls.Add(Me.txtIn)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me._Label1_2)
        Me.Controls.Add(Me._Label1_1)
        Me.Controls.Add(Me._Label1_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(120, 207)
        Me.Name = "Form1"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "ORMAP MapNumber tester"
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNewOutput As System.Windows.Forms.TextBox
#End Region 
End Class