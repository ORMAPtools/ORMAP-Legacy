<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SortCancelledNumbersForm
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
        Me.uxCancelledNumbers = New System.Windows.Forms.ListBox
        Me.uxTop = New System.Windows.Forms.Button
        Me.uxUp = New System.Windows.Forms.Button
        Me.uxDown = New System.Windows.Forms.Button
        Me.uxBottom = New System.Windows.Forms.Button
        Me.uxCancel = New System.Windows.Forms.Button
        Me.uxOK = New System.Windows.Forms.Button
        Me.uxAdd = New System.Windows.Forms.Button
        Me.uxDelete = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'uxCancelledNumbers
        '
        Me.uxCancelledNumbers.FormattingEnabled = True
        Me.uxCancelledNumbers.Location = New System.Drawing.Point(12, 12)
        Me.uxCancelledNumbers.Name = "uxCancelledNumbers"
        Me.uxCancelledNumbers.Size = New System.Drawing.Size(225, 420)
        Me.uxCancelledNumbers.TabIndex = 0
        '
        'uxTop
        '
        Me.uxTop.Location = New System.Drawing.Point(243, 129)
        Me.uxTop.Name = "uxTop"
        Me.uxTop.Size = New System.Drawing.Size(52, 23)
        Me.uxTop.TabIndex = 1
        Me.uxTop.Text = "Top"
        Me.uxTop.UseVisualStyleBackColor = True
        '
        'uxUp
        '
        Me.uxUp.Location = New System.Drawing.Point(243, 158)
        Me.uxUp.Name = "uxUp"
        Me.uxUp.Size = New System.Drawing.Size(52, 23)
        Me.uxUp.TabIndex = 2
        Me.uxUp.Text = "Up"
        Me.uxUp.UseVisualStyleBackColor = True
        '
        'uxDown
        '
        Me.uxDown.Location = New System.Drawing.Point(243, 200)
        Me.uxDown.Name = "uxDown"
        Me.uxDown.Size = New System.Drawing.Size(52, 23)
        Me.uxDown.TabIndex = 3
        Me.uxDown.Text = "Down"
        Me.uxDown.UseVisualStyleBackColor = True
        '
        'uxBottom
        '
        Me.uxBottom.Location = New System.Drawing.Point(243, 229)
        Me.uxBottom.Name = "uxBottom"
        Me.uxBottom.Size = New System.Drawing.Size(52, 23)
        Me.uxBottom.TabIndex = 4
        Me.uxBottom.Text = "Bottom"
        Me.uxBottom.UseVisualStyleBackColor = True
        '
        'uxCancel
        '
        Me.uxCancel.Location = New System.Drawing.Point(220, 438)
        Me.uxCancel.Name = "uxCancel"
        Me.uxCancel.Size = New System.Drawing.Size(75, 23)
        Me.uxCancel.TabIndex = 5
        Me.uxCancel.Text = "Cancel"
        Me.uxCancel.UseVisualStyleBackColor = True
        '
        'uxOK
        '
        Me.uxOK.Location = New System.Drawing.Point(139, 438)
        Me.uxOK.Name = "uxOK"
        Me.uxOK.Size = New System.Drawing.Size(75, 23)
        Me.uxOK.TabIndex = 6
        Me.uxOK.Text = "OK"
        Me.uxOK.UseVisualStyleBackColor = True
        '
        'uxAdd
        '
        Me.uxAdd.Location = New System.Drawing.Point(243, 326)
        Me.uxAdd.Name = "uxAdd"
        Me.uxAdd.Size = New System.Drawing.Size(52, 23)
        Me.uxAdd.TabIndex = 7
        Me.uxAdd.Text = "Add"
        Me.uxAdd.UseVisualStyleBackColor = True
        '
        'uxDelete
        '
        Me.uxDelete.Location = New System.Drawing.Point(243, 355)
        Me.uxDelete.Name = "uxDelete"
        Me.uxDelete.Size = New System.Drawing.Size(53, 23)
        Me.uxDelete.TabIndex = 8
        Me.uxDelete.Text = "Delete"
        Me.uxDelete.UseVisualStyleBackColor = True
        '
        'SortCancelledNumbersForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(308, 473)
        Me.Controls.Add(Me.uxDelete)
        Me.Controls.Add(Me.uxAdd)
        Me.Controls.Add(Me.uxOK)
        Me.Controls.Add(Me.uxCancel)
        Me.Controls.Add(Me.uxBottom)
        Me.Controls.Add(Me.uxDown)
        Me.Controls.Add(Me.uxUp)
        Me.Controls.Add(Me.uxTop)
        Me.Controls.Add(Me.uxCancelledNumbers)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "SortCancelledNumbersForm"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Sort Cancelled Numbers"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents uxCancelledNumbers As System.Windows.Forms.ListBox
    Friend WithEvents uxTop As System.Windows.Forms.Button
    Friend WithEvents uxUp As System.Windows.Forms.Button
    Friend WithEvents uxDown As System.Windows.Forms.Button
    Friend WithEvents uxBottom As System.Windows.Forms.Button
    Friend WithEvents uxCancel As System.Windows.Forms.Button
    Friend WithEvents uxOK As System.Windows.Forms.Button
    Friend WithEvents uxAdd As System.Windows.Forms.Button
    Friend WithEvents uxDelete As System.Windows.Forms.Button
End Class
