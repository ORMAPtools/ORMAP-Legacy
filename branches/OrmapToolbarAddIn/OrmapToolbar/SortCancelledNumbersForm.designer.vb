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
        Me.uxCancelledNumbers = New System.Windows.Forms.ListBox()
        Me.uxTop = New System.Windows.Forms.Button()
        Me.uxUp = New System.Windows.Forms.Button()
        Me.uxDown = New System.Windows.Forms.Button()
        Me.uxBottom = New System.Windows.Forms.Button()
        Me.uxCancel = New System.Windows.Forms.Button()
        Me.uxOK = New System.Windows.Forms.Button()
        Me.uxAdd = New System.Windows.Forms.Button()
        Me.uxDelete = New System.Windows.Forms.Button()
        Me.uxMapIndex = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.uxFind = New System.Windows.Forms.Button()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.LineShape1 = New Microsoft.VisualBasic.PowerPacks.LineShape()
        Me.SuspendLayout()
        '
        'uxCancelledNumbers
        '
        Me.uxCancelledNumbers.FormattingEnabled = True
        Me.uxCancelledNumbers.Location = New System.Drawing.Point(12, 51)
        Me.uxCancelledNumbers.Name = "uxCancelledNumbers"
        Me.uxCancelledNumbers.Size = New System.Drawing.Size(225, 381)
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
        'uxMapIndex
        '
        Me.uxMapIndex.AllowDrop = True
        Me.uxMapIndex.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.uxMapIndex.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.uxMapIndex.Location = New System.Drawing.Point(114, 12)
        Me.uxMapIndex.Name = "uxMapIndex"
        Me.uxMapIndex.Size = New System.Drawing.Size(123, 20)
        Me.uxMapIndex.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(99, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Enter Map Number:"
        '
        'uxFind
        '
        Me.uxFind.Location = New System.Drawing.Point(243, 10)
        Me.uxFind.Name = "uxFind"
        Me.uxFind.Size = New System.Drawing.Size(52, 23)
        Me.uxFind.TabIndex = 11
        Me.uxFind.Text = "Find"
        Me.uxFind.UseVisualStyleBackColor = True
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineShape1})
        Me.ShapeContainer1.Size = New System.Drawing.Size(308, 473)
        Me.ShapeContainer1.TabIndex = 12
        Me.ShapeContainer1.TabStop = False
        '
        'LineShape1
        '
        Me.LineShape1.Name = "LineShape1"
        Me.LineShape1.X1 = 6
        Me.LineShape1.X2 = 301
        Me.LineShape1.Y1 = 41
        Me.LineShape1.Y2 = 41
        '
        'SortCancelledNumbersForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(308, 473)
        Me.Controls.Add(Me.uxFind)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.uxMapIndex)
        Me.Controls.Add(Me.uxDelete)
        Me.Controls.Add(Me.uxAdd)
        Me.Controls.Add(Me.uxOK)
        Me.Controls.Add(Me.uxCancel)
        Me.Controls.Add(Me.uxBottom)
        Me.Controls.Add(Me.uxDown)
        Me.Controls.Add(Me.uxUp)
        Me.Controls.Add(Me.uxTop)
        Me.Controls.Add(Me.uxCancelledNumbers)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "SortCancelledNumbersForm"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Sort Cancelled Numbers"
        Me.ResumeLayout(False)
        Me.PerformLayout()

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
    Friend WithEvents uxMapIndex As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents uxFind As System.Windows.Forms.Button
    Friend WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    Friend WithEvents LineShape1 As Microsoft.VisualBasic.PowerPacks.LineShape
End Class
