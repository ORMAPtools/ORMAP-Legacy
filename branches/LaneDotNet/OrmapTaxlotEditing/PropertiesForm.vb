#Region "Copyright 2006 ESRI"

' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
'
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
'
' See use restrictions at /arcgis/developerkit/userestrictions.

#End Region

#Region "Copyright 2008 ORMAP Tech Group"

' File: PropertiesForm.vb

' Author: .NET Migration Team (Shad Campbell, James Moore, Nick Seigal)
' Created: January 8, 2008

' All rights reserved. Reproduction or transmission of this file, or a portion thereof,
' is forbidden without prior written permission of the ORMAP Tech Group.

#End Region

Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

''' <summary>
''' Summary description for PropertiesForm.
''' </summary>
<ComVisible(False)> _
Friend NotInheritable Class PropertiesForm : Inherits System.Windows.Forms.Form

#Region "Friend Fields"

    Friend WithEvents uxEnableTools As System.Windows.Forms.CheckBox
    Friend WithEvents uxEnableAutoUpdate As System.Windows.Forms.CheckBox
    Friend WithEvents uxDescription As System.Windows.Forms.Label
    Friend WithEvents uxLogo As System.Windows.Forms.PictureBox
    Friend WithEvents uxMinimumFieldsOption As System.Windows.Forms.RadioButton
    Friend WithEvents uxAllFieldsOption As System.Windows.Forms.RadioButton

#End Region

#Region "Private Fields"

    ''' <summary>
    ''' Required designer variable.
    ''' </summary>
    Private components As System.ComponentModel.Container = Nothing

#End Region

#Region "Constructors"

    Public Sub New()
        '
        ' Required for Windows Form Designer support
        '
        InitializeComponent()

        '
        ' TODO: Add any constructor code after InitializeComponent call
        '
    End Sub

#End Region

#Region "Protected Methods"

    ''' <summary>
    ''' Clean up any resources being used.
    ''' </summary>
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region

#Region "Windows Form Designer generated code"
    ''' <summary>
    ''' Required method for Designer support - do not modify
    ''' the contents of this method with the code editor.
    ''' </summary>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PropertiesForm))
        Me.uxDescription = New System.Windows.Forms.Label
        Me.uxLogo = New System.Windows.Forms.PictureBox
        Me.uxMinimumFieldsOption = New System.Windows.Forms.RadioButton
        Me.uxAllFieldsOption = New System.Windows.Forms.RadioButton
        Me.uxEnableAutoUpdate = New System.Windows.Forms.CheckBox
        Me.uxEnableTools = New System.Windows.Forms.CheckBox
        CType(Me.uxLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'uxDescription
        '
        Me.uxDescription.Location = New System.Drawing.Point(18, 23)
        Me.uxDescription.Name = "uxDescription"
        Me.uxDescription.Size = New System.Drawing.Size(370, 54)
        Me.uxDescription.TabIndex = 2
        Me.uxDescription.Text = "The ORMAP Taxlot Editing Toolbar gives the user tools to edit taxlots and related" & _
            " features in compliance with the ORMAP standard."
        '
        'uxLogo
        '
        Me.uxLogo.Image = CType(resources.GetObject("uxLogo.Image"), System.Drawing.Image)
        Me.uxLogo.Location = New System.Drawing.Point(21, 279)
        Me.uxLogo.Name = "uxLogo"
        Me.uxLogo.Size = New System.Drawing.Size(86, 82)
        Me.uxLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.uxLogo.TabIndex = 3
        Me.uxLogo.TabStop = False
        '
        'uxMinimumFieldsOption
        '
        Me.uxMinimumFieldsOption.AutoSize = True
        Me.uxMinimumFieldsOption.Location = New System.Drawing.Point(60, 136)
        Me.uxMinimumFieldsOption.Name = "uxMinimumFieldsOption"
        Me.uxMinimumFieldsOption.Size = New System.Drawing.Size(285, 17)
        Me.uxMinimumFieldsOption.TabIndex = 4
        Me.uxMinimumFieldsOption.Text = "Minimum fields only (e.g. AUTODATE and AUTOWHO)"
        Me.uxMinimumFieldsOption.UseVisualStyleBackColor = True
        '
        'uxAllFieldsOption
        '
        Me.uxAllFieldsOption.AutoSize = True
        Me.uxAllFieldsOption.Checked = True
        Me.uxAllFieldsOption.Location = New System.Drawing.Point(60, 160)
        Me.uxAllFieldsOption.Name = "uxAllFieldsOption"
        Me.uxAllFieldsOption.Size = New System.Drawing.Size(63, 17)
        Me.uxAllFieldsOption.TabIndex = 5
        Me.uxAllFieldsOption.TabStop = True
        Me.uxAllFieldsOption.Text = "All fields"
        Me.uxAllFieldsOption.UseVisualStyleBackColor = True
        '
        'uxEnableAutoUpdate
        '
        Me.uxEnableAutoUpdate.Checked = True
        Me.uxEnableAutoUpdate.CheckState = System.Windows.Forms.CheckState.Checked
        Me.uxEnableAutoUpdate.Location = New System.Drawing.Point(40, 112)
        Me.uxEnableAutoUpdate.Name = "uxEnableAutoUpdate"
        Me.uxEnableAutoUpdate.Size = New System.Drawing.Size(284, 17)
        Me.uxEnableAutoUpdate.TabIndex = 1
        Me.uxEnableAutoUpdate.Text = "Enable field auto-updates"
        '
        'uxEnableTools
        '
        Me.uxEnableTools.Checked = True
        Me.uxEnableTools.CheckState = System.Windows.Forms.CheckState.Checked
        Me.uxEnableTools.Location = New System.Drawing.Point(21, 89)
        Me.uxEnableTools.Name = "uxEnableTools"
        Me.uxEnableTools.Size = New System.Drawing.Size(284, 17)
        Me.uxEnableTools.TabIndex = 0
        Me.uxEnableTools.Text = "Enable taxlot editing tools"
        '
        'PropertiesForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(456, 398)
        Me.Controls.Add(Me.uxAllFieldsOption)
        Me.Controls.Add(Me.uxMinimumFieldsOption)
        Me.Controls.Add(Me.uxLogo)
        Me.Controls.Add(Me.uxDescription)
        Me.Controls.Add(Me.uxEnableAutoUpdate)
        Me.Controls.Add(Me.uxEnableTools)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "PropertiesForm"
        Me.Padding = New System.Windows.Forms.Padding(8, 0, 0, 0)
        Me.Text = "PropertiesForm"
        CType(Me.uxLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region

End Class

