#Region "Copyright 2008 ORMAP Tech Group"

' File:  PropertiesForm.vb
'
' Original Author:  OPET.NET Migration Team (Shad Campbell, James Moore, 
'                   Nick Seigal)
'
' Date Created:  January 8, 2008
'
' Copyright Holder:  ORMAP Tech Group  
' Contact Info:  ORMAP Tech Group (a.k.a. opet developers) may be reached at 
' opet-developers@lists.sourceforge.net
'
' This file is part of the ORMAP Taxlot Editing Toolbar.
'
' ORMAP Taxlot Editing Toolbar is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License as published by 
' the Free Software Foundation; either version 3 of the License, or (at your 
' option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT 
' ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or 
' FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License located
' in the COPYING.txt file for more details.
'
' You should have received a copy of the GNU General Public License along
' with the ORMAP Taxlot Editing Toolbar; if not, write to the Free Software 
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.

#End Region

Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

<ComVisible(False)> _
Friend NotInheritable Class PropertiesForm : Inherits System.Windows.Forms.Form


#Region "Class-Level Constants And Enumerations"
    ' None
#End Region

#Region "Built-In Class Members (Properties, Methods, Events, Event Handlers, Delegates, Etc.)"

#Region "Constructors"

    Public Sub New()
        '
        ' Required for Windows Form Designer support
        '
        InitializeComponent()

        ' TODO: Add any constructor code after InitializeComponent call

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    ''' <summary>
    ''' Required designer variable.
    ''' </summary>
    Private components As System.ComponentModel.Container = Nothing

    Friend WithEvents uxEnableTools As CheckBox
    Friend WithEvents uxEnableAutoUpdate As CheckBox
    Friend WithEvents uxDescription As Label
    Friend WithEvents uxLogo As PictureBox
    Friend WithEvents uxMinimumFieldsOption As RadioButton
    Friend WithEvents uxAllFieldsOption As RadioButton
    Friend WithEvents uxSettings As Button

#End Region

#Region "Properties"
    ' None
#End Region

#Region "Event Handlers"
    ' None
#End Region

#Region "Methods"
    ' None
#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"
    ' None
#End Region

#Region "Methods"

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

#End Region

#Region "Implemented Interface Members"
    ' None
#End Region

#Region "Other Members"

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
        Me.uxSettings = New System.Windows.Forms.Button
        CType(Me.uxLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'uxDescription
        '
        Me.uxDescription.Location = New System.Drawing.Point(11, 13)
        Me.uxDescription.Name = "uxDescription"
        Me.uxDescription.Size = New System.Drawing.Size(370, 34)
        Me.uxDescription.TabIndex = 2
        Me.uxDescription.Text = "The ORMAP Taxlot Editing Toolbar gives the user tools to edit taxlots and related" & _
            " features in compliance with the ORMAP standard."
        '
        'uxLogo
        '
        Me.uxLogo.Image = CType(resources.GetObject("uxLogo.Image"), System.Drawing.Image)
        Me.uxLogo.Location = New System.Drawing.Point(14, 302)
        Me.uxLogo.Name = "uxLogo"
        Me.uxLogo.Size = New System.Drawing.Size(86, 82)
        Me.uxLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.uxLogo.TabIndex = 3
        Me.uxLogo.TabStop = False
        '
        'uxMinimumFieldsOption
        '
        Me.uxMinimumFieldsOption.AutoSize = True
        Me.uxMinimumFieldsOption.Location = New System.Drawing.Point(53, 97)
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
        Me.uxAllFieldsOption.Location = New System.Drawing.Point(53, 121)
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
        Me.uxEnableAutoUpdate.Location = New System.Drawing.Point(33, 73)
        Me.uxEnableAutoUpdate.Name = "uxEnableAutoUpdate"
        Me.uxEnableAutoUpdate.Size = New System.Drawing.Size(284, 17)
        Me.uxEnableAutoUpdate.TabIndex = 1
        Me.uxEnableAutoUpdate.Text = "Enable field auto-updates"
        '
        'uxEnableTools
        '
        Me.uxEnableTools.Checked = True
        Me.uxEnableTools.CheckState = System.Windows.Forms.CheckState.Checked
        Me.uxEnableTools.Location = New System.Drawing.Point(14, 50)
        Me.uxEnableTools.Name = "uxEnableTools"
        Me.uxEnableTools.Size = New System.Drawing.Size(284, 17)
        Me.uxEnableTools.TabIndex = 0
        Me.uxEnableTools.Text = "Enable taxlot editing tools"
        '
        'uxSettings
        '
        Me.uxSettings.Location = New System.Drawing.Point(14, 164)
        Me.uxSettings.Name = "uxSettings"
        Me.uxSettings.Size = New System.Drawing.Size(86, 23)
        Me.uxSettings.TabIndex = 6
        Me.uxSettings.Text = "Settings..."
        Me.uxSettings.UseVisualStyleBackColor = True
        '
        'PropertiesForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(456, 398)
        Me.Controls.Add(Me.uxSettings)
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

#End Region

End Class

