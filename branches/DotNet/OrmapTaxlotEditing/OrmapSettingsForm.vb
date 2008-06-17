#Region "Copyright 2008 ORMAP Tech Group"

' File:  OrmapSettingsForm.vb
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
' modify it under the terms of the Lesser GNU General Public License as 
' published by the Free Software Foundation; either version 3 of the License, 
' or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT 
' ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or 
' FITNESS FOR A PARTICULAR PURPOSE.  See the Lesser GNU General Public License 
' located in the COPYING.LESSER.txt file for more details.
'
' You should have received a copy of the Lesser GNU General Public License 
' along with the ORMAP Taxlot Editing Toolbar; if not, write to the Free 
' Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 
' 02110-1301 USA.

#End Region

#Region "Subversion Keyword Expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision$
'Date of Last Change: $Date$
#End Region

#Region "Imported Namespaces"
Imports System.Runtime.InteropServices
Imports System.Configuration
Imports System.Windows.Forms
#End Region

<ComVisible(False)> _
Public Class OrmapSettingsForm

#Region "Class-Level Constants and Enumerations (none)"
#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        SetBindings()

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields (none)"
#End Region

#Region "Properties (none)"
#End Region

#Region "Event Handlers"

    ''' <summary>
    ''' Closes the form without saving any modified settings.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub uxCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxCancel.Click

        ReloadSettings()

        If Me.Modal Then
            ' Modal form is closed automatically by the 
            ' uxCancel.DialogResult = Cancel property. 
        Else
            Me.Close()
        End If

    End Sub

    ''' <summary>
    ''' Reloads application settings values and keeps the dialog open. 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub uxReload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxReload.Click

        ReloadSettings()

    End Sub

    ''' <summary>
    ''' Resets application settings values and keeps the dialog open. 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub uxReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxReset.Click

        ResetSettings()

    End Sub

    ''' <summary>
    ''' Saves the current application settings values and closes the dialog.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub uxSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxSave.Click

        SaveSettings()

        If Me.Modal Then
            ' Modal form is closed automatically by the 
            ' uxSave.DialogResult = Cancel property. 
        Else
            Me.Close()
        End If

    End Sub

#End Region

#Region "Methods"

    ''' <summary>
    ''' Set the control binding sources for the form.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetBindings()

        ' Set the control binding sources for all the controls on the settings tabs.
        TableNamesSettingsBindingSource.DataSource = EditorExtension.TableNamesSettings
        AnnoTableNamesSettingsBindingSource.DataSource = EditorExtension.AnnoTableNamesSettings
        AllTablesSettingsBindingSource.DataSource = EditorExtension.AllTablesSettings
        MapIndexSettingsBindingSource.DataSource = EditorExtension.MapIndexSettings
        TaxLotSettingsBindingSource.DataSource = EditorExtension.TaxLotSettings
        TaxLotLinesSettingsBindingSource.DataSource = EditorExtension.TaxLotLinesSettings
        CartographicLinesSettingsBindingSource.DataSource = EditorExtension.CartographicLinesSettings
        TaxlotAcreageAnnoSettingsBindingSource.DataSource = EditorExtension.TaxlotAcreageAnnoSettings
        TaxlotNumberAnnoSettingsBindingSource.DataSource = EditorExtension.TaxlotNumberAnnoSettings
        DefaultValuesSettingsBindingSource.DataSource = EditorExtension.DefaultValuesSettings

    End Sub

    ''' <summary>
    ''' Stores the current values of the application settings in persistent storage.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SaveSettings()

        Dim settings As ApplicationSettingsBase
        settings = DirectCast(TableNamesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Save()
        settings = DirectCast(AnnoTableNamesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Save()
        settings = DirectCast(AllTablesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Save()
        settings = DirectCast(MapIndexSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Save()
        settings = DirectCast(TaxLotSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Save()
        settings = DirectCast(TaxLotLinesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Save()
        settings = DirectCast(CartographicLinesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Save()
        settings = DirectCast(TaxlotAcreageAnnoSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Save()
        settings = DirectCast(TaxlotNumberAnnoSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Save()
        settings = DirectCast(DefaultValuesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Save()

    End Sub

    ''' <summary>
    ''' Refreshes the application settings values from persistent 
    ''' storage.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReloadSettings()

        Dim settings As ApplicationSettingsBase
        settings = DirectCast(TableNamesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reload()
        settings = DirectCast(AnnoTableNamesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reload()
        settings = DirectCast(AllTablesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reload()
        settings = DirectCast(MapIndexSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reload()
        settings = DirectCast(TaxLotSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reload()
        settings = DirectCast(TaxLotLinesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reload()
        settings = DirectCast(CartographicLinesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reload()
        settings = DirectCast(TaxlotAcreageAnnoSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reload()
        settings = DirectCast(TaxlotNumberAnnoSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reload()
        settings = DirectCast(DefaultValuesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reload()

    End Sub

    ''' <summary>
    ''' Restores the persisted application settings values to their 
    ''' corresponding default properties.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ResetSettings()

        Dim settings As ApplicationSettingsBase
        settings = DirectCast(TableNamesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reset()
        settings = DirectCast(AnnoTableNamesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reset()
        settings = DirectCast(AllTablesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reset()
        settings = DirectCast(MapIndexSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reset()
        settings = DirectCast(TaxLotSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reset()
        settings = DirectCast(TaxLotLinesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reset()
        settings = DirectCast(CartographicLinesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reset()
        settings = DirectCast(TaxlotAcreageAnnoSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reset()
        settings = DirectCast(TaxlotNumberAnnoSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reset()
        settings = DirectCast(DefaultValuesSettingsBindingSource.DataSource, ApplicationSettingsBase)
        settings.Reset()

    End Sub

#End Region

#End Region

#Region "Inherited Class Members (none)"

#Region "Properties (none)"
#End Region

#Region "Methods (none)"
#End Region

#End Region

#Region "Implemented Interface Members (none) "
#End Region

#Region "Other Members (none)"
#End Region

    Private Sub TableLayoutPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub
End Class