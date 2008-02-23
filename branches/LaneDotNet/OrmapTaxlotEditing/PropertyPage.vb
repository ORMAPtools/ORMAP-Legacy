#Region "Copyright 2008 ORMAP Tech Group"

' File:  PropertyPage.vb
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
#Region "Subversion Keyword expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision:$
'Date of Last Change: $Date:$
#End Region
Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase

<ComVisible(True)> _
<ComClass(PropertyPage.ClassId, PropertyPage.InterfaceId, PropertyPage.EventsId), _
ProgId("ORMAPTaxlotEditing.PropertyPage")> _
Public NotInheritable Class PropertyPage
    Implements IComPropertyPage

#Region "Class-Level Constants And Enumerations"
    ' None
#End Region

#Region "Built-In Class Members (Properties, Methods, Events, Event Handlers, Delegates, Etc.)"

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"
    ' None
#End Region

#Region "Properties"

    Private _pageDirty As Boolean '= False

    Friend ReadOnly Property PageDirty() As Boolean
        Get
            Return _pageDirty
        End Get
    End Property

    Private Sub SetPageDirty(ByVal value As Boolean)
        ' TODO: Add validation code?
        _pageDirty = value
    End Sub

    Private _propertiesPageSite As IComPropertyPageSite

    Friend ReadOnly Property PropertiesPageSite() As IComPropertyPageSite
        Get
            Return _propertiesPageSite
        End Get
    End Property

    Private Sub SetPropertiesPageSite(ByVal value As IComPropertyPageSite)
        ' TODO: Add validation code?
        _propertiesPageSite = value
    End Sub

    Private WithEvents _propertiesForm As PropertiesForm  ' TODO: Is WithEvents needed here?

    Friend ReadOnly Property PropertiesForm() As PropertiesForm
        Get
            Return _propertiesForm
        End Get
    End Property

    Private Sub SetPropertiesForm(ByVal value As PropertiesForm)
        ' TODO: Add validation code?
        _propertiesForm = value
    End Sub

#End Region

#Region "Event Handlers"

    Private Sub uxEnableTools_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)

        PropertiesForm.uxEnableAutoUpdate.Enabled = PropertiesForm.uxEnableTools.Checked
        PropertiesForm.uxMinimumFieldsOption.Enabled = PropertiesForm.uxEnableTools.Checked
        PropertiesForm.uxAllFieldsOption.Enabled = PropertiesForm.uxEnableTools.Checked

        ' Set dirty flag.
        SetPageDirty(True)

        If Not PropertiesPageSite Is Nothing Then
            PropertiesPageSite.PageChanged()
        End If

    End Sub

    Private Sub uxEnableAutoUpdate_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)

        PropertiesForm.uxMinimumFieldsOption.Enabled = PropertiesForm.uxEnableAutoUpdate.Checked
        PropertiesForm.uxAllFieldsOption.Enabled = PropertiesForm.uxEnableAutoUpdate.Checked

        ' Set dirty flag.
        SetPageDirty(True)

        If Not PropertiesPageSite Is Nothing Then
            PropertiesPageSite.PageChanged()
        End If

    End Sub

    Private Sub uxSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim settingsForm As New OrmapSettingsForm
        settingsForm.ShowDialog(DirectCast(sender, Control).FindForm)

    End Sub

    Private Sub uxAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim aboutForm As New AboutForm
        aboutForm.ShowDialog(DirectCast(sender, Control).FindForm)

    End Sub

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
    ' None
#End Region

#End Region

#Region "Implemented Interface Members"

#Region "IComPropertyPage Implementations"

    Public ReadOnly Property Height() As Integer Implements IComPropertyPage.Height
        Get
            Return PropertiesForm.Height
        End Get
    End Property

    Public ReadOnly Property HelpFile() As String Implements IComPropertyPage.HelpFile
        Get
            Return Nothing  ' TODO: Implement Help File
        End Get
    End Property

    Public ReadOnly Property HelpContextID(ByVal controlID As Integer) As Integer Implements IComPropertyPage.HelpContextID
        Get
            Return 0  ' TODO: Implement Help File
        End Get
    End Property

    Public ReadOnly Property IsPageDirty() As Boolean Implements IComPropertyPage.IsPageDirty
        Get
            Return PageDirty
        End Get
    End Property

    Public WriteOnly Property PageSite() As ESRI.ArcGIS.Framework.IComPropertyPageSite Implements IComPropertyPage.PageSite
        Set(ByVal value As ESRI.ArcGIS.Framework.IComPropertyPageSite)
            SetPropertiesPageSite(value)
        End Set
    End Property

    Public Property Priority() As Integer Implements IComPropertyPage.Priority
        Get
            Return 0  'Lowest number = last/rightmost tab position in the Properties window.
        End Get
        Set(ByVal value As Integer)
            ' Do not set anything
        End Set
    End Property

    Public Property Title() As String Implements IComPropertyPage.Title
        Get
            Return "ORMAP Taxlot Editor"
        End Get
        Set(ByVal value As String)
            ' Do not set anything
        End Set
    End Property

    Public ReadOnly Property Width() As Integer Implements IComPropertyPage.Width
        Get
            Return PropertiesForm.Width
        End Get
    End Property

    Public Function Activate() As Integer Implements IComPropertyPage.Activate
        Return PropertiesForm.Handle.ToInt32()
    End Function

    Public Function Applies(ByVal objects As ESRI.ArcGIS.esriSystem.ISet) As Boolean Implements IComPropertyPage.Applies

        ' Do not affirm if the objects list is empty.
        If objects Is Nothing OrElse objects.Count = 0 Then
            Return False
        End If
        objects.Reset()

        ' Get a reference to the editor.
        ' Do not affirm if the editor is not found.
        Dim editor As IEditor = TryCast(objects.Next(), IEditor)
        If editor Is Nothing Then
            Return False
        End If

        ' Do not affirm if the user is not editing.
        If editor.EditState <> esriEditState.esriStateEditing Then
            Return False
        End If

        ' Do not affirm if the user is editing a file-based workspace (e.g. coverages, shapefiles).
        If editor.EditWorkspace.Type = esriWorkspaceType.esriFileSystemWorkspace Then
            Return False
        End If

        ' Otherwise, affirm.
        Return True

    End Function

    Public Sub Apply() Implements IComPropertyPage.Apply
        ' Write to the EditorExtension.CanEdit shared (i.e. by all class objects) property
        EditorExtension.CanEditTaxlots = PropertiesForm.uxEnableTools.Checked
        EditorExtension.CanAutoUpdate = PropertiesForm.uxEnableAutoUpdate.Checked
        EditorExtension.CanAutoUpdateAllFields = Not PropertiesForm.uxAllFieldsOption.Checked
        SetPageDirty(False)
    End Sub

    Public Sub Cancel() Implements IComPropertyPage.Cancel
        ' TODO: Implement this?
    End Sub

    Public Sub Deactivate() Implements IComPropertyPage.Deactivate
        If Not _propertiesForm Is Nothing Then
            PropertiesForm.Dispose()
        End If
        SetPropertiesForm(Nothing)
        SetPropertiesPageSite(Nothing)
    End Sub

    Public Sub Hide() Implements IComPropertyPage.Hide
        PropertiesForm.Hide()
    End Sub

    Public Sub SetObjects(ByVal objects As ESRI.ArcGIS.esriSystem.ISet) Implements IComPropertyPage.SetObjects
        ' Note: The Applies() method should have done preliminary checking of 
        ' editor states before this method is called.

        ' TODO: Move (to where)?
        SetPropertiesForm(New PropertiesForm())
        PropertiesForm.uxEnableTools.Checked = EditorExtension.CanEditTaxlots
        PropertiesForm.uxEnableAutoUpdate.Checked = EditorExtension.CanAutoUpdate
        PropertiesForm.uxMinimumFieldsOption.Checked = Not EditorExtension.CanAutoUpdateAllFields
        PropertiesForm.uxAllFieldsOption.Checked = EditorExtension.CanAutoUpdateAllFields

        ' Wire up form events.
        AddHandler PropertiesForm.uxEnableTools.CheckedChanged, AddressOf uxEnableTools_CheckedChanged
        AddHandler PropertiesForm.uxEnableAutoUpdate.CheckedChanged, AddressOf uxEnableAutoUpdate_CheckedChanged
        AddHandler PropertiesForm.uxSettings.Click, AddressOf uxSettings_Click
        AddHandler PropertiesForm.uxAbout.Click, AddressOf uxAbout_Click

    End Sub

    Public Sub Show() Implements IComPropertyPage.Show
        PropertiesForm.Show()
    End Sub

#End Region

#End Region

#Region "Other Members"

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "050c23da-ebd8-4a1d-871b-b7a9354d331b"
    Public Const InterfaceId As String = "bae36023-8a03-43b6-bea6-fab534ff7c5e"
    Public Const EventsId As String = "8ab94224-407b-4139-a003-48f5789bf3b3"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisible(False)> _
    Private Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        '
        ' TODO: Add any COM registration code here
        '
    End Sub

    <ComUnregisterFunction(), ComVisible(False)> _
    Private Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        '
        ' TODO: Add any COM unregistration code here
        '
    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        EditorPropertyPages.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        EditorPropertyPages.Unregister(regKey)

    End Sub

#End Region
#End Region

#End Region

End Class

