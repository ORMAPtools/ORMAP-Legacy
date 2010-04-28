#Region "Copyright 2008 ORMAP Tech Group"

' File:  CombineTaxlots.vb
'
' Original Author:  OPET.NET Migration Team (Shad Campbell, James Moore, 
'                   Nick Seigal)
'
' Date Created:  January 8, 2008
'
' Copyright Holder:  ORMAP Tech Group  
' Contact Info:  ORMAP Tech Group may be reached at 
' ORMAP_ESRI_Programmers@listsmart.osl.state.or.us
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
'SCC revision number: $Revision: 406 $
'Date of Last Change: $Date: 2009-11-30 22:49:20 -0800 (Mon, 30 Nov 2009) $
#End Region

#Region "Imported NameSpaces"
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Environment
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports OrmapTaxlotEditing.DataMonitor
Imports OrmapTaxlotEditing.LocateFeatureDockWin
#End Region

<ComClass(MapDefinition.ClassId, MapDefinition.InterfaceId, MapDefinition.EventsId), _
 ProgId("OrmapTaxlotEditing.MapDefinition")> _
Public NotInheritable Class MapDefinition
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4fb41cb8-b7ca-4a5e-8722-be5267223dfe"
    Public Const InterfaceId As String = "915ba534-4d27-49e4-aca5-2f5c7368566b"
    Public Const EventsId As String = "26a1be93-0f53-4dbf-a41f-2bf860170df3"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region


    Private _application As IApplication

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "OrmapToolbar"  'localizable text 
        MyBase.m_caption = "Map Definition Query"   'localizable text 
        MyBase.m_message = "Sets a Defintion for the Map currently selected."   'localizable text 
        MyBase.m_toolTip = "Sets a Defintion for the Map currently selected." 'localizable text 
        MyBase.m_name = MyBase.m_category & "_MapDefinitionQuery"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        'Try
        '    'TODO: change bitmap name if necessary
        '    Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
        '    MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        'Catch ex As Exception
        '    EditorExtension.ProcessUnhandledException(ex)
        'End Try


    End Sub
#Region "Properties"

    Private WithEvents _partnerMapDefinitionForm As MapDefinitionForm
    Private _theMapNumber As String
    Private _ApplyQuery As Boolean = False

    Friend ReadOnly Property PartnerMapDefinitionForm() As MapDefinitionForm
        Get
            If _partnerMapDefinitionForm Is Nothing OrElse _partnerMapDefinitionForm.IsDisposed Then
                setPartnerMapDefinitionForm(New MapDefinitionForm())
            End If
            Return _partnerMapDefinitionForm
        End Get
    End Property

    Private Sub setPartnerMapDefinitionForm(ByVal value As MapDefinitionForm)

        If value IsNot Nothing Then
            _partnerMapDefinitionForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerMapDefinitionForm.Load, AddressOf PartnerMapDefinitionForm_Load
            AddHandler _partnerMapDefinitionForm.uxCancelSetDefinitionQuery.Click, AddressOf CancelSetDefinitionQuery_Click
            AddHandler _partnerMapDefinitionForm.uxSetMapDefinitionQuery.Click, AddressOf SetMapDefinitionQuery_Click
        Else
            ' Unsubscribe to partner form events.
            RemoveHandler _partnerMapDefinitionForm.Load, AddressOf PartnerMapDefinitionForm_Load
            RemoveHandler _partnerMapDefinitionForm.uxCancelSetDefinitionQuery.Click, AddressOf CancelSetDefinitionQuery_Click
            RemoveHandler _partnerMapDefinitionForm.uxSetMapDefinitionQuery.Click, AddressOf SetMapDefinitionQuery_Click
        End If

    End Sub

#End Region

#Region "Event Handlers"

    Private Sub PartnerMapDefinitionForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles _ParnterMapDefinitionForm.Load

        _partnerMapDefinitionForm.uxDefinitonQueryTextBox.Text = _theMapNumber

    End Sub
    Private Sub CancelSetDefinitionQuery_Click(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles _partnerMapDefinitionForm.uxCancelSetDefinitionQuery.Click

        _ApplyQuery = False
        _partnerMapDefinitionForm.Dispose()

    End Sub
    Private Sub SetMapDefinitionQuery_Click(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles _partnerMapDefinitionForm.uxSetMapDefinitionQuery.Click

        _theMapNumber = _partnerMapDefinitionForm.uxDefinitonQueryTextBox.Text
        _ApplyQuery = True
        _partnerMapDefinitionForm.Dispose()

    End Sub
#End Region

#Region "Properties"
    Public ReadOnly Property MapNumber() As String 'Returns the MapNumber to Process
        Get
            Return _theMapNumber
        End Get
    End Property
    Public ReadOnly Property ApplyQuery() As Boolean 'Returns a boolean value
        Get
            Return _ApplyQuery
        End Get
    End Property

#End Region
#Region "Methods"
    
#End Region

    Public Sub CallForm(ByVal hook As Object, ByVal theMapNumber As String)
        If Not hook Is Nothing Then

            'Disable tool if parent application is not ArcMap
            If TypeOf hook Is IMxApplication Then
                _application = DirectCast(hook, IApplication)
                _theMapNumber = theMapNumber
                setPartnerMapDefinitionForm(New MapDefinitionForm())
                _partnerMapDefinitionForm.ShowDialog()
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If
    End Sub

    Public Overrides Sub onCreate(ByVal hook As Object)
        If Not hook Is Nothing Then

            'Disable tool if parent application is not ArcMap
            If TypeOf hook Is IMxApplication Then
                _application = DirectCast(hook, IApplication)
                setPartnerMapDefinitionForm(New MapDefinitionForm())
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If
    End Sub

End Class



