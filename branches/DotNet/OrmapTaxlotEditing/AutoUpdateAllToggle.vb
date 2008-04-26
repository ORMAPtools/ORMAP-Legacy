#Region "Copyright 2008 ORMAP Tech Group"

' File:  AutoUpdateAllToggle.vb
'
' Original Author:  OPET.NET Migration Team (Shad Campbell, James Moore, 
'                   Nick Seigal)
'
' Date Created:  March 30, 2008
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

#Region "Subversion Keyword Expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision: 239 $
'Date of Last Change: $Date: 2008-03-18 02:11:11 -0700 (Tue, 18 Mar 2008) $
#End Region

#Region "Imported Namespaces"
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework
#End Region

<ComVisible(True)> _
<ComClass(AutoUpdateAllToggle.ClassId, AutoUpdateAllToggle.InterfaceId, AutoUpdateAllToggle.EventsId), _
ProgId("ORMAPTaxlotEditing.AutoUpdateAllToggle")> _
Public NotInheritable Class AutoUpdateAllToggle
    Inherits BaseCommand

#Region "Class-Level Constants And Enumerations (none)"
#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' Define protected instance field values for the public properties
        MyBase.m_category = "OrmapToolbar"  'localizable text 
        MyBase.m_caption = "AutoUpdateAllToggle"   'localizable text 
        MyBase.m_message = "Turn on automatic update of all taxlot fields"   'localizable text 
        MyBase.m_toolTip = "Automatic Update" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_AutoUpdateAllToggle"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")
        MyBase.m_checked = False

        Try
            ' Set the bitmap based on the name of the class.
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As ArgumentException
            Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication

#End Region

#Region "Properties (none)"
#End Region

#Region "Event Handlers (none)"
#End Region

#Region "Methods (none)"
#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            If EditorExtension.AllowedToAutoUpdateAllFields Then
                m_checked = True
            Else
                m_checked = False
            End If
            Dim canEnable As Boolean
            canEnable = EditorExtension.CanEnableExtendedEditing
            Return canEnable
        End Get
    End Property

    Public Overrides ReadOnly Property Checked() As Boolean
        Get
            Return MyBase.Checked
        End Get
    End Property

#End Region

#Region "Methods"

    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            _application = DirectCast(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If

        ' NOTE: Add other initialization code here...

    End Sub

    Public Overrides Sub OnClick()
        ' Toggle the checked state of the button.
        MyBase.m_checked = Not MyBase.m_checked
        ' Synch up the extension-level flag for auto updates.
        EditorExtension.AllowedToAutoUpdateAllFields = MyBase.m_checked
        If EditorExtension.AllowedToAutoUpdateAllFields Then
            MessageBox.Show("Auto update of taxlot fields is ON. The minimum fields" & vbNewLine & _
                    " (e.g. autodate, autowho) will be updated, as well as all taxlot fields.", _
                    "Auto Update All Toggle", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Auto update of taxlot fields is OFF. Only the minimum fields" & vbNewLine & _
                    "(e.g. autodate, autowho) will be updated.", _
                    "Auto Update All Toggle", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

#End Region

#End Region

#Region "Implemented Interface Members (none)"
#End Region

#Region "Other Members"

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "6a793dd3-2e75-4b39-a3a3-a261ebeee4e9"
    Public Const InterfaceId As String = "12aedc31-e94b-4015-ae1c-ddb004f7f953"
    Public Const EventsId As String = "415ab563-e444-4c50-b54a-9bfa92fdcac9"
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
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub

    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

#End Region

End Class


