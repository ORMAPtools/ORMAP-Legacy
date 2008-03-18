#Region "Copyright 2008 ORMAP Tech Group"

' File:  LocateFeature.vb
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
'SCC revision number: $Revision$
'Date of Last Change: $Date$
#End Region

Imports System.Drawing
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework


<ComVisible(True)> _
<ComClass(LocateFeature.ClassId, LocateFeature.InterfaceId, LocateFeature.EventsId), _
ProgId("ORMAPTaxlotEditing.LocateFeature")> _
Public NotInheritable Class LocateFeature
    Inherits BaseCommand

#Region "Class-Level Constants And Enumerations"
    ' None
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
        MyBase.m_caption = "LocateFeature"   'localizable text 
        MyBase.m_message = "Locate a Taxlot or Mapindex"   'localizable text 
        MyBase.m_toolTip = "Locate Taxlot or Mapindex" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_LocateFeature"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            ' Set the bitmap based on the name of the class.
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As ArgumentException
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication

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

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Dim canEnable As Boolean
            canEnable = EditorExtension.CanEnableExtendedEditing
            Return canEnable
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

        ' TODO: Add other initialization code
    End Sub

    Public Overrides Sub OnClick()
        ' TODO: Port LocateFeature.OnClick implementation
        System.Windows.Forms.MessageBox.Show("Port LocateFeature.OnClick implementation")
    End Sub

#End Region

#End Region

#Region "Implemented Interface Members"
    ' None
#End Region

#Region "Other Members"

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "ecace054-c346-4be5-a980-8d81ede68fcb"
    Public Const InterfaceId As String = "effc83d3-a320-460c-9b1e-7f48a4b282d8"
    Public Const EventsId As String = "591a6993-fcb1-4c0c-879f-5488e0c88312"
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



