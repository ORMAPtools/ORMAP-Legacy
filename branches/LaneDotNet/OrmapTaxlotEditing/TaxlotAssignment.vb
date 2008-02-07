#Region "Copyright 2008 ORMAP Tech Group"

' File: TaxlotAssignment.vb

' Author: .NET Migration Team (Shad Campbell, James Moore, Nick Seigal)
' Created: January 8, 2008

' All rights reserved. Reproduction or transmission of this file, or a portion thereof,
' is forbidden without prior written permission of the ORMAP Tech Group.

#End Region

Imports System.Drawing
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework

<ComVisible(True)> _
<ComClass(TaxlotAssignment.ClassId, TaxlotAssignment.InterfaceId, TaxlotAssignment.EventsId), _
ProgId("ORMAPTaxlotEditing.TaxlotAssignment")> _
Public NotInheritable Class TaxlotAssignment
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "d091f7ea-0626-4d05-9a7c-533e0961f1cc"
    Public Const InterfaceId As String = "938bb5ab-a827-4731-be72-9db650fb8ed3"
    Public Const EventsId As String = "04e3b3ef-1fb3-46ed-b8d7-11dc19763f32"
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

#Region "Private Fields"

    Private _application As IApplication

#End Region

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' Define protected instance field values for the public properties.
        MyBase.m_category = "OrmapToolbar"  'localizable text 
        MyBase.m_caption = "TaxlotAssignment"   'localizable text 
        MyBase.m_message = "Populate values in the Taxlots feature class based on a starting value and an increment value."   'localizable text 
        MyBase.m_toolTip = "Assign Taxlots" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_TaxlotAssignment"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            ' Set the bitmap based on the name of the class.
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As ArgumentException
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

    End Sub

#End Region

#Region "Inherited Properties and Methods"

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Return MyBase.Enabled AndAlso _
                EditorExtension.Editor IsNot Nothing AndAlso _
                EditorExtension.Editor.EditState = esriEditState.esriStateEditing AndAlso _
                EditorExtension.IsValidWorkspace AndAlso _
                EditorExtension.HasValidLicense AndAlso _
                EditorExtension.CanEditTaxlots
        End Get
    End Property

    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            _application = CType(hook, IApplication)

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
        'TODO: Add TaxlotAssignment.OnClick implementation
        System.Windows.Forms.MessageBox.Show("Add TaxlotAssignment.OnClick implementation")
    End Sub

#End Region

End Class



