#Region "Copyright 2008 ORMAP Tech Group"

' File: CombineTaxlots.vb

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
<ComClass(CombineTaxlots.ClassId, CombineTaxlots.InterfaceId, CombineTaxlots.EventsId), _
ProgId("ORMAPTaxlotEditing.CombineTaxlots")> _
Public NotInheritable Class CombineTaxlots
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "cd9a4052-836a-42a1-bd09-b3bee8dee1a9"
    Public Const InterfaceId As String = "3429b024-92e6-4e59-b2fb-720efce1a98c"
    Public Const EventsId As String = "c612e825-3371-407d-8fb9-f478981a22dc"
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

        ' Define protected instance field values for the public properties
        MyBase.m_category = "OrmapToolbar"  'localizable text 
        MyBase.m_caption = "CombineSelectedTaxlots"   'localizable text 
        MyBase.m_message = "Combine taxlots in the Taxlot polygon feature class."   'localizable text 
        MyBase.m_toolTip = "Combine Selected Taxlots" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_CombineTaxlots"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add CombineTaxlots.OnClick implementation
        System.Windows.Forms.MessageBox.Show("Add CombineTaxlots.OnClick implementation")
    End Sub

#End Region

End Class



