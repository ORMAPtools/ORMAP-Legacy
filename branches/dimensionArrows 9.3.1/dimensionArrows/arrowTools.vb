Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.esriSystem
'Imports ESRI.ArcGIS.Geometry
'Imports ESRI.ArcGIS.Geodatabase
'Imports ESRI.ArcGIS.Carto
'Imports System.Drawing
Imports ESRI.ArcGIS.Framework
'Imports ESRI.ArcGIS.ArcMapUI
'Imports System.Windows.Forms
'Imports ESRI.ArcGIS.Display
'Imports ESRI.ArcGIS.Editor
'Imports ESRI.ArcGIS.SystemUI

<ComClass(arrowTools.ClassId, arrowTools.InterfaceId, arrowTools.EventsId), _
 ProgId("dimensionArrows.dimensionTools")> _
Public NotInheritable Class arrowTools
    Inherits BaseToolbar

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
        MxCommandBars.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "0ec01c46-f345-49f1-ae8a-6bd084ddec99"
    Public Const InterfaceId As String = "8194ea23-666e-4c64-93bb-254fe25cee23"
    Public Const EventsId As String = "dfd2b5cd-13b9-4cad-ba38-677c9ff48557"
#End Region

    Private targetExtID As New UIDClass()

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()

        AddItem("dimensionArrows.straightArrows", 1)
        AddItem("dimensionArrows.landHook", 1)
        AddItem("dimensionArrows.curvedArrows", 1)
        AddItem("dimensionArrows.singleArrow", 1)
        'BeginGroup() 'Separator
    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "Dimension Arrow Tools"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "dimensionArrowTools"
        End Get
    End Property
End Class
