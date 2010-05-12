Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices

<ComClass(AssessorToolbar.ClassId, AssessorToolbar.InterfaceId, AssessorToolbar.EventsId), _
 ProgId("AssessorToolbar.AssessorToolbar")> _
Public NotInheritable Class AssessorToolbar
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
    Public Const ClassId As String = "e2453546-feb3-40df-8c69-5402735f720a"
    Public Const InterfaceId As String = "0ae1c53d-b66a-4a48-bb9e-5b173b0d011d"
    Public Const EventsId As String = "21944634-ff0d-4c2a-9cdb-f474325179f7"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()

        AddItem("AssessorToolbar.SnappingCommand")
        BeginGroup() 'separator
        AddItem("AssessorToolbar.CalculateDecimalPlaces")
        AddItem("AssessorToolbar.DrawSelectedDangles")
        AddItem("AssessorToolbar.DrawSelectedArrows")
        AddItem("AssessorToolbar.DefinitionQuery")
        AddItem("AssessorToolbar.ToggleReferenceScale")
        AddItem("AssessorToolbar.SortCancelledNumbers")
        AddItem("AssessorToolbar.DrawNeatLine")
        AddItem("AssessorToolbar.DrawSectionGraphic")
        AddItem("AssessorToolbar.FilterAnnoScale")

    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "Assessor Toolbar"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "AssessorToolbar"
        End Get
    End Property
End Class
