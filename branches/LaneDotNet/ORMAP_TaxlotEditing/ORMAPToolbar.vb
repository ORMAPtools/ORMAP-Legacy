Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.esriSystem





<ComClass(ORMAPToolbar.ClassId, ORMAPToolbar.InterfaceId, ORMAPToolbar.EventsId), _
 ProgId("ORMAP_TaxlotEditing.ORMAPToolbar")> _
Public NotInheritable Class ORMAPToolbar
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
    Public Const ClassId As String = "07b32247-abf9-4643-b078-4e68be56599e"
    Public Const InterfaceId As String = "98eff0b1-3e68-4881-b5f8-5f8a9e242682"
    Public Const EventsId As String = "8f2ffc36-8a61-44ef-9127-8e4935dcc563"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()



        '
        'TODO: Define your toolbar here by adding items
        '
        AddItem("ORMAPTaxlotEditing.TaxlotAssignment")
        AddItem("ORMAPTaxlotEditing.LocateFeature")
        BeginGroup() 'Separator
        AddItem("ORMAPTaxlotEditing.EditMapIndex")
        BeginGroup() 'Separator
        AddItem("ORMAPTaxlotEditing.CombineTaxlots")
        BeginGroup() 'Separator
        AddItem("ORMAPTaxlotEditing.AddArrows")


        'AddItem("{FBF8C3FB-0480-11D2-8D21-080009EE4E51}", 1) 'undo command
        'AddItem(New Guid("FBF8C3FB-0480-11D2-8D21-080009EE4E51"), 2) 'redo command


    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "ORMAP Taxlot Editor (.NET)"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "ORMAPToolbar"
        End Get
    End Property
End Class
