Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Framework

<ComClass(dimensionArrowExtension.ClassId, dimensionArrowExtension.InterfaceId, dimensionArrowExtension.EventsId), _
 ProgId("dimensionArrows.dimensionArrowExtension")> _
Public Class dimensionArrowExtension
    Implements IExtension
    Implements IPersistVariant

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
        MxExtensionJIT.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxExtensionJIT.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "a8e825ce-e1e1-4d90-9afb-2ce369d6cc33"
    Public Const InterfaceId As String = "5baa7e83-ec76-45c9-a549-5c9d934efa7e"
    Public Const EventsId As String = "722069c0-a85d-4f06-93b8-d7719322a8fc"
#End Region

    Private m_application As IApplication
    Private m_enableState As esriExtensionState

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    ''' <summary>
    ''' Determine extension state
    ''' </summary>
    Private Function StateCheck(ByVal requestEnable As Boolean) As esriExtensionState
        'TODO: Replace with advanced extension state checking if needed
        'Turn on or off extension directly 
        If requestEnable Then
            Return esriExtensionState.esriESEnabled
        Else
            Return esriExtensionState.esriESDisabled
        End If
    End Function

#Region "IExtension Members"
    ''' <summary>
    ''' Name of extension. Do not exceed 31 characters
    ''' </summary>
    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
        Get
            'TODO: Modify string to uniquely identify extension
            Return "dimensionArrowExtension"
        End Get
    End Property

    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        'TODO: Clean up resources

        m_application = Nothing
    End Sub

    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        m_application = CType(initializationData, IApplication)
        If m_application Is Nothing Then Return

        'TODO: Add code to initialize the extension
    End Sub
#End Region

#Region "IPersistVariant Members"
    Public ReadOnly Property ID() As ESRI.ArcGIS.esriSystem.UID Implements ESRI.ArcGIS.esriSystem.IPersistVariant.ID
        Get
            Dim typeID As UID = New UIDClass()
            typeID.Value = Me.GetType().GUID.ToString("B")
            Return typeID
        End Get
    End Property

    Public Sub Load(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Load
        'TODO: Load persisted data from document stream
        _arrowheadIsSwitched = Stream.Read()
        _flipArrows = Stream.Read()
        Marshal.ReleaseComObject(Stream)
    End Sub

    Public Sub Save(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Save
        'TODO: Save extension related data to document stream
        Stream.Write(_arrowheadIsSwitched)
        Stream.Write(_flipArrows)
        Marshal.ReleaseComObject(Stream)
    End Sub
#End Region

End Class


