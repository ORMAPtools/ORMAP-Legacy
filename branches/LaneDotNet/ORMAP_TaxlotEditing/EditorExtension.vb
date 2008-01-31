Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework

Imports System.Runtime.InteropServices

<ComClass(EditorExtension.ClassId, EditorExtension.InterfaceId, EditorExtension.EventsId), _
 ProgId("ORMAPTaxlotEditing.EditorExtension")> _
Public Class EditorExtension
    Implements ESRI.ArcGIS.esriSystem.IExtension
    Implements ESRI.ArcGIS.esriSystem.IExtensionAccelerators
    Implements ESRI.ArcGIS.esriSystem.IExtensionConfig
    Implements ESRI.ArcGIS.esriSystem.IPersistVariant


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
        EditorExtensions.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        EditorExtensions.Unregister(regKey)

    End Sub

#End Region
#End Region


#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "3ffddc1a-bf54-45b4-a9dc-88740d97bcc2"
    Public Const InterfaceId As String = "cf8fd284-b76e-4012-a738-bce6e0cbbff4"
    Public Const EventsId As String = "e5719155-369f-4b3e-9e5e-99856449f05b"
#End Region

    Private m_Application As IApplication
    Private m_ExtensionState As esriExtensionState

    Private Const RequiredProductCode As esriLicenseProductCode = esriLicenseProductCode.esriLicenseProductCodeArcEditor

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
        Get
            Return "EditorExtension"
        End Get
    End Property

    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        m_Application = Nothing
    End Sub

    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        m_Application = CType(initializationData, IApplication)
    End Sub

    Public Sub CreateAccelerators() Implements ESRI.ArcGIS.esriSystem.IExtensionAccelerators.CreateAccelerators
        'TODO: Implement this
    End Sub

    Public ReadOnly Property Description() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.Description
        Get
            Return "ORMAP Taxlot Editor Extension (.NET) Developed by the ORMAP technical group."
        End Get
    End Property

    Public ReadOnly Property ProductName() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.ProductName
        Get
            Return "ORMAP Taxlot Editor Extension"
        End Get
    End Property

    Public Property State() As ESRI.ArcGIS.esriSystem.esriExtensionState Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.State
        Get
            Return m_ExtensionState
        End Get
        Set(ByVal value As ESRI.ArcGIS.esriSystem.esriExtensionState)

            If value = m_ExtensionState Then
                '[New setting and current value are the same...]
                Exit Property
            End If

            'Check if OK to enable or disable extension
            'Determine if state can be changed
            Dim ValidatedExtensionState As esriExtensionState = ValidateState(value)
            If ValidatedExtensionState = esriExtensionState.esriESUnavailable Then
                '[Cannot enable if it's already in unavailable state...]
                Throw New COMException("Not running the appropriate product license.")
            Else
                m_ExtensionState = ValidatedExtensionState
            End If

        End Set
    End Property

    Public ReadOnly Property ID() As ESRI.ArcGIS.esriSystem.UID Implements ESRI.ArcGIS.esriSystem.IPersistVariant.ID
        Get
            Dim pUID As ESRI.ArcGIS.esriSystem.UID
            pUID = "ORMAPTaxlotEditing.EditorExtension"
            ID = pUID
        End Get
    End Property

    Public Sub Load(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Load
        'TODO: Implement this
    End Sub

    Public Sub Save(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Save
        'TODO: Implement this
    End Sub

    Private Function ValidateState(ByVal desiredState As esriExtensionState) As esriExtensionState
        'Validate the desired extension state based on whether the correct license is running

        Dim AoInitTestProduct As IAoInitialize = New AoInitializeClass()
        Dim ProductCode As esriLicenseProductCode = AoInitTestProduct.InitializedProduct()

        If ProductCode >= RequiredProductCode Then
            Return desiredState
        Else
            Return esriExtensionState.esriESUnavailable
        End If

    End Function

End Class

