Imports System
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework

<ComVisible(True)> _
<ComClass(EditorExtension.ClassId, EditorExtension.InterfaceId, EditorExtension.EventsId), _
ProgId("ORMAPTaxlotEditing.EditorExtension")> _
Public NotInheritable Class EditorExtension
    Implements ESRI.ArcGIS.esriSystem.IExtension
    Implements ESRI.ArcGIS.esriSystem.IExtensionAccelerators
    Implements ESRI.ArcGIS.esriSystem.IPersistVariant

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "3ffddc1a-bf54-45b4-a9dc-88740d97bcc2"
    Public Const InterfaceId As String = "cf8fd284-b76e-4012-a738-bce6e0cbbff4"
    Public Const EventsId As String = "e5719155-369f-4b3e-9e5e-99856449f05b"
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

#Region "Private Fields"

    Private _EditEvents As IEditEvents_Event

#End Region

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
    End Sub

#End Region

#Region "Friend Properties"

    Private Shared _Editor As IEditor

    Friend Shared ReadOnly Property Editor() As IEditor
        Get
            Return _Editor
        End Get
    End Property

    Private Shared _HasValidLicense As Boolean '= False

    Friend Shared Property HasValidLicense() As Boolean
        Get
            Return _HasValidLicense
        End Get
        Set(ByVal value As Boolean)
            _HasValidLicense = value
        End Set
    End Property

    Private Shared _IsValidWorkspace As Boolean '= False

    Friend Shared Property IsValidWorkspace() As Boolean
        Get
            Return _IsValidWorkspace
        End Get
        Set(ByVal value As Boolean)
            _IsValidWorkspace = value
        End Set
    End Property

    Private Shared _CanEditTaxlots As Boolean = True

    Friend Shared Property CanEditTaxlots() As Boolean
        Get
            Return _CanEditTaxlots
        End Get
        Set(ByVal value As Boolean)
            _CanEditTaxlots = value
        End Set
    End Property

    Private Shared _CanAutoUpdate As Boolean = True

    Friend Shared Property CanAutoUpdate() As Boolean
        Get
            Return _CanAutoUpdate
        End Get
        Set(ByVal value As Boolean)
            _CanAutoUpdate = value
        End Set
    End Property

    Private Shared _CanAutoUpdateAllFields As Boolean = True

    Friend Shared Property CanAutoUpdateAllFields() As Boolean
        Get
            Return _CanAutoUpdateAllFields
        End Get
        Set(ByVal value As Boolean)
            _CanAutoUpdateAllFields = value
        End Set
    End Property

#End Region

#Region "IExtension Interface Implementations"

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
        Get
            Return "EditorExtension"
        End Get
    End Property

    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        _Editor = Nothing
        _EditEvents = Nothing
    End Sub

    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        If Not initializationData Is Nothing AndAlso TypeOf initializationData Is IEditor Then
            _Editor = CType(initializationData, IEditor)

            'Wire up editor events.
            _EditEvents = CType(_Editor, IEditEvents_Event)
            AddHandler _EditEvents.OnStartEditing, AddressOf EditEvents_OnStartEditing
        End If
    End Sub

#End Region

#Region "IExtensionAccelerators Interface Implementations"

    Public Sub CreateAccelerators() Implements ESRI.ArcGIS.esriSystem.IExtensionAccelerators.CreateAccelerators
        ' Create the keyboard accelerators for this extension.
        ' TODO: Test this (not sure this will work with an editor extension)
        Dim key As Integer
        Dim usesCtrl As Boolean
        Dim usesAlt As Boolean
        Dim usesShift As Boolean
        Dim uid As New UID
        Dim doc As IDocument = EditorExtension.Editor.Parent.Document
        Dim acceleratorTable As IAcceleratorTable = doc.Accelerators

        ' Set TaxlotAssignment accelerator keys to Ctrl + Shift + T
        key = Convert.ToInt32(System.Windows.Forms.Keys.T)
        usesCtrl = True
        usesAlt = False
        usesShift = True
        uid.Value = "{" & OrmapTaxlotEditing.TaxlotAssignment.ClassId & "}"
        SetAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set LocateFeature accelerator keys to Ctrl + Shift + L
        key = Convert.ToInt32(System.Windows.Forms.Keys.L)
        usesCtrl = True
        usesAlt = False
        usesShift = True
        uid.Value = "{" & OrmapTaxlotEditing.LocateFeature.ClassId & "}"
        SetAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set EditMapIndex accelerator keys to Ctrl + Shift + E
        key = Convert.ToInt32(System.Windows.Forms.Keys.E)
        usesCtrl = True
        usesAlt = False
        usesShift = True
        uid.Value = "{" & OrmapTaxlotEditing.EditMapIndex.ClassId & "}"
        SetAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set CombineTaxlots accelerator keys to Ctrl + Shift + C
        key = Convert.ToInt32(System.Windows.Forms.Keys.C)
        usesCtrl = True
        usesAlt = False
        usesShift = True
        uid.Value = "{" & OrmapTaxlotEditing.CombineTaxlots.ClassId & "}"
        SetAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set AddArrows accelerator keys to Ctrl + Shift + A
        key = Convert.ToInt32(System.Windows.Forms.Keys.A)
        usesCtrl = True
        usesAlt = False
        usesShift = True
        uid.Value = "{" & OrmapTaxlotEditing.AddArrows.ClassId & "}"
        SetAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

    End Sub

#End Region

#Region "IPersistVariant Interface Implementations"

    Public ReadOnly Property ID() As ESRI.ArcGIS.esriSystem.UID Implements ESRI.ArcGIS.esriSystem.IPersistVariant.ID
        Get
            Dim uid As UID = New UIDClass()
            uid.Value = "{" & OrmapTaxlotEditing.EditorExtension.ClassId & "}"
            Return uid
        End Get
    End Property

    Public Sub Load(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Load

        If Stream Is Nothing Then
            Throw New ArgumentNullException("Stream")
        End If

        _CanEditTaxlots = CBool(Stream.Read())
        _CanAutoUpdate = CBool(Stream.Read())
        _CanAutoUpdateAllFields = CBool(Stream.Read())

    End Sub

    Public Sub Save(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Save

        If Stream Is Nothing Then
            Throw New ArgumentNullException("Stream")
        End If

        Stream.Write(_CanEditTaxlots)
        Stream.Write(_CanAutoUpdate)
        Stream.Write(_CanAutoUpdateAllFields)

    End Sub

#End Region

#Region "Editor Event Handlers"

    Private Sub EditEvents_OnStartEditing()
        ' Test for valid workspace and license
        If _Editor.EditWorkspace.Type = esriWorkspaceType.esriFileSystemWorkspace Then
            _IsValidWorkspace = False
        Else
            _IsValidWorkspace = True
            'Wire up editor events.
            AddHandler _EditEvents.OnChangeFeature, AddressOf EditEvents_OnChangeFeature
            AddHandler _EditEvents.OnCreateFeature, AddressOf EditEvents_OnCreateFeature
        End If
        _HasValidLicense = (ValidateLicense(esriLicenseProductCode.esriLicenseProductCodeArcEditor) OrElse _
                ValidateLicense(esriLicenseProductCode.esriLicenseProductCodeArcInfo))
    End Sub

    Private Sub EditEvents_OnChangeFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject)
        ' TODO: Connect to field AutoUpdate, etc. (see VB6 code)
    End Sub

    Private Sub EditEvents_OnCreateFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject)
        ' TODO: Connect to field AutoUpdate, etc. (see VB6 code)
    End Sub

#End Region

#Region "Private Methods"

    ' TODO: Test (not sure this how this will work with editor extension)
    Private Shared Sub SetAccelerator(ByRef acceleratorTable As IAcceleratorTable, _
            ByVal classID As UID, ByVal key As Integer, _
            ByVal usesCtrl As Boolean, ByVal usesAlt As Boolean, _
            ByVal usesShift As Boolean)
        ' Create accelerator only if nothing else is using it

        Dim accelerator As IAccelerator

        accelerator = acceleratorTable.FindByKey(key, usesCtrl, usesAlt, usesShift)
        If accelerator Is Nothing Then
            'The clsid of one of the commands in the ext
            acceleratorTable.Add(classID, key, usesCtrl, usesAlt, usesShift)
        End If

    End Sub

    Private Shared Function ValidateLicense(ByVal requiredProductCode As esriLicenseProductCode) As Boolean
        ' Validate the license (e.g. ArcEditor or ArcInfo).

        Dim aoInitTestProduct As IAoInitialize = New AoInitializeClass()
        Dim productCode As esriLicenseProductCode = aoInitTestProduct.InitializedProduct()

        If productCode = requiredProductCode Then
            Return True
        Else
            Return False
        End If

    End Function

#End Region

End Class

