
#Region "Copyright 2008 ORMAP Tech Group"

' File: EditorExtension.vb

' Author: .NET Migration Team (Shad Campbell, James Moore, Nick Seigal)
' Created: January 8, 2008

' All rights reserved. Reproduction or transmission of this file, or a portion thereof,
' is forbidden without prior written permission of the ORMAP Tech Group.

#End Region

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

#Region "Class-Level Constants And Enumerations"

#End Region

#Region "Built-In Class Members (Properties, Methods, Events, Event Handlers, Delegates, Etc.)"

#End Region

#Region "Custom Class Members"

#End Region

#Region "Inherited Class Members"

#End Region

#Region "Implemented Interface Members"

#End Region

#Region "Other Members"

#End Region


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

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
    End Sub

#End Region

#Region "Friend Properties"

    Private Shared _editor As IEditor

    Friend Shared ReadOnly Property Editor() As IEditor
        Get
            Return _editor
        End Get
    End Property

    Private Sub SetEditor(ByVal value As IEditor)
        ' TODO: Add validation code?
        _editor = value
    End Sub

    Private Shared _editEvents As IEditEvents_Event

    Friend Shared ReadOnly Property EditEvents() As IEditEvents_Event
        Get
            Return _editEvents
        End Get
    End Property

    Private Sub SetEditEvents(ByVal value As IEditEvents_Event)
        ' TODO: Add validation code?
        _editEvents = value
    End Sub

    Private Shared _hasValidLicense As Boolean '= False

    Friend Shared Property HasValidLicense() As Boolean
        Get
            Return _hasValidLicense
        End Get
        Set(ByVal value As Boolean)
            _hasValidLicense = value
        End Set
    End Property

    Private Shared _isValidWorkspace As Boolean '= False

    Friend Shared Property IsValidWorkspace() As Boolean
        Get
            Return _isValidWorkspace
        End Get
        Set(ByVal value As Boolean)
            _isValidWorkspace = value
        End Set
    End Property

    Private Shared _canEditTaxlots As Boolean = True

    Friend Shared Property CanEditTaxlots() As Boolean
        Get
            Return _canEditTaxlots
        End Get
        Set(ByVal value As Boolean)
            _canEditTaxlots = value
        End Set
    End Property

    Private Shared _canAutoUpdate As Boolean = True

    Friend Shared Property CanAutoUpdate() As Boolean
        Get
            Return _canAutoUpdate
        End Get
        Set(ByVal value As Boolean)
            _canAutoUpdate = value
        End Set
    End Property

    Private Shared _canAutoUpdateAllFields As Boolean = True

    Friend Shared Property CanAutoUpdateAllFields() As Boolean
        Get
            Return _canAutoUpdateAllFields
        End Get
        Set(ByVal value As Boolean)
            _canAutoUpdateAllFields = value
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
        SetEditor(Nothing)
        SetEditEvents(Nothing)
    End Sub

    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        If Not initializationData Is Nothing AndAlso TypeOf initializationData Is IEditor Then
            SetEditor(CType(initializationData, IEditor))

            ' Subscribe to editor events.
            SetEditEvents(CType(EditorExtension.Editor, IEditEvents_Event))
            AddHandler _editEvents.OnStartEditing, AddressOf EditEvents_OnStartEditing
            AddHandler _editEvents.OnStopEditing, AddressOf EditEvents_OnStopEditing
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

        ' Set LocateFeature accelerator keys to Ctrl + Alt + L
        key = Convert.ToInt32(System.Windows.Forms.Keys.L)
        usesCtrl = True
        usesAlt = True
        usesShift = False
        uid.Value = "{" & OrmapTaxlotEditing.LocateFeature.ClassId & "}"
        SetAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set TaxlotAssignment accelerator keys to Ctrl + Alt + T
        key = Convert.ToInt32(System.Windows.Forms.Keys.T)
        usesCtrl = True
        usesAlt = True
        usesShift = False
        uid.Value = "{" & OrmapTaxlotEditing.TaxlotAssignment.ClassId & "}"
        SetAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set EditMapIndex accelerator keys to Ctrl + Alt + E
        key = Convert.ToInt32(System.Windows.Forms.Keys.E)
        usesCtrl = True
        usesAlt = True
        usesShift = False
        uid.Value = "{" & OrmapTaxlotEditing.EditMapIndex.ClassId & "}"
        SetAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set CombineTaxlots accelerator keys to Ctrl + Alt + C
        key = Convert.ToInt32(System.Windows.Forms.Keys.C)
        usesCtrl = True
        usesAlt = True
        usesShift = False
        uid.Value = "{" & OrmapTaxlotEditing.CombineTaxlots.ClassId & "}"
        SetAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set AddArrows accelerator keys to Ctrl + Alt + A
        key = Convert.ToInt32(System.Windows.Forms.Keys.A)
        usesCtrl = True
        usesAlt = True
        usesShift = False
        uid.Value = "{" & OrmapTaxlotEditing.AddArrows.ClassId & "}"
        SetAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

    End Sub

#End Region

#Region "IPersistVariant Interface Implementations"

    Public ReadOnly Property ID() As ESRI.ArcGIS.esriSystem.UID Implements ESRI.ArcGIS.esriSystem.IPersistVariant.ID
        Get
            Dim uid As New UIDClass()
            uid.Value = "{" & OrmapTaxlotEditing.EditorExtension.ClassId & "}"
            Return uid
        End Get
    End Property

    Public Sub Load(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Load

        If Stream Is Nothing Then
            Throw New ArgumentNullException("Stream")
        End If

        CanEditTaxlots = CBool(Stream.Read())
        CanAutoUpdate = CBool(Stream.Read())
        CanAutoUpdateAllFields = CBool(Stream.Read())

    End Sub

    Public Sub Save(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Save

        If Stream Is Nothing Then
            Throw New ArgumentNullException("Stream")
        End If

        Stream.Write(CanEditTaxlots)
        Stream.Write(CanAutoUpdate)
        Stream.Write(CanAutoUpdateAllFields)

    End Sub

#End Region

#Region "Editor Event Handlers"

    Private Sub EditEvents_OnChangeFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject)
        ' TODO: Connect to field AutoUpdate, etc. (see VB6 code)
    End Sub

    Private Sub EditEvents_OnCreateFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject)
        ' TODO: Connect to field AutoUpdate, etc. (see VB6 code)
    End Sub

    Private Sub EditEvents_OnStartEditing()
        ' Test for valid workspace and license
        If EditorExtension.Editor.EditWorkspace.Type = esriWorkspaceType.esriFileSystemWorkspace Then
            IsValidWorkspace = False
        Else
            IsValidWorkspace = True
            ' Subscribe to editor events.
            AddHandler EditEvents.OnChangeFeature, AddressOf EditEvents_OnChangeFeature
            AddHandler EditEvents.OnCreateFeature, AddressOf EditEvents_OnCreateFeature

            ' Set up document keyboard accelerators for extension commands.
            CreateAccelerators()
        End If
        _hasValidLicense = (ValidateLicense(esriLicenseProductCode.esriLicenseProductCodeArcEditor) OrElse _
                ValidateLicense(esriLicenseProductCode.esriLicenseProductCodeArcInfo))
    End Sub

    Private Sub EditEvents_OnStopEditing(ByVal save As Boolean)
        ' Unsubscribe to editor events.
        RemoveHandler EditEvents.OnChangeFeature, AddressOf EditEvents_OnChangeFeature
        RemoveHandler EditEvents.OnCreateFeature, AddressOf EditEvents_OnCreateFeature
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

        Dim aoInitTestProduct As New AoInitializeClass()
        Dim productCode As esriLicenseProductCode = aoInitTestProduct.InitializedProduct()

        Return (productCode = requiredProductCode)
    End Function

#End Region

End Class

