Imports System
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI

<ComVisible(True)> _
<ComClass(PropertyPage.ClassId, PropertyPage.InterfaceId, PropertyPage.EventsId), _
ProgId("ORMAPTaxlotEditing.PropertyPage")> _
Public NotInheritable Class PropertyPage
    Implements IComPropertyPage

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "050c23da-ebd8-4a1d-871b-b7a9354d331b"
    Public Const InterfaceId As String = "bae36023-8a03-43b6-bea6-fab534ff7c5e"
    Public Const EventsId As String = "8ab94224-407b-4139-a003-48f5789bf3b3"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisible(False)> _
    Private Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        '
        ' TODO: Add any COM registration code here
        '
    End Sub

    <ComUnregisterFunction(), ComVisible(False)> _
    Private Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        '
        ' TODO: Add any COM unregistration code here
        '
    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        EditorPropertyPages.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        EditorPropertyPages.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "Private Fields"

    Private _pageDirty As Boolean '= False
    Private _propPageSite As IComPropertyPageSite
    Private WithEvents _propForm As PropertiesForm

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

    Friend Property PropertiesForm() As PropertiesForm
        Get
            Return _propForm
        End Get
        Set(ByVal value As PropertiesForm)
            _propForm = value
        End Set
    End Property

#End Region

#Region "IComPropertyPage Implementations"

    Public ReadOnly Property Height() As Integer Implements IComPropertyPage.Height
        Get
            Return _propForm.Height
        End Get
    End Property

    Public ReadOnly Property HelpFile() As String Implements IComPropertyPage.HelpFile
        Get
            Return Nothing  ' TODO: Implement Help File
        End Get
    End Property

    Public ReadOnly Property HelpContextID(ByVal controlID As Integer) As Integer Implements IComPropertyPage.HelpContextID
        Get
            Return 0  ' TODO: Implement Help File
        End Get
    End Property

    Public ReadOnly Property IsPageDirty() As Boolean Implements IComPropertyPage.IsPageDirty
        Get
            Return _pageDirty
        End Get
    End Property

    Public WriteOnly Property PageSite() As ESRI.ArcGIS.Framework.IComPropertyPageSite Implements IComPropertyPage.PageSite
        Set(ByVal value As ESRI.ArcGIS.Framework.IComPropertyPageSite)
            _propPageSite = value
        End Set
    End Property

    Public Property Priority() As Integer Implements IComPropertyPage.Priority
        Get
            Return 0  'Lowest number = last/rightmost tab position in the Properties window.
        End Get
        Set(ByVal value As Integer)
        End Set
    End Property

    Public Property Title() As String Implements IComPropertyPage.Title
        Get
            Return "ORMAP Taxlot Editor"
        End Get
        Set(ByVal value As String)
        End Set
    End Property

    Public ReadOnly Property Width() As Integer Implements IComPropertyPage.Width
        Get
            Return _propForm.Width
        End Get
    End Property

    Public Function Activate() As Integer Implements IComPropertyPage.Activate
        Return _propForm.Handle.ToInt32()
    End Function

    Public Function Applies(ByVal objects As ESRI.ArcGIS.esriSystem.ISet) As Boolean Implements IComPropertyPage.Applies

        ' Do not affirm if the objects list is empty.
        If objects Is Nothing OrElse objects.Count = 0 Then
            Return False
        End If
        objects.Reset()

        ' Get a reference to the editor.
        ' Do not affirm if the editor is not found.
        Dim editor As IEditor = TryCast(objects.Next(), IEditor)
        If editor Is Nothing Then
            Return False
        End If

        ' Do not affirm if the user is not editing.
        If editor.EditState <> esriEditState.esriStateEditing Then
            Return False
        End If

        ' Do not affirm if the user is editing a file-based workspace (e.g. coverages, shapefiles).
        If editor.EditWorkspace.Type = esriWorkspaceType.esriFileSystemWorkspace Then
            Return False
        End If

        ' Otherwise, affirm.
        Return True

    End Function

    Public Sub Apply() Implements IComPropertyPage.Apply
        ' Write to the EditorExtension.CanEdit shared (i.e. by all class objects) property
        EditorExtension.CanEditTaxlots = _propForm.uxEnableTools.Checked
        EditorExtension.CanAutoUpdate = _propForm.uxEnableAutoUpdate.Checked
        EditorExtension.CanAutoUpdateAllFields = Not _propForm.uxAllFieldsOption.Checked
        EditorExtension.CanAutoUpdateAllFields = _propForm.uxAllFieldsOption.Checked
        _pageDirty = False
    End Sub

    Public Sub Cancel() Implements IComPropertyPage.Cancel
        ' TODO: Implement this?
    End Sub

    Public Sub Deactivate() Implements IComPropertyPage.Deactivate
        If Not _propForm Is Nothing Then
            _propForm.Dispose()
        End If
        _propForm = Nothing
        _propPageSite = Nothing
    End Sub

    Public Sub Hide() Implements IComPropertyPage.Hide
        _propForm.Hide()
    End Sub

    Public Sub SetObjects(ByVal objects As ESRI.ArcGIS.esriSystem.ISet) Implements IComPropertyPage.SetObjects
        ' Note: The Applies() method should have done preliminary checking of 
        ' editor states before this method is called.

        ' TODO: Move (to where)?
        _propForm = New PropertiesForm()
        _propForm.uxEnableTools.Checked = EditorExtension.CanEditTaxlots
        _propForm.uxEnableAutoUpdate.Checked = EditorExtension.CanAutoUpdate
        _propForm.uxMinimumFieldsOption.Checked = Not EditorExtension.CanAutoUpdateAllFields
        _propForm.uxAllFieldsOption.Checked = EditorExtension.CanAutoUpdateAllFields

        ' Wire up form events.
        AddHandler _propForm.uxEnableTools.CheckedChanged, AddressOf uxEnableTools_CheckedChanged
        AddHandler _propForm.uxEnableAutoUpdate.CheckedChanged, AddressOf uxEnableAutoUpdate_CheckedChanged

    End Sub

    Public Sub Show() Implements IComPropertyPage.Show
        _propForm.Show()
    End Sub

#End Region

#Region "Private Methods"

    Private Sub uxEnableTools_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)

        _propForm.uxEnableAutoUpdate.Enabled = _propForm.uxEnableTools.Checked
        _propForm.uxMinimumFieldsOption.Enabled = _propForm.uxEnableTools.Checked
        _propForm.uxAllFieldsOption.Enabled = _propForm.uxEnableTools.Checked

        ' Set dirty flag.
        _pageDirty = True

        If Not _propPageSite Is Nothing Then
            _propPageSite.PageChanged()
        End If
    End Sub

    Private Sub uxEnableAutoUpdate_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)

        _propForm.uxMinimumFieldsOption.Enabled = _propForm.uxEnableAutoUpdate.Checked
        _propForm.uxAllFieldsOption.Enabled = _propForm.uxEnableAutoUpdate.Checked

        ' Set dirty flag.
        _pageDirty = True

        If Not _propPageSite Is Nothing Then
            _propPageSite.PageChanged()
        End If
    End Sub

#End Region

End Class

