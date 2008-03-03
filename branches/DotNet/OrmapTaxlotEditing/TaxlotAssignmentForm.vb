Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ADF.CATIDs

<ComClass(TaxlotAssignmentForm.ClassId, TaxlotAssignmentForm.InterfaceId, TaxlotAssignmentForm.EventsId), _
 ProgId("OrmapTaxlotEditing.TaxlotAssignmentForm")> _
Public Class TaxlotAssignmentForm
    Implements IDockableWindowDef

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "e844dd61-81e9-4ca6-ade5-c55bde8ec51e"
    Public Const InterfaceId As String = "78448d32-4b12-4178-b52c-2b200857a330"
    Public Const EventsId As String = "6e2de113-be00-47bc-9a35-f03ee692a972"
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
        MxDockableWindows.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxDockableWindows.Unregister(regKey)

    End Sub

#End Region
#End Region

    Private _application As IApplication

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'Populate multi-value controls and set default settings
        
        'Populate multi-value controls
        uxType.Items.Add("NUMBER")
        uxType.Items.Add("ROADS")
        uxType.Items.Add("WATER")
        uxType.Items.Add("RAILS")
        uxType.Items.Add("NONTL")

        ' Set control defaults
        uxType.Text = "NUMBER"
        uxIncrementByNone.Checked = True
        uxStartingFrom.Text = "0"

    End Sub

#Region "IDockableWindowDef Members"

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.Framework.IDockableWindowDef.Caption
        Get
            'Note: Can replace with locale-based initial title bar caption
            Return "Taxlot Assignment"
        End Get
    End Property

    Public ReadOnly Property ChildHWND() As Integer Implements ESRI.ArcGIS.Framework.IDockableWindowDef.ChildHWND
        Get
            Return Me.Handle.ToInt32()
        End Get
    End Property

    Public ReadOnly Property Name1() As String Implements ESRI.ArcGIS.Framework.IDockableWindowDef.Name
        Get
            'Note: Can replace with any non-localizable string
            Return Me.Name
        End Get
    End Property

    Public Sub OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.Framework.IDockableWindowDef.OnCreate
        _application = CType(hook, IApplication)
    End Sub

    Public Sub OnDestroy() Implements ESRI.ArcGIS.Framework.IDockableWindowDef.OnDestroy
        'TODO: NIS Release resources and call dispose of any ActiveX control initialized (needed?)
    End Sub

    Public ReadOnly Property UserData() As Object Implements ESRI.ArcGIS.Framework.IDockableWindowDef.UserData
        Get
            Return Nothing
        End Get
    End Property
#End Region


    ' TODO: NIS Add Standard Regions

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uxHelp.Click
        'TODO: NIS Finish implementation of this -- could be replaced with new help mechanism.
        'Open a custom help file
        'Requires a file in the same dir as the application dll
        Dim filePath As String
        filePath = My.Application.Info.DirectoryPath & "\" & "TaxlotAssignmentHelp.rtf"
        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(filePath) Then
            'TODO: NIS Implement open of rtf document
            MessageBox.Show("Implement function to open help file from the current directory.")
            'gsb_StartDoc(Me.Handle.ToInt32, filePath)
        Else
            MessageBox.Show("No help file available in the current directory.")
        End If
    End Sub

End Class
