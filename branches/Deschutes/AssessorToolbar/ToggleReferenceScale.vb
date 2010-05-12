Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.carto
Imports System.Windows.forms
Imports AssessorToolbar.Utilities

<ComClass(ToggleReferenceScale.ClassId, ToggleReferenceScale.InterfaceId, ToggleReferenceScale.EventsId), _
 ProgId("AssessorToolbar.ToggleReferenceScale")> _
Public NotInheritable Class ToggleReferenceScale
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "118e84bd-78c9-49a6-95ee-6966a356f912"
    Public Const InterfaceId As String = "c74d6563-82a0-4221-b765-a37a74e7f6f6"
    Public Const EventsId As String = "3d23376a-56c4-480e-b39d-465093b374e6"
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
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

    Private _application As IApplication
    Private _buttonChecked As Boolean

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "AssessorToolbar"  'localizable text 
        MyBase.m_caption = "ToggleReferenceScale"   'localizable text 
        MyBase.m_message = "Sets/Clears The Data Frame Reference Scale"   'localizable text 
        MyBase.m_toolTip = "Sets/Clears The Data Frame Reference Scale" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_ToggleReferenceScale"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")


        Try
            'TODO: change bitmap name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            'Dim bitmapResourceName As String = "page.bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try


    End Sub


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

        ' TODO:  Add other initialization code
    End Sub

    Public Overrides Sub OnClick()


        Dim theMxDoc As IMxDocument = _application.Document
        Dim theMap As IMap = theMxDoc.FocusMap
        Dim theActiveView As IActiveView = theMxDoc.FocusMap

        '-- Check to see if the tool is not active but the reference scale was already set
        If Me.m_checked = False AndAlso theMap.ReferenceScale <> 0 AndAlso ReferenceScale = 0 Then
            ReferenceScale = theMap.ReferenceScale
            Me.m_checked = True
            MessageBox.Show("Note:  The reference scale was set in a previous session and is still set to 1:" & ReferenceScale & ".  Please make sure this is correct for the map your working on.  If not you may want to run the Filter Anno Scale tool again.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If ReferenceScale = 0 Then
            Dim theSelectMapScaleDialog As FilterAnnoScaleForm = New FilterAnnoScaleForm
            If theSelectMapScaleDialog.ShowDialog = DialogResult.Cancel Then Exit Sub
            Dim theMapScale As Integer = Convert.ToInt32(theSelectMapScaleDialog.uxMapScale.Text) * 12
            ReferenceScale = theMapScale
        End If

        If theMap.ReferenceScale = 0 Then
            theMap.ReferenceScale = ReferenceScale
            Me.m_checked = True
        Else
            theMap.ReferenceScale = 0
            Me.m_checked = False
        End If

        theActiveView.Refresh()

    End Sub




End Class



