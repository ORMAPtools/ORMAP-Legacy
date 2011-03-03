Imports ESRI.ArcGIS.SystemUI
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

''' <summary>
''' Creates a land hook
''' </summary>
''' <remarks>The basic land hook definition is stored in the XML file</remarks>
<ComClass(landHook.ClassId, landHook.InterfaceId, landHook.EventsId), _
 ProgId("dimensionArrows.landHook")> _
Public NotInheritable Class landHook
    Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "7ea0f36f-c44d-4f57-aaa4-f90fa3eae08f"
    Public Const InterfaceId As String = "c0beb700-2ee7-4e19-b162-fe9dc3117e0f"
    Public Const EventsId As String = "63a094f5-ad1a-4118-ae49-647d61b5faf1"
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

    Private _dStartEditing As IEditEvents_OnStartEditingEventHandler
    Private _dStopEditing As IEditEvents_OnStopEditingEventHandler
    Private kill As Boolean = False

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        MyBase.m_category = "OR-DOR"
        MyBase.m_caption = "Land hook"
        MyBase.m_message = "Place a land hook"
        MyBase.m_toolTip = "Place a land hook"
        MyBase.m_name = "OR-DOR_landHook"
        MyBase.m_enabled = False
        _installationFolder = System.IO.Path.GetDirectoryName(Me.GetType().Assembly.Location)

        Try
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
            MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), _
                Me.GetType().Name + ".cur")
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            _app = CType(hook, IApplication)

            Dim editorUID As New ESRI.ArcGIS.esriSystem.UID
            editorUID.Value = "esriEditor.editor"

            Dim extension As ESRI.ArcGIS.esriSystem.IExtension
            extension = _app.FindExtensionByCLSID(editorUID)

            _editor = _app.FindExtensionByCLSID(editorUID)
            _editEvents = _editor

            _dStartEditing = New IEditEvents_OnStartEditingEventHandler( _
         AddressOf onStartEditing)
            AddHandler _editEvents.OnStartEditing, _dStartEditing

            _dStopEditing = New IEditEvents_OnStopEditingEventHandler( _
             AddressOf onStopEditing)
            AddHandler _editEvents.OnStopEditing, _dStopEditing

        End If
    End Sub

    Private Sub onStartEditing()
        MyBase.m_enabled = True
    End Sub

    Private Sub onStopEditing(ByVal save As Boolean)
        MyBase.m_enabled = False
    End Sub

    Public Overrides Sub OnClick()
        If checkDataView() = False Then
            kill = True
            setDefaultTool()
        Else
            clearAll()
            _thisArrow.category = arrowCategories.LandHook
        End If
    End Sub

    ''' <summary>
    ''' Clean up the screen and reset the subtype if the command exits
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function Deactivate() As Boolean
        If Not kill Then
            clearAll()
            resetSubtype()
        End If
        Return True
    End Function

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
          ByVal X As Integer, ByVal Y As Integer)
        If Button = EsriMouseButtons.Left Then
            If _pointNumber = 1 Then
                saveSubtype()
                setArrowSubtype(101)
                setLineFeedback(getDataFrameCoords(X, Y))
                _pointNumber = 2
            Else
                placeArrows(getDataFrameCoords(X, Y))
                resetSubtype()
                _pointNumber = 1
            End If
        End If
    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
      ByVal X As Integer, ByVal Y As Integer)
        If _pointNumber = 2 Then
            _angleIsSet = False
            showLineFeedback(getDataFrameCoords(X, Y))
        End If
    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
        ByVal X As Integer, ByVal Y As Integer)
    End Sub

    Public Overrides Sub OnKeyDown(ByVal keyCode As Integer, ByVal Shift As Integer)
        keyCommands(keyCode, Shift)
    End Sub

    Public Overrides Function OnContextMenu(ByVal X As Integer, ByVal Y As Integer) _
        As Boolean
        CreateContextMenu(_app)
        Return True
    End Function

    '''<summary>Create a context menu for this class</summary>
    '''  
    '''<param name="application"></param>
    '''   
    '''<remarks>Uses constants to fill in the list</remarks>
    Public Sub CreateContextMenu(ByVal application As IApplication)
        Dim commandBars As ICommandBars = application.Document.CommandBars
        Dim commandBar As ICommandBar = commandBars.Create("TemporaryContextMenu", _
            esriCmdBarType.esriCmdBarTypeShortcutMenu)

        Dim optionalIndex As System.Object = Type.Missing
        Dim uid As ESRI.ArcGIS.esriSystem.UID = New ESRI.ArcGIS.esriSystem.UIDClass

        uid.Value = "dimensionArrows.ContextMenu"
        uid.SubType = 0
        commandBar.Add(uid, optionalIndex)
        ReDim _menuitems(16)
        _menuitems(0) = SHORTER
        _menuitems(1) = LONGER
        _menuitems(2) = FLIP
        _menuitems(3) = CANCEL
        _menuitems(4) = SEPARATOR
        _menuitems(5) = SCALE10
        _menuitems(6) = SCALE20
        _menuitems(7) = SCALE30
        _menuitems(8) = SCALE40
        _menuitems(9) = SCALE50
        _menuitems(10) = SCALE100
        _menuitems(11) = SCALE200
        _menuitems(12) = SCALE400
        _menuitems(13) = SCALE800
        _menuitems(14) = SCALE1000
        _menuitems(15) = SCALE2000
        _menuitems(16) = HELP

        commandBar.Popup()
    End Sub

End Class

