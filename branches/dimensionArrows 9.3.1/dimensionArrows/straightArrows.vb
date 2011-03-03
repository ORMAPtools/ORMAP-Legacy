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
''' Places a pair of opposing straight arrows
''' </summary>
''' <remarks></remarks>
<ComClass(straightArrows.ClassId, straightArrows.InterfaceId, straightArrows.EventsId), _
 ProgId("dimensionArrows.straightArrows")> _
Public NotInheritable Class straightArrows
  Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4dc3bfa0-237e-483a-91a8-83759cf73be2"
    Public Const InterfaceId As String = "be12f4a7-1ac7-400b-add8-2865de198787"
    Public Const EventsId As String = "c3f2b698-a521-4778-8856-c2a3557fcf32"
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

#Region "Module Variables"
    Private _dStartEditing As IEditEvents_OnStartEditingEventHandler
    Private _dStopEditing As IEditEvents_OnStopEditingEventHandler
    Private kill As Boolean = False
#End Region


    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        MyBase.m_category = "OR-DOR"
        MyBase.m_caption = "Straight width arrows"
        MyBase.m_message = "Straight width arrows"
        MyBase.m_toolTip = "Place two straight width arrows"
        MyBase.m_name = "OR-DOR_straightArrows"
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
        End If

        Dim arrowUID As New ESRI.ArcGIS.esriSystem.UID
        arrowUID.Value = "dimensionarrows.dimensionArrowExtension"
        Dim arrowExtension As ESRI.ArcGIS.esriSystem.IExtension
        arrowExtension = _app.FindExtensionByCLSID(arrowUID)

        Dim editorUID As New ESRI.ArcGIS.esriSystem.UID
        editorUID.Value = "esriEditor.editor"
        _editor = _app.FindExtensionByCLSID(editorUID)
        _editEvents = _editor

        _dStartEditing = New IEditEvents_OnStartEditingEventHandler( _
            AddressOf onStartEditing)
        AddHandler _editEvents.OnStartEditing, _dStartEditing

        _dStopEditing = New IEditEvents_OnStopEditingEventHandler( _
            AddressOf onStopEditing)
        AddHandler _editEvents.OnStopEditing, _dStopEditing

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
            _thisArrow.category = arrowCategories.Straight
        End If
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
      ByVal X As Integer, ByVal Y As Integer)
        Dim point As IPoint = getDataFrameCoords(X, Y)

        If Button = EsriMouseButtons.Left Then
            If _pointNumber = 1 Then
                arrowUtilities.saveSubtype()
                Dim arrowSubtype As Long = findArrowSubtype()
                If arrowSubtype = Nothing Then Exit Sub
                setArrowSubtype(arrowSubtype)
                setLineFeedback(point)
                _pointNumber = 2
            Else
                placeArrows(_lastPoint)
                arrowUtilities.resetSubtype()
                _pointNumber = 1
            End If
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

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
          ByVal X As Integer, ByVal Y As Integer)

        Dim point As IPoint = getDataFrameCoords(X, Y)
        point = getSnapPoint(point)
        If _pointNumber = 2 Then
            showLineFeedback(point)
        End If
    End Sub

    Public Overrides Sub OnKeyDown(ByVal keyCode As Integer, ByVal Shift As Integer)
        keyCommands(keyCode, Shift)
    End Sub

    Public Overrides Function OnContextMenu(ByVal X As Integer, ByVal Y As Integer) _
    As Boolean
        CreateContextMenu(_app)
        Return True
    End Function

    '''<summary>Create a context menu</summary>
    '''  
    '''<param name="application"></param>
    '''   
    '''<remarks></remarks>
    Public Sub CreateContextMenu(ByVal application As IApplication)
        Dim commandBars As ICommandBars = application.Document.CommandBars
        Dim commandBar As ICommandBar = commandBars.Create("TemporaryContextMenu", _
            esriCmdBarType.esriCmdBarTypeShortcutMenu)

        Dim optionalIndex As System.Object = Type.Missing
        Dim uid As ESRI.ArcGIS.esriSystem.UID = New ESRI.ArcGIS.esriSystem.UIDClass

        uid.Value = "dimensionArrows.ContextMenu"
        uid.SubType = 0
        commandBar.Add(uid, optionalIndex)
        ReDim _menuitems(18)
        _menuitems(0) = SHORTER
        _menuitems(1) = LONGER
        _menuitems(2) = FLIP
        _menuitems(3) = UNLOCK
        _menuitems(4) = SWITCH
        _menuitems(5) = CANCEL
        _menuitems(6) = SEPARATOR
        _menuitems(7) = SCALE10
        _menuitems(8) = SCALE20
        _menuitems(9) = SCALE30
        _menuitems(10) = SCALE40
        _menuitems(11) = SCALE50
        _menuitems(12) = SCALE100
        _menuitems(13) = SCALE200
        _menuitems(14) = SCALE400
        _menuitems(15) = SCALE800
        _menuitems(16) = SCALE1000
        _menuitems(17) = SCALE2000
        _menuitems(18) = HELP

        commandBar.Popup()
    End Sub
End Class

