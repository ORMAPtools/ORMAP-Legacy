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
''' Place a pair of curved dimension arrows for bearings and distances.
''' Two clicks position the points of the arrows and a third click sets the offset from
''' an imaginary line between the points.
''' 0 to 4 dashes may be added to the ends of the arrows using the keyboard.
''' </summary>
''' <remarks></remarks>
<ComClass(curvedArrows.ClassId, curvedArrows.InterfaceId, curvedArrows.EventsId), _
 ProgId("dimensionArrows.curvedArrows")> _
Public NotInheritable Class curvedArrows
    Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "f8270f75-86f3-4d8d-ad07-70d1e2976748"
    Public Const InterfaceId As String = "20ca56d9-10f5-4491-b43f-f1968e2fc4d2"
    Public Const EventsId As String = "fa7f17e4-8e79-4e5d-9347-4015887c0a76"
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
    Public prevPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
    Private kill As Boolean = False

    Public Sub New()
        MyBase.New()

        MyBase.m_category = "OR-DOR"
        MyBase.m_caption = "Curved dimension arrows"
        MyBase.m_message = "Place a pair of curved dimension arrows"
        MyBase.m_toolTip = "Place a pair of curved dimension arrows"
        MyBase.m_name = "OR-DOR_curvedArrows"
        MyBase.m_enabled = False

        _installationFolder = System.IO.Path.GetDirectoryName(Me.GetType().Assembly.Location)
        Try
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
            MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), "curved0.cur")
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
        _thisArrow.category = arrowCategories.NoDashes
    End Sub

    Private Sub onStopEditing(ByVal save As Boolean)
        MyBase.m_enabled = False
    End Sub

    Public Overrides Sub OnClick()
        If checkDataView() = False Then
            Kill = True
            setDefaultTool()
        Else
            clearAll()
            'default to no dashes
            _thisArrow.category = arrowCategories.NoDashes
            MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), "curved0.cur")
        End If
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
      ByVal X As Integer, ByVal Y As Integer)
        If Button = EsriMouseButtons.Left Then
            If _pointNumber = 1 Then
                saveSubtype()
                setArrowSubtype(134)
                setLineFeedback(getSnapPoint(getDataFrameCoords(X, Y)))
                _pointNumber = 2
            ElseIf _pointNumber = 2 Then
                showLineFeedback(_lastPoint)
                _pointNumber = 3
            Else
                setArrowOffset(X, Y)
                placeArrows(_lastPoint)
                prevPoint.PutCoords(0, 0)
                resetSubtype()
                _pointNumber = 1
            End If
        End If
    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
      ByVal X As Integer, ByVal Y As Integer)
        Dim geoPoint As IPoint = getDataFrameCoords(X, Y)

        getSnapPoint(geoPoint)
        _angleIsSet = False

        _editor.InvertAgent(prevPoint, 0)
        _editor.InvertAgent(geoPoint, 0)

        If _pointNumber = 3 Then
            setArrowOffset(X, Y)
            showLineFeedback(_lastPoint)
        ElseIf _pointNumber = 2 Then
            showLineFeedback(geoPoint)
        End If

        prevPoint = geoPoint
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

    ''' <summary>
    ''' Handles keyDown events specific to this class
    ''' </summary>
    ''' <param name="keyCode"></param>
    ''' <param name="Shift"></param>
    ''' <remarks></remarks>
    Public Overrides Sub OnKeyDown(ByVal keyCode As Integer, ByVal Shift As Integer)
        If Shift = 0 Then
            Select Case keyCode
                Case Windows.Forms.Keys.D0, Windows.Forms.Keys.NumPad0
                    MyBase.m_cursor = New System.Windows.Forms.Cursor( _
                        Me.GetType(), "curved0.cur")
                    _thisArrow.category = arrowCategories.NoDashes
                Case Windows.Forms.Keys.D1, Windows.Forms.Keys.NumPad1
                    _thisArrow.category = arrowCategories.OneDash
                    MyBase.m_cursor = New System.Windows.Forms.Cursor( _
                        Me.GetType(), "curved1.cur")
                Case Windows.Forms.Keys.D2, Windows.Forms.Keys.NumPad2
                    _thisArrow.category = arrowCategories.TwoDashes
                    MyBase.m_cursor = New System.Windows.Forms.Cursor( _
                        Me.GetType(), "curved2.cur")
                Case Windows.Forms.Keys.D3, Windows.Forms.Keys.NumPad3
                    _thisArrow.category = arrowCategories.ThreeDashes
                    MyBase.m_cursor = New System.Windows.Forms.Cursor( _
                        Me.GetType(), "curved3.cur")
                Case Windows.Forms.Keys.D4, Windows.Forms.Keys.NumPad4
                    _thisArrow.category = arrowCategories.FourDashes
                    MyBase.m_cursor = New System.Windows.Forms.Cursor( _
                        Me.GetType(), "curved4.cur")
                Case Else
                    keyCommands(keyCode, Shift)
            End Select

            If Not _pointNumber = 1 Then
                showLineFeedback(_lastPoint)
            End If
        Else
            keyCommands(keyCode, Shift)
        End If
    End Sub

    ''' <summary>
    ''' Sets the offset distance from an imaginary line between the two arrow points
    ''' </summary>
    ''' <param name="X"></param>
    ''' <param name="Y"></param>
    ''' <remarks></remarks>
    Private Sub setArrowOffset(ByVal X As Integer, ByVal Y As Integer)
        Dim displayTransformation As ESRI.ArcGIS.Display.IDisplayTransformation
        displayTransformation = _app.Display.DisplayTransformation

        Dim endPoint As ESRI.ArcGIS.Geometry.IPoint = _lastPoint

        'use a new line to get the angle of the vector between the points
        Dim vector As ILine = New Line
        vector.FromPoint = _firstPoint
        vector.ToPoint = endPoint

        'find the nearest point on the vector
        Dim nearPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
        nearPoint = displayTransformation.ToMapPoint(X, Y)

        Dim onlinePoint As IPoint = New ESRI.ArcGIS.Geometry.Point

        Dim rightSide As Boolean

        vector.QueryPointAndDistance( _
            esriSegmentExtension.esriExtendAtFrom, nearPoint, False, _
            Nothing, Nothing, _arrowOffset, rightSide)
        If rightSide Then _arrowOffset = _arrowOffset * -1
    End Sub

    Public Overrides Function OnContextMenu(ByVal X As Integer, ByVal Y As Integer) _
        As Boolean
        If _pointNumber <> 1 Then
            CreateContextMenu(_app)
            Return True
        End If
    End Function

    '''<summary>Create a context menu for this class</summary>
    '''  
    '''<param name="application">An IApplication Interface</param>
    '''   
    '''<remarks>Uses constant values to fill in the list</remarks>
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
        _menuitems(2) = SWITCH
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

