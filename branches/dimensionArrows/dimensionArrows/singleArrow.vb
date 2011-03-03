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
''' Places a single arrow in one of four styles
''' </summary>
''' <remarks></remarks>
<ComClass(singleArrow.ClassId, singleArrow.InterfaceId, singleArrow.EventsId), _
 ProgId("dimensionArrows.singleArrow")> _
Public NotInheritable Class singleArrow
    Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4840f003-0f8f-4259-8b84-cf3aea6387a0"
    Public Const InterfaceId As String = "57616c2d-6637-465e-82b2-cd7164d07c8d"
    Public Const EventsId As String = "4baea7cd-0b00-4cbf-8ab3-31ef43e668e8"
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
    Private secondPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
    Public prevPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
    Private kill As Boolean = False

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        MyBase.m_category = "OR_DOR"
        MyBase.m_caption = "Single Arrow"
        MyBase.m_message = "Place single arrow"
        MyBase.m_toolTip = "Place single arrow"
        MyBase.m_name = "OR_DOR_singleArrow"
        MyBase.m_enabled = False

        _installationFolder = System.IO.Path.GetDirectoryName(Me.GetType().Assembly.Location)

        Try
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
            MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), Me.GetType().Name + ".cur")
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

    Friend Sub setCursor()
        'Select the cursor image as an indicator of the category of arrow being placed
        Dim cursorName As String = ""
        Select Case _thisArrow.style
            Case arrowStyles.Straight
                cursorname = "SingleStraight.cur"
            Case arrowStyles.Leader
                cursorname = "SingleLeader.cur"
            Case arrowStyles.Zigzag
                cursorname = "SingleZigzag.cur"
            Case arrowStyles.Freeform
                cursorname = "SingleFreeform.cur"
        End Select
        MyBase.m_cursor = New System.Windows.Forms.Cursor( _
            Me.GetType(), cursorName)
    End Sub

    Private Sub onStartEditing()
        MyBase.m_enabled = True
        secondPoint = Nothing
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
            'Get the default arrow category and style from the registry
            _thisArrow.category = GetSetting("OR_DOR_dimensionArrows", "default", "category", _
                arrowCategories.SingleArrow)
            _thisArrow.style = GetSetting("OR_DOR_dimensionArrows", "default", "style", _
                arrowStyles.Straight)
            setCursor()
            _angleIsSet = True
        End If
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
      ByVal X As Integer, ByVal Y As Integer)
        If Button = EsriMouseButtons.Left Then
            Dim theDisplayTransformation As ESRI.ArcGIS.Display.IDisplayTransformation
            theDisplayTransformation = _app.Display.DisplayTransformation

            If _pointNumber = 1 Then
                _firstPoint = theDisplayTransformation.ToMapPoint(X, Y)
                saveSubtype()
                Dim arrowSubtype As Long = findArrowSubtype()
                If arrowSubtype = Nothing Then Exit Sub
                setArrowSubtype(arrowSubtype)
                setLineFeedback(getSnapPoint(getDataFrameCoords(X, Y)))
                _pointNumber = 2
            ElseIf _pointNumber = 2 And _
                _thisArrow.style = arrowStyles.Leader Then
                secondPoint = theDisplayTransformation.ToMapPoint(X, Y)
                showLineFeedback(getSnapPoint(getDataFrameCoords(X, Y)), secondPoint)
                _pointNumber = 3
            ElseIf _thisArrow.style = arrowStyles.Freeform Then
                ReDim Preserve _freeformPoints(_pointNumber - 1)
                _freeformPoints(_pointNumber - 1) = theDisplayTransformation.ToMapPoint(X, Y)
                showLineFeedback(getDataFrameCoords(X, Y))
                _pointNumber = _pointNumber + 1
            Else
                placeArrows(_lastPoint, secondPoint)
                resetSubtype()
                _pointNumber = 1
                secondPoint = Nothing
                prevPoint.PutCoords(0, 0)
            End If
        End If
    End Sub

    Public Overrides Sub OnDblClick()
        If _thisArrow.style = arrowStyles.Freeform Then
            'remove the last point that was created with the mouse down event 
            ' that was triggered when double-clicking
            _pointNumber = _pointNumber - 1
            placeFreeformArrow()
            resetSubtype()
        End If
    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
      ByVal X As Integer, ByVal Y As Integer)
        setCursor()
        If _pointNumber = 1 Then
            Exit Sub
        End If

        Dim geoPoint As IPoint = getDataFrameCoords(X, Y)

        getSnapPoint(geoPoint)

        _editor.InvertAgent(prevPoint, 0)
        _editor.InvertAgent(geoPoint, 0)

        If _thisArrow.style = arrowStyles.Freeform Then
            If _pointNumber > 1 Then
                showLineFeedback(geoPoint)
            End If
        ElseIf _pointNumber = 2 Then
            '_angleIsSet = False
            showLineFeedback(geoPoint)
        ElseIf _pointNumber = 3 Then
            _arrowAngle = 0
            showLineFeedback(geoPoint, secondPoint)
        ElseIf _pointNumber > 3 Then
            showLineFeedback(geoPoint)
        End If
        prevPoint = geoPoint
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
        Select Case _thisArrow.style
            Case arrowStyles.Straight
                ReDim _menuitems(5)
                _menuitems(0) = STYLE_LEADER
                _menuitems(1) = STYLE_ZIGZAG
                _menuitems(2) = STYLE_FREEFORM
                _menuitems(3) = SWITCH
                _menuitems(4) = CANCEL
                _menuitems(5) = HELP
            Case arrowStyles.Freeform
                ReDim _menuitems(6)
                _menuitems(0) = FINISH
                _menuitems(1) = STYLE_STRAIGHT
                _menuitems(2) = STYLE_LEADER
                _menuitems(3) = STYLE_ZIGZAG
                _menuitems(4) = SWITCH
                _menuitems(5) = CANCEL
                _menuitems(6) = HELP
            Case arrowStyles.Leader
                ReDim _menuitems(6)
                _menuitems(0) = STYLE_STRAIGHT
                _menuitems(1) = STYLE_ZIGZAG
                _menuitems(2) = STYLE_FREEFORM
                _menuitems(3) = UNLOCK
                _menuitems(4) = SWITCH
                _menuitems(5) = CANCEL
                _menuitems(6) = HELP
            Case arrowStyles.Zigzag
                ReDim _menuitems(13)
                _menuitems(0) = STYLE_STRAIGHT
                _menuitems(1) = STYLE_LEADER
                _menuitems(2) = STYLE_FREEFORM
                _menuitems(3) = TOPOINT
                _menuitems(4) = TOEND
                _menuitems(5) = NARROWER
                _menuitems(6) = WIDER
                _menuitems(7) = CURVELESS
                _menuitems(8) = CURVEMORE
                _menuitems(9) = FLIP
                _menuitems(10) = SWITCH
                _menuitems(11) = CANCEL
                _menuitems(12) = SAVEDEFAULT
                _menuitems(13) = HELP
        End Select

        commandBar.Popup()
    End Sub

End Class

