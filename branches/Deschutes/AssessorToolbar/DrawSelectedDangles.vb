Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports System.Windows.Forms
Imports ESRI.ArcGIS.EditorExt

<ComClass(DrawSelectedDangles.ClassId, DrawSelectedDangles.InterfaceId, DrawSelectedDangles.EventsId), _
 ProgId("AssessorToolbar.DrawSelectedDangles")> _
Public NotInheritable Class DrawSelectedDangles
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "a2562442-7ed6-4caa-882d-72c6ccba50d3"
    Public Const InterfaceId As String = "39ac65bb-d82d-42b4-b77f-346682cd19e0"
    Public Const EventsId As String = "d4f549a4-5fbe-4dd1-aac2-2a88ca4810c2"
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
    Private _editor As IEditor
    Private WithEvents _editorEvents As Editor
    Private WithEvents _activeViewEvents As Map

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "AssessorToolbar"  'localizable text 
        MyBase.m_caption = "DrawSelectedDangles"   'localizable text 
        MyBase.m_message = "Draws Dangle Nodes for Selected Features."   'localizable text 
        MyBase.m_toolTip = "Draws Dangle Nodes for Selected Features." 'localizable text 
        MyBase.m_name = MyBase.m_category & "_DrawSelectedDangles"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            'TODO: change bitmap name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try


    End Sub


    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then

            _application = CType(hook, IApplication)
            _editor = _application.FindExtensionByName("ESRI Object Editor")
            _editorEvents = _editor
            Dim pMxDoc As IMxDocument = _application.Document
            _activeViewEvents = pMxDoc.FocusMap

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

        'Dim theMxDoc As IMxDocument = _application.Document
        'Dim theActiveView As IActiveView = theMxDoc.FocusMap

        '_buttonChecked = Not _buttonChecked
        'If _buttonChecked Then
        '    _activeViewEvents = theMxDoc.FocusMap
        'End If
        'theActiveView.Refresh()

        Call DrawNodes()

    End Sub


    ''' <summary>
    ''' Called by ArcMap once per second to check if the command is enabled.
    ''' </summary>
    ''' <remarks>WARNING: Do not put computation-intensive code here.</remarks>
    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Dim canEnable As Boolean
            canEnable = _editor.EditState
            Return canEnable
        End Get
    End Property


    Public Overrides ReadOnly Property Checked() As Boolean
        Get
            Return _buttonChecked
        End Get
    End Property



    Private Sub DrawNodes()

        '-- Get the document, map, and activeview
        Dim theMxDoc As IMxDocument = _application.Document
        Dim theMap As IMap = theMxDoc.FocusMap
        If theMap.LayerCount = 0 Then Exit Sub
        Dim theActiveView As IActiveView = theMap

        '-- Set up the topology stuff and clear the classes
        Dim pTopologyExtension As ITopologyExtension = _application.FindExtensionByName("ESRI Topology Extension")
        If pTopologyExtension Is Nothing Then Exit Sub
        Dim pMapTopology As IMapTopology = pTopologyExtension.MapTopology
        pMapTopology.ClearClasses()

        '-- Loop through the layers in the TOC and add visible polyline features to the map topology
        Dim pEnum As IEnumLayer = theMap.Layers
        Dim pFLayer As IFeatureLayer
        Dim pLayer As ILayer = pEnum.Next
        Do Until pLayer Is Nothing
            If TypeOf pLayer Is IFeatureLayer And pLayer.Valid And pLayer.Visible Then
                pFLayer = pLayer
                If pFLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                    pMapTopology.AddClass(pFLayer.FeatureClass)
                End If
            End If
            pLayer = pEnum.Next
        Loop

        '-- Setup status bar
        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)
        Dim pStatusBar As ESRI.ArcGIS.esriSystem.IStatusBar = _application.StatusBar
        Dim pProgbar As ESRI.ArcGIS.esriSystem.IStepProgressor = pStatusBar.ProgressBar
        pStatusBar.Message(0) = "Finding nodes... please wait."

        '-- Build the maptopology cache
        Dim theActiveViewEnv As IEnvelope = theActiveView.Extent
        theActiveViewEnv.Expand(1.05, 10.5, True)
        If Not Contains(pMapTopology.Cache.BuildExtent, theActiveViewEnv) Then
            pMapTopology.Cache.Build(theActiveViewEnv, True)
        End If

        '-- Build the maptopology cache
        'If Not Contains(pMapTopology.Cache.BuildExtent, theActiveView.Extent) Then
        ' pMapTopology.Cache.Build(theActiveView.Extent, True)
        ' End If

        If pMapTopology.Cache.Nodes.Count > 0 Then
            pStatusBar.ShowProgressBar("Drawing nodes...", 0, pMapTopology.Cache.Nodes.Count, 1, True)

            '-- Loop through the topology nodes and display them on the map...
            Dim pEnumNode As IEnumTopologyNode
            pEnumNode = pMapTopology.Cache.Nodes
            Dim pNode As ITopologyNode = pEnumNode.Next
            With theActiveView.ScreenDisplay
                .StartDrawing(theActiveView.ScreenDisplay.hDC, 0)
                Do Until pNode Is Nothing

                    If pNode.Degree = 1 Then
                        .SetSymbol(MakeMarkerSym(Color.Red))
                        .DrawPoint(pNode.Geometry)
                    End If

                    'Select Case pNode.Degree
                    '    Case 1
                    '        .SetSymbol(MakeMarkerSym(Color.Red))
                    '    Case 2
                    '        .SetSymbol(MakeMarkerSym(Color.Blue))
                    '    Case 3
                    '        .SetSymbol(MakeMarkerSym(Color.Green))
                    '    Case Else
                    '        .SetSymbol(MakeMarkerSym(Color.Yellow))
                    'End Select
                    '.DrawPoint(pNode.Geometry)

                    pStatusBar.StepProgressBar()
                    pNode = pEnumNode.Next
                Loop
                .FinishDrawing()
            End With

        End If

        pStatusBar.HideProgressBar()
        pMapTopology.ClearClasses()


        '    If _editor.EditState = esriEditState.esriStateEditing Then

        '        Dim theMxDoc As IMxDocument = _application.Document
        '        Dim theActiveView As IActiveView = theMxDoc.FocusMap
        '        Dim theScreenDisplay As IScreenDisplay = theActiveView.ScreenDisplay

        '        Dim theColor As System.Drawing.Color = Drawing.Color.Yellow
        '        Dim theRGBColor As IRgbColor = New RgbColor
        '        With theRGBColor
        '            .Red = theColor.R
        '            .Green = theColor.G
        '            .Blue = theColor.B
        '            .Transparency = 1
        '        End With

        '        Dim theMarkerSym As ISimpleMarkerSymbol = New SimpleMarkerSymbol
        '        With theMarkerSym
        '            .Style = esriSimpleMarkerStyle.esriSMSCircle
        '            .Outline = True
        '            .Color = theRGBColor
        '            .Size = 8
        '        End With

        '        Dim theSpatialFilter As ISpatialFilter = New SpatialFilter
        '        Dim thisPolyCurve As IPolycurve2
        '        Dim thisPoint As IPoint
        '        Dim thisRelOpFrom As IRelationalOperator
        '        Dim thisRelOpTo As IRelationalOperator
        '        Dim thisFeature As IFeature
        '        Dim thisPolyCV As IPolycurve2
        '        Dim n As Integer


        '        '-- Start Shad

        '        Dim theFeatureLayer As IFeatureLayer
        '        theFeatureLayer = theMxDoc.FocusMap.Layer(0)
        '        Dim theFeatureClass As IFeatureClass = theFeatureLayer.FeatureClass

        '        Dim theMapEnvelope As IEnvelope = theActiveView.Extent.Envelope

        '        Dim theMapSpatialFilter As ISpatialFilter
        '        theMapSpatialFilter = New SpatialFilter
        '        With theMapSpatialFilter
        '            .Geometry = theMapEnvelope
        '            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects 'esriSpatialRelIntersects
        '        End With

        '        theMapSpatialFilter.GeometryField = theFeatureClass.ShapeFieldName


        '        Dim theFeatureCursor As IFeatureCursor = theFeatureLayer.FeatureClass.Search(theMapSpatialFilter, False)

        '        Dim thisCursorFeature As IFeature = theFeatureCursor.NextFeature

        '        Dim x As Integer = 0

        '        Do While Not thisCursorFeature Is Nothing

        '            thisPolyCurve = thisCursorFeature.Shape
        '            thisPoint = thisPolyCurve.FromPoint

        '            x += 1
        '            For i As Integer = 1 To 2  'for 2 points of from and to

        '                With theSpatialFilter
        '                    .Geometry = thisPoint
        '                    .GeometryField = "SHAPE"
        '                    .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
        '                    .OutputSpatialReference(theFeatureClass.ShapeFieldName) = thisPoint.SpatialReference
        '                End With

        '                Dim pFeatureCursor As IFeatureCursor = theFeatureClass.Search(theSpatialFilter, False) 'NOT WORK for curve end
        '                thisFeature = pFeatureCursor.NextFeature
        '                n = 0

        '                Do While Not thisFeature Is Nothing
        '                    thisPolyCV = thisFeature.Shape
        '                    thisRelOpFrom = thisPolyCV.FromPoint
        '                    thisRelOpTo = thisPolyCV.ToPoint

        '                    'Check for points with the same coordinates within the precision of the dataset
        '                    If thisRelOpTo.Equals(thisPoint) Then n = n + 1
        '                    If thisRelOpFrom.Equals(thisPoint) Then n = n + 1
        '                    If n > 1 Then Exit Do

        '                    thisFeature = pFeatureCursor.NextFeature
        '                Loop

        '                If n = 1 Then
        '                    With theScreenDisplay
        '                        .StartDrawing(theScreenDisplay.hDC, -1)
        '                        .SetSymbol(theMarkerSym)
        '                        .DrawPoint(thisPoint)
        '                        .FinishDrawing()
        '                    End With
        '                End If

        '                thisPoint = thisPolyCurve.ToPoint
        '            Next i


        '            thisCursorFeature = theFeatureCursor.NextFeature
        '        Loop

        '        '-- End Shad

        '    Else
        '        MessageBox.Show("Must be editing to display dangle nodes", "Start Edit Session", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    End If

    End Sub



    'Private Sub _pActiveViewEvents_AfterDraw(ByVal Display As IDisplay, ByVal phase As esriViewDrawPhase) Handles _activeViewEvents.AfterDraw

    'End Sub

    'Private Sub _pEditorEvents_OnCurrentLayerChanged() Handles _editorEvents.OnCurrentLayerChanged
    '    RefreshMap()
    'End Sub

    'Private Sub _pEditorEvents_OnSelectionChanged() Handles _editorEvents.OnSelectionChanged
    '    RefreshMap()
    'End Sub



    Function Contains(ByVal pRelOp As IRelationalOperator, _
    ByVal pGeom As IGeometry) As Boolean
        Contains = pRelOp.Contains(pGeom)
    End Function

    Function MakeMarkerSym(ByVal theColor As Color, Optional ByVal dSize As Double = 8) As ISimpleMarkerSymbol

        Dim theRGBColor As IRgbColor = New RgbColor
        With theRGBColor
            .Red = theColor.R
            .Green = theColor.G
            .Blue = theColor.B
            .Transparency = 1
        End With

        MakeMarkerSym = New SimpleMarkerSymbol
        MakeMarkerSym.Color = theRGBColor
        MakeMarkerSym.Size = dSize
        MakeMarkerSym.Style = esriSimpleMarkerStyle.esriSMSCircle

    End Function

    Private Sub RefreshMap()
        Dim theMxDoc As IMxDocument
        theMxDoc = _application.Document
        Dim theActiveView As IActiveView
        theActiveView = theMxDoc.ActiveView
        theActiveView.Refresh()
    End Sub

End Class



