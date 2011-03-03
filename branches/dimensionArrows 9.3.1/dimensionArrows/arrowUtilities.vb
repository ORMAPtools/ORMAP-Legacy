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
Imports System.Xml
Imports System.IO

Module arrowUtilities

#Region "Module Variables"

    ''' <summary>
    ''' Enumeration of ESRI mouse button constant values.
    ''' </summary>
    Public Enum EsriMouseButtons
        None = 0
        Left = 1
        Right = 2
        Middle = 4
    End Enum

    ''' <summary>
    ''' Enumeration of arrow types.
    ''' </summary>
    Public Enum arrowCategories
        Straight = 0
        LandHook = 1
        NoDashes = 2
        OneDash = 3
        TwoDashes = 4
        ThreeDashes = 5
        FourDashes = 6
        SingleArrow = 7
    End Enum

    ''' <summary>
    ''' Enumeration of single arrow styles
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum arrowStyles
        Straight = 0
        Leader = 1
        Zigzag = 2
        Freeform = 3
    End Enum

    Structure arrowType
        Dim category As Integer
        Dim style As Integer
    End Structure

#Region "constants"
    Friend Const SEPARATOR = "-"
    Friend Const SHORTER = "Shorter - Ctrl+Down"
    Friend Const LONGER = "Longer - Ctrl+Up"
    Friend Const FLIP = "Flip Arrows - F"
    Friend Const UNLOCK = "Unlock/Lock Angle - U"
    Friend Const SWITCH = "Switch Arrowheads - S"
    Friend Const CANCEL = "Cancel - Esc"
    Friend Const SCALE10 = "10 Scale"
    Friend Const SCALE20 = "20 Scale"
    Friend Const SCALE30 = "30 Scale"
    Friend Const SCALE40 = "40 Scale"
    Friend Const SCALE50 = "50 Scale"
    Friend Const SCALE100 = "100 Scale"
    Friend Const SCALE200 = "200 Scale"
    Friend Const SCALE400 = "400 Scale"
    Friend Const SCALE800 = "800 Scale"
    Friend Const SCALE1000 = "1000 Scale"
    Friend Const SCALE2000 = "2000 Scale"
    Friend Const NARROWER = "Narrower"
    Friend Const WIDER = "Wider"
    Friend Const TOPOINT = "Slide toward point"
    Friend Const TOEND = "Slide toward end"
    Friend Const CURVELESS = "Less curve"
    Friend Const CURVEMORE = "More curve"
    Friend Const STYLE_STRAIGHT = "Straight arrow style"
    Friend Const STYLE_LEADER = "Leader arrow style"
    Friend Const STYLE_ZIGZAG = "Zigzag arrow style"
    Friend Const STYLE_FREEFORM = "Freeform arrow style"
    Friend Const SAVEDEFAULT = "Save as default zigzag arrow"
    Friend Const HELP = "Help"
    Friend Const FINISH = "Finish Arrow"

#End Region

    Friend _app As IMxApplication
    Friend _editor As IEditor2
    Friend _editEvents As IEditEvents_Event
    Friend _firstPoint As ESRI.ArcGIS.Geometry.IPoint
    Friend _flipArrows As Boolean = False
    Friend _arrowScale As Double
    Friend _arrowAngle As Double
    Friend _angleIsSet As Boolean
    Friend _featureClass As IFeatureLayer
    Friend _subtype As Long
    Friend _lastPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
    Friend _pointNumber As Integer = 1
    Friend _menuitems() As String
    Friend _arrowheadIsSwitched As Boolean = False
    Friend _arrowOffset As Double
    Friend _installationFolder As String
    Friend _zigzagWidth As Double = _
        GetSetting("OR_DOR_dimensionArrows", "default", "zigzagWidth", 5.0)
    Friend _zigzagCurve As Double = _
        GetSetting("OR_DOR_dimensionArrows", "default", "zigzagCurve", 5.0)
    Friend _zigzagPosition As Double = _
        GetSetting("OR_DOR_dimensionArrows", "default", "zigzagPosition", 10.0)
    Friend _thisArrow As arrowType
    Friend _freeformPoints() As IPoint

#End Region

    ''' <summary>
    ''' Converts mouse screen coordinates to data frame coordinates
    ''' </summary>
    ''' <param name="X">The mouse X position</param>
    ''' <param name="Y">The mouse Y position</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function getDataFrameCoords(ByVal X As Integer, ByVal Y As Integer) As IPoint
        Dim displayTransformation As ESRI.ArcGIS.Display.IDisplayTransformation
        displayTransformation = _app.Display.DisplayTransformation

        Return displayTransformation.ToMapPoint(X, Y)
    End Function

    ''' <summary>
    ''' If snapping is on, finds the snap point
    ''' </summary>
    ''' <param name="point"></param>
    ''' <returns></returns>
    ''' <remarks>If there is no snap point the original point is returned</remarks>
    Function getSnapPoint(ByVal point As IPoint) As IPoint
        Dim snapEnv As ISnapEnvironment = CType(_editor, ISnapEnvironment)
        snapEnv.SnapPoint(point)
        Return point
    End Function

    ''' <summary>
    ''' Converts a geographic to a screen point
    ''' </summary>
    ''' <param name="geoPoint"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getScreenCoords(ByVal geoPoint As IPoint) As System.Drawing.Point
        Dim mxDoc As IMxDocument = _app.document
        Dim activeView As IActiveView = mxDoc.ActivatedView
        Dim screenDisplay As IScreenDisplay = activeView.ScreenDisplay
        Dim displayTransformation As ESRI.ArcGIS.Display.IDisplayTransformation
        displayTransformation = screenDisplay.DisplayTransformation
        Dim screenPoint As System.Drawing.Point

        displayTransformation.FromMapPoint(geoPoint, screenPoint.X, screenPoint.Y)
        screenPoint.X = screenPoint.X + displayTransformation.DeviceFrame.left
        screenPoint.Y = screenPoint.Y + displayTransformation.DeviceFrame.top
        Return screenPoint
    End Function

    ''' <summary>
    ''' Set up the arrow display after the first mouse click
    ''' </summary>
    ''' <param name="point">Mouse position in data frame coordinates</param>
    ''' <remarks></remarks>
    Friend Sub setLineFeedback(ByVal point As IPoint)
        _firstPoint = point
        If _thisArrow.style = arrowStyles.Freeform Then
            ReDim _freeformPoints(0)
            _freeformPoints(0) = point
        End If
        If _thisArrow.category <> arrowCategories.SingleArrow Then
            SetScale(_firstPoint)
            _arrowAngle = perpendicularAngle()
            If _arrowAngle = Nothing Then
                _angleIsSet = False
            Else
                _angleIsSet = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' Places the arrows after the last mouse click
    ''' </summary>
    ''' <param name="endPoint">Mouse position in data frame coordinates</param>
    ''' <param name="secondPoint">Option second mouse point for arrows that 
    ''' require three points</param>
    ''' <remarks></remarks>
    Friend Sub placeArrows(ByVal endPoint As IPoint, _
        Optional ByVal secondPoint As IPoint = Nothing)

        Dim editorUID As New UID
        Dim theEditSketch As IEditSketch2

        editorUID.Value = "esriEditor.Editor"
        theEditSketch = _app.FindExtensionByCLSID(editorUID)

        theEditSketch.GeometryType = esriGeometryType.esriGeometryPolyline

        Dim count As Integer
        Dim theEditTask As IEditTask
        Dim taskName As String
        taskName = _editor.CurrentTask.Name
        For count = 0 To _editor.TaskCount - 1
            theEditTask = _editor.Task(count)
            If theEditTask.Name = "Create New Feature" Then
                _editor.CurrentTask = theEditTask
                Exit For
            End If
        Next count

        If _thisArrow.category = arrowCategories.SingleArrow Then
            Dim polyline As IPolyline = getSingleArrowGeometry(endPoint, secondPoint)
            theEditSketch.Geometry = polyline
            theEditSketch.FinishSketch()
        Else
            Dim polyLineArray As IPolylineArray = getArrowGeometry(endPoint)
            For count = 0 To polyLineArray.Count - 1
                theEditSketch.Geometry = polyLineArray.Element(count)
                theEditSketch.FinishSketch()
            Next
        End If

        For count = 0 To _editor.TaskCount - 1
            theEditTask = _editor.Task(count)
            If theEditTask.Name = taskName Then
                _editor.CurrentTask = theEditTask
                Exit For
            End If
        Next count

        Dim mxDoc As IMxDocument = _app.document
        clearAll()
        mxDoc.FocusMap.ClearSelection()

    End Sub

    ''' <summary>
    ''' Finds the angle perpendicular to the selected line for straight opposing arrows
    ''' </summary>
    ''' <returns>the perpendicular angle</returns>
    ''' <remarks></remarks>
    Friend Function perpendicularAngle() As Double
        Try
            Dim pGeometry As ESRI.ArcGIS.Geometry.IGeometry
            Dim pEnv As ESRI.ArcGIS.Geometry.IEnvelope
            Dim pSpatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter
            Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
            Dim pEditLayers As ESRI.ArcGIS.Editor.IEditLayers
            Dim pFCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
            Dim pFeature As ESRI.ArcGIS.Geodatabase.IFeature
            Dim pMap As ESRI.ArcGIS.Carto.IMap
            Dim ShapeFieldName As String

            perpendicularAngle = Nothing

            'Get the Map from the editor
            pEditLayers = TryCast(_editor, ESRI.ArcGIS.Editor.IEditLayers)
            pMap = TryCast(_editor.Map, ESRI.ArcGIS.Carto.IMap)

            'Pass point to CreateSearchShape which creates a geometry around the point
            'The larger geometry is an envelope and will give us better search results
            'The click therefore doesn't have to be exactly on the feature
            pGeometry = _editor.CreateSearchShape(_firstPoint)
            pEnv = TryCast(pGeometry, ESRI.ArcGIS.Geometry.IEnvelope)

            'Create a new spatial filter and use the new envelope as the geometry
            pSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
            pSpatialFilter.Geometry = pEnv
            ShapeFieldName = pEditLayers.CurrentLayer.FeatureClass.ShapeFieldName
            pSpatialFilter.OutputSpatialReference(ShapeFieldName) = pMap.SpatialReference
            pSpatialFilter.GeometryField = _
             pEditLayers.CurrentLayer.FeatureClass.ShapeFieldName
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

            Dim enumLayers As IEnumLayer
            Dim eachLayer As IFeatureLayer
            Dim pUID As UID = New UIDClass()

            pUID.Value = "{40A9E885-5533-11d0-98BE-00805F7CED21}" 'only select IfeatureLayer

            'Get all the feature layers from the map
            enumLayers = pMap.Layers(pUID, True)
            enumLayers.Reset()
            eachLayer = enumLayers.Next

            Dim featureSet As IFeature = New Feature
            Dim editLayer As IEditLayers = DirectCast(_editor, IEditLayers)

            Do Until eachLayer Is Nothing
                'Only search for lines or polygons
                If eachLayer.Valid Then
                    If eachLayer.FeatureClass.ShapeType = _
                     esriGeometryType.esriGeometryPolyline Then
                        pFeatClass = eachLayer.FeatureClass
                        pFCursor = pFeatClass.Search(pSpatialFilter, False) 'Do the search
                        pFeature = pFCursor.NextFeature 'Get the first feature
                        If Not pFeature Is Nothing Then
                            Dim polyLine As IPolyline
                            polyLine = DirectCast(pFeature.Shape, IPolyline)
                            Dim fromPoint As ESRI.ArcGIS.Geometry.Point _
                                = New ESRI.ArcGIS.Geometry.Point
                            Dim toPoint As ESRI.ArcGIS.Geometry.Point _
                                = New ESRI.ArcGIS.Geometry.Point
                            Dim pointDist As Double
                            Dim perpendicularLine As ILine = New Line

                            polyLine.QueryPointAndDistance( _
                                esriSegmentExtension.esriExtendAtFrom, _firstPoint, False, _
                                fromPoint, pointDist, Nothing, False)
                            polyLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, _
                                pointDist, False, 50, perpendicularLine)
                            perpendicularLine.QueryFromPoint(fromPoint)
                            perpendicularLine.QueryToPoint(toPoint)
                            perpendicularAngle = perpendicularLine.Angle
                            Exit Function
                        End If
                    End If
                End If
                eachLayer = enumLayers.Next
            Loop

        Catch ex As Exception
            MsgBox("perpendicularAngle" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
    End Function

    Friend Sub showHelp()
        Dim myproc As System.Diagnostics.Process = New System.Diagnostics.Process
        myproc.EnableRaisingEvents = False
        myproc.StartInfo.FileName = _installationFolder & "\DimensionArrows.chm"
        myproc.Start()
    End Sub

    ''' <summary>
    ''' Finds the map scale based on the underlying MapIndex polygon and saves it in
    ''' the _arrowScale variable
    ''' </summary>
    ''' <param name="point">The first arrow point</param>
    ''' <remarks>If there is no map index a scale of 1 (1"=100') is set</remarks>
    Friend Sub SetScale(ByVal point As IPoint)
        Try
            Dim pGeometry As ESRI.ArcGIS.Geometry.IGeometry
            Dim pEnv As ESRI.ArcGIS.Geometry.IEnvelope
            Dim pSpatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter
            Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
            Dim pEditLayers As ESRI.ArcGIS.Editor.IEditLayers
            Dim pFCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
            Dim pFeature As ESRI.ArcGIS.Geodatabase.IFeature
            Dim pMap As ESRI.ArcGIS.Carto.IMap
            Dim ShapeFieldName As String

            'Get the Map from the editor
            pEditLayers = TryCast(_editor, ESRI.ArcGIS.Editor.IEditLayers)
            pMap = TryCast(_editor.Map, ESRI.ArcGIS.Carto.IMap)

            'Pass point to CreateSearchShape which creates a geometry around the point
            'The larger geometry is an envelope and will give us better search results
            'The click therefore doesn't have to be exactly on the feature
            pGeometry = _editor.CreateSearchShape(point)
            pEnv = TryCast(pGeometry, ESRI.ArcGIS.Geometry.IEnvelope)

            'Create a new spatial filter and use the new envelope as the geometry
            pSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
            pSpatialFilter.Geometry = pEnv
            ShapeFieldName = pEditLayers.CurrentLayer.FeatureClass.ShapeFieldName
            pSpatialFilter.OutputSpatialReference(ShapeFieldName) = pMap.SpatialReference
            pSpatialFilter.GeometryField = pEditLayers.CurrentLayer.FeatureClass.ShapeFieldName
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

            Dim enumLayers As IEnumLayer
            Dim eachLayer As IFeatureLayer
            Dim pUID As UID = New UIDClass()

            pUID.Value = "{40A9E885-5533-11d0-98BE-00805F7CED21}" 'only select IfeatureLayer

            'Get all the feature layers from the map
            enumLayers = pMap.Layers(pUID, True)
            enumLayers.Reset()
            eachLayer = enumLayers.Next

            Do Until eachLayer Is Nothing
                If InStr(UCase(eachLayer.Name), "MAPINDEX") Then
                    'Only search the specified geometry type
                    If eachLayer.FeatureClass.ShapeType = _
                        esriGeometryType.esriGeometryPolygon Then
                        pFeatClass = eachLayer.FeatureClass
                        pFCursor = pFeatClass.Search(pSpatialFilter, False) 'Do the search
                        pFeature = pFCursor.NextFeature 'Get the first feature
                        If Not pFeature Is Nothing Then
                            _arrowScale = _
                                pFeature.Value( _
                                pFeature.Fields.FindField("MapScale")) / 1200
                            Exit Sub
                        End If
                    End If
                End If
                eachLayer = enumLayers.Next
            Loop
            'if there is no map index then set it to 100 scale
            _arrowScale = 1
        Catch ex As Exception
            MsgBox("SetScale" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' Get the geometry of the paired arrows
    ''' </summary>
    ''' <param name="endPoint">Mouse position in data frame coordinates</param>
    ''' <returns>An IPolyLineArray of two elements</returns>
    ''' <remarks></remarks>
    Friend Function getArrowGeometry( _
        ByVal endPoint As IPoint) As IPolylineArray

        getArrowGeometry = Nothing

        If _angleIsSet Then
            endPoint.ConstrainAngle(_arrowAngle, _firstPoint, True)
        End If

        Dim count As Integer

        Dim geoString As String
        geoString = ReadXML(_thisArrow.category)
        Dim segmentCount As Integer
        segmentCount = CInt(Left(geoString, InStr(geoString, ",") - 1))
        Dim segmentString() As String = geoString.Split(",")

        Dim segment1() As ILine
        Dim segment2() As ILine
        Dim missing As Object = Type.Missing

        Dim arrow1 As IPolyline = New Polyline
        Dim arrow2 As IPolyline = New Polyline

        'Draw the left arrow
        Dim path1 As ISegmentCollection = New Polyline
        Dim path2 As ISegmentCollection = New Polyline

        Dim scaleFactor As Double
        If _pointNumber = 3 Then
            scaleFactor = _arrowOffset / segmentString(UBound(segmentString)) / _arrowScale
        Else
            scaleFactor = 1
        End If

        ReDim segment1(segmentCount - 1)
        For count = 0 To segmentCount - 1
            segment1(count) = New ESRI.ArcGIS.Geometry.Line
        Next

        Dim coordStep As Integer = 1
        For count = 0 To segmentCount - 1
            Dim point1 As IPoint = New ESRI.ArcGIS.Geometry.Point
            Dim point2 As IPoint = New ESRI.ArcGIS.Geometry.Point
            point1.X = CDbl(segmentString(coordStep))
            If count = 0 Then
                point1.Y = CDbl(segmentString(coordStep + 1))
            Else
                point1.Y = CDbl(segmentString(coordStep + 1)) * scaleFactor
            End If
            segment1(count).FromPoint = point1
            point2.X = CDbl(segmentString(coordStep + 2))
            point2.Y = CDbl(segmentString(coordStep + 3)) * scaleFactor
            segment1(count).ToPoint = point2
            path1.AddSegment(CType(segment1(count), ISegment), missing, missing)
            coordStep = coordStep + 4
        Next

        'Draw the right arrow
        ReDim segment2(segmentCount - 1)
        For count = 0 To segmentCount - 1
            segment2(count) = New ESRI.ArcGIS.Geometry.Line
        Next

        coordStep = 1
        For count = 0 To segmentCount - 1
            Dim point1 As IPoint = New ESRI.ArcGIS.Geometry.Point
            Dim point2 As IPoint = New ESRI.ArcGIS.Geometry.Point
            point1.X = CDbl(segmentString(coordStep)) * -1
            If count = 0 Then
                point1.Y = CDbl(segmentString(coordStep + 1))
            Else
                point1.Y = CDbl(segmentString(coordStep + 1)) * scaleFactor
            End If
            segment2(count).FromPoint = point1
            point2.X = CDbl(segmentString(coordStep + 2)) * -1
            point2.Y = CDbl(segmentString(coordStep + 3)) * scaleFactor
            If _thisArrow.category = arrowCategories.LandHook And count = segmentCount - 1 Then
                point2.Y = point2.Y * -1
            End If
            segment2(count).ToPoint = point2
            path2.AddSegment(CType(segment2(count), ISegment), missing, missing)
            coordStep = coordStep + 4
        Next

        arrow1 = path1
        arrow2 = path2

        If _arrowheadIsSwitched Then
            arrow1.ReverseOrientation()
            arrow2.ReverseOrientation()
        End If

        If _thisArrow.category = arrowCategories.Straight Then
            If _flipArrows Then
                arrow1 = path2
                arrow2 = path1
            End If
        End If

        'use a new line to get the angle of the vector between the points
        Dim vector As ILine = New Line
        vector.FromPoint = _firstPoint
        vector.ToPoint = endPoint
        Dim vectorAngle As Double = vector.Angle
        vector = Nothing

        'Transform the polylines by moving, rotating and scaling them to the proper positions
        Dim transform As ITransform2D = arrow1
        transform.Move(_firstPoint.X, _firstPoint.Y)
        transform.Rotate(_firstPoint, vectorAngle)
        transform.Scale(_firstPoint, _arrowScale, _arrowScale)

        transform = arrow2
        transform.Move(endPoint.X, endPoint.Y)
        transform.Rotate(endPoint, vectorAngle)
        transform.Scale(endPoint, _arrowScale, _arrowScale)

        Dim arrowArray As IPolylineArray = New PolylineArray
        arrowArray.Add(arrow1)
        arrowArray.Add(arrow2)

        Return arrowArray
    End Function

    ''' <summary>
    ''' Save the current feature class and subtype
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub saveSubtype()
        Dim editLayer As IEditLayers = _editor
        _featureClass = editLayer.CurrentLayer
        _subtype = editLayer.CurrentSubtype
    End Sub

    ''' <summary>
    ''' Set the subtype of the arrows
    ''' </summary>
    ''' <param name="subtypeNumber">The subtype number to set</param>
    ''' <remarks></remarks>
    Friend Sub setArrowSubtype(ByVal subtypeNumber As Long)
        Dim i As Integer
        Dim theMap As ESRI.ArcGIS.Carto.IMap
        Dim theDoc As IMxDocument
        Try
            theDoc = _editor.Parent.Document
            theMap = theDoc.FocusMap

            For i = 0 To theMap.LayerCount - 1
                If LCase(theMap.Layer(i).Name) Like "*cartographic*" Then

                    Dim featLayer As IFeatureLayer
                    featLayer = TryCast(theMap.Layer(i), IFeatureLayer)

                    Dim featClass As IFeatureClass
                    featClass = featLayer.FeatureClass

                    Dim editLayer As IEditLayers = _editor

                    Dim enumSubtypes As IEnumSubtype
                    Dim subtype As ISubtypes
                    subtype = featClass
                    enumSubtypes = subtype.Subtypes

                    Dim subtypeCode As Long

                    enumSubtypes.Next(subtypeCode)
                    Do While subtypeCode
                        If subtypeCode = subtypeNumber Then
                            editLayer.SetCurrentLayer(featLayer, subtypeCode)
                            Exit Sub
                        End If
                        enumSubtypes.Next(subtypeCode)
                    Loop
                End If
            Next
        Catch ex As Exception
            MsgBox("setArrowSubtype - " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Reset the feature class and subtype after placing the arrow
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub resetSubtype()
        Dim editLayer As IEditLayers = _editor
        editLayer.SetCurrentLayer(_featureClass, _subtype)
    End Sub

    ''' <summary>
    ''' Show the arrows while moving the mouse
    ''' </summary>
    ''' <param name="endPoint">Mouse position in data frame coordinates</param>
    ''' <param name="secondPoint">Option second mouse point for arrows that 
    ''' require three points</param>
    ''' <remarks></remarks>
    Friend Sub showLineFeedback(ByVal endPoint As IPoint, _
        Optional ByVal secondPoint As IPoint = Nothing)
        Dim mxDoc As IMxDocument = _app.document

        _lastPoint = endPoint

        Dim graphicsContainer As IGraphicsContainer = mxDoc.ActiveView.GraphicsContainer
        graphicsContainer.DeleteAllElements()

        If _thisArrow.category = arrowCategories.SingleArrow Then
            Dim polyLine As IPolyline = getSingleArrowGeometry(endPoint, secondPoint)
            drawArrowImage(polyLine)
        Else
            Dim polyLineArray As IPolylineArray = getArrowGeometry(endPoint)
            drawArrowImage(polyLineArray.Element(0))
            drawArrowImage(polyLineArray.Element(1))
        End If
    End Sub

    ''' <summary>
    ''' Get the geometry for a single arrow
    ''' </summary>
    ''' <param name="endPoint">Mouse position in data frame coordinates</param>
    ''' <param name="secondPoint">Optional second mouse point for arrows that 
    ''' require three points</param>
    ''' <returns>The arrow geometry as an IPolyLine</returns>
    ''' <remarks></remarks>
    Friend Function getSingleArrowGeometry( _
        ByVal endPoint As IPoint, ByVal secondPoint As IPoint) As IPolyline

        getSingleArrowGeometry = Nothing
        Dim count As Integer
        Dim segpoints(3) As IPoint

        If Not _thisArrow.style = arrowStyles.Freeform Then
            If _angleIsSet And _pointNumber = 3 Then
                endPoint.ConstrainAngle(0, secondPoint, True)
            End If

            If secondPoint Is Nothing Then
                secondPoint = endPoint
            End If

            For count = 0 To 3
                segpoints(count) = New ESRI.ArcGIS.Geometry.Point
            Next
        End If

        Dim offset As Integer = 1
        If _flipArrows Then
            offset = -1
        End If

        Dim polyline As IPolyline = New Polyline

        Select Case _thisArrow.style
            Case arrowStyles.Straight
                Dim path As ISegmentCollection = New Polyline
                Dim firstSegment As ISegment = New ESRI.ArcGIS.Geometry.Line
                firstSegment.FromPoint = _firstPoint
                firstSegment.ToPoint = secondPoint
                path.AddSegment(firstSegment)
                polyline = path
            Case arrowStyles.Leader
                Dim path As ISegmentCollection = New Polyline
                Dim firstSegment As ISegment = New ESRI.ArcGIS.Geometry.Line
                Dim lastSegment As ISegment = New ESRI.ArcGIS.Geometry.Line
                firstSegment.FromPoint = _firstPoint
                firstSegment.ToPoint = secondPoint
                lastSegment.FromPoint = secondPoint
                lastSegment.ToPoint = endPoint
                path.AddSegment(firstSegment)
                path.AddSegment(lastSegment)
                polyline = path
            Case arrowStyles.Zigzag
                segpoints(0).PutCoords(0, 0)
                segpoints(1).PutCoords(_zigzagPosition, 0)
                segpoints(2).PutCoords(_zigzagPosition, _zigzagWidth * offset)
                segpoints(3).PutCoords(20, _zigzagWidth * offset)

                Dim firstSegment As ISegment = New ESRI.ArcGIS.Geometry.Line
                Dim lastSegment As ISegment = New ESRI.ArcGIS.Geometry.Line
                firstSegment.FromPoint = segpoints(0)
                firstSegment.ToPoint = segpoints(1)
                lastSegment.FromPoint = segpoints(2)
                lastSegment.ToPoint = segpoints(3)

                Dim bSpline As IBezierCurveGEN = New BezierCurve
                Dim points(3) As IPoint
                For count = 0 To 3
                    points(count) = New ESRI.ArcGIS.Geometry.Point
                Next

                points(0) = segpoints(1)
                points(1).PutCoords(segpoints(1).X + _zigzagCurve, segpoints(1).Y)
                points(2).PutCoords(segpoints(2).X - _zigzagCurve, segpoints(2).Y)
                points(3) = segpoints(2)

                bSpline.PutCoords(points)

                Dim path As ISegmentCollection = New Polyline
                path.AddSegment(firstSegment)
                path.AddSegment(bSpline)
                path.AddSegment(lastSegment)

                polyline = path

                'use a new line to get the angle of the vector between the points
                Dim polyVector As ILine = New Line
                polyVector.FromPoint = polyline.FromPoint
                polyVector.ToPoint = polyline.ToPoint
                Dim polyAngle As Double
                polyAngle = polyVector.Angle

                Dim vector As ILine = New Line
                vector.FromPoint = _firstPoint
                vector.ToPoint = endPoint
                Dim vectorAngle As Double = vector.Angle - polyAngle

                Dim scale As Double
                scale = vector.Length / polyVector.Length

                'Transform the polylines by moving, rotating and scaling them to the _
                'proper positions
                Dim transform As ITransform2D = polyline
                transform.Move(_firstPoint.X, _firstPoint.Y)
                transform.Rotate(_firstPoint, vectorAngle)
                transform.Scale(_firstPoint, scale, scale)
            Case arrowStyles.Freeform
                Dim points(_pointNumber - 1) As IPoint
                ReDim Preserve _freeformPoints(_pointNumber - 1)

                For count = 0 To UBound(points)
                    points(count) = _freeformPoints(count)
                Next

                points(UBound(points)) = endPoint

                Dim bezierFeedback As INewBezierCurveFeedback
                Try
                    bezierFeedback = New NewBezierCurveFeedback
                    bezierFeedback.Start(points(0))
                    For count = 1 To UBound(points)
                        bezierFeedback.AddPoint(points(count))
                    Next
                    'bezierFeedback.AddPoint(points(UBound(points)))
                    Dim geoMetry As IGeometry
                    geoMetry = bezierFeedback.Stop
                    polyline = CType(geoMetry, Polyline)
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
        End Select

        If _arrowheadIsSwitched Then
            polyline.ReverseOrientation()
        End If

        Return polyline
    End Function

    ''' <summary>
    ''' Finds the arrow subtype based on the current edit layer
    ''' </summary>
    ''' <returns>The subtype number</returns>
    ''' <remarks></remarks>
    Friend Function findArrowSubtype() As Integer
        Dim editLayer As IEditLayers = _editor
        Dim featureLayer As IFeatureLayer2 = editLayer.CurrentLayer
        Dim subtype As Long = editLayer.CurrentSubtype
        Dim featureClass As IFeatureClass = featureLayer.FeatureClass
        Dim sName As String
        Dim subtypeList As ISubtypes = CType(featureClass, ISubtypes)

        If featureClass.AliasName.ToLower Like "*cartographic*" Then
            Return subtype
        ElseIf _thisArrow.category = arrowCategories.SingleArrow Then
            If featureClass.AliasName.ToLower Like "*anno*" Then
                sName = subtypeList.SubtypeName(subtype).ToLower

                Select Case True
                    Case sName Like "*bearing*"
                        Return 134
                    Case sName Like "*block*"
                        Return 137
                    Case sName Like "*dlc*", sName Like "*d.l.c.*"
                        Return 147
                    Case sName Like "*easement*", sName Like "*landmark*", _
                        sName Like "*meanander*", sName Like "*hydro*", _
                        sName Like "*ref*", sName Like "*water*"
                        Return 136
                    Case sName Like "*index*"
                        Return 162
                    Case sName Like "*sub*"
                        Return 141
                    Case sName Like "*code*"
                        Return 154
                    Case sName Like "taxlot", sName Like "*tax lot*"
                        Return 137
                    Case Else
                        Return 136
                End Select
            End If
        ElseIf _thisArrow.category = arrowCategories.Straight Then
            If featureClass.AliasName.ToLower Like "*anno*" Then
                sName = subtypeList.SubtypeName(subtype).ToLower

                If sName Like "*easement*" Or sName Like "*landmark*" Or _
                    sName Like "*meanander*" Or sName Like "*hydro*" Or _
                    sName Like "*ref*" Or sName Like "*water*" Then
                    Return 136
                Else
                    Return 134
                End If
            End If
        Else
            MsgBox("Arrow subtype could not be determined." & vbCrLf & _
                "Please select an arrow subtype", MsgBoxStyle.OkOnly)
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' Draws the arrow image to the screen
    ''' </summary>
    ''' <param name="arrow">Input polyline</param>
    ''' <remarks></remarks>
    Friend Sub drawArrowImage(ByVal arrow As IPolyline)
        Dim mxDoc As IMxDocument = _app.document
        Dim graphicsContainer As IGraphicsContainer = mxDoc.ActiveView.GraphicsContainer
        Dim symbol As ISymbol

        Dim editLayer As IEditLayers = _editor

        Dim uvr As IUniqueValueRenderer
        Dim geoLayer As IGeoFeatureLayer = editLayer.CurrentLayer
        uvr = geoLayer.Renderer

        symbol = uvr.DefaultSymbol

        Dim count As Integer
        For count = 0 To uvr.ValueCount
            If uvr.Value(count) = editLayer.CurrentSubtype Then
                symbol = CType(uvr.Symbol(uvr.Value(count)), ISymbol)
                Exit For
            End If
        Next

        Dim lineElement As ILineElement = New LineElement
        lineElement.Symbol = symbol
        Dim element As ESRI.ArcGIS.Carto.IElement
        element = CType(lineElement, IElement)
        element.Geometry = CType(arrow, IGeometry)

        graphicsContainer.AddElement(element, 0)
        mxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
    End Sub

    ''' <summary>
    ''' KeyDown events 
    ''' </summary>
    ''' <param name="keyCode">the key code</param>
    ''' <param name="shift">the control/shift/alt key status</param>
    ''' <remarks></remarks>
    Public Sub keyCommands(ByVal keyCode As Integer, ByVal shift As Integer)
        If keyCode = Windows.Forms.Keys.Down And shift > 0 Then
            _arrowScale = _arrowScale * 0.9
            showLineFeedback(_lastPoint)
        ElseIf keyCode = Windows.Forms.Keys.Up And shift > 0 Then
            _arrowScale = _arrowScale * 1.2
            showLineFeedback(_lastPoint)
        ElseIf keyCode = Windows.Forms.Keys.F Then
            _flipArrows = Not _flipArrows
            If Not _pointNumber = 1 Then
                showLineFeedback(_lastPoint)
            End If
        ElseIf keyCode = Windows.Forms.Keys.U Then
            _angleIsSet = Not _angleIsSet
        ElseIf keyCode = Windows.Forms.Keys.Escape Then
            If _thisArrow.style = arrowStyles.Freeform And _pointNumber > 2 Then
                _pointNumber -= 1
            Else
                clearAll()
            End If
        ElseIf keyCode = Windows.Forms.Keys.S Then
            _arrowheadIsSwitched = Not _arrowheadIsSwitched
            showLineFeedback(_lastPoint)
        End If
    End Sub

    ''' <summary>
    ''' Activates the Select Features Tool to unselect an arrow tool due to 
    ''' being in the layout view
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub setDefaultTool()
        Dim uid As New UID
        uid.Value = "esriArcMapUI.selectFeaturesTool"
        Dim application As IApplication = _app
        application.CurrentTool = application.Document.CommandBars.Find(uid)
    End Sub

    ''' <summary>
    ''' Places a freeform arrow when the user double-clicks or selects finish from
    ''' the context menu
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub placeFreeformArrow()
        placeArrows(_lastPoint)
        resetSubtype()
        _pointNumber = 1
        ReDim _freeformPoints(0)
        'secondPoint = Nothing
        'prevPoint.PutCoords(0, 0)
    End Sub

    ''' <summary>
    ''' Checks to see that the map in is data view. If not gives the option to change it.
    ''' </summary>
    ''' <returns>If the return is false the tool is cancelled</returns>
    ''' <remarks></remarks>
    Friend Function checkDataView() As Boolean
        Dim retVal As Boolean = True
        Dim mxDoc As IMxDocument = _app.document
        If mxDoc.ActiveView Is mxDoc.PageLayout Then
            If MsgBox("You must be in Data View to use this tool. " & _
                "Click OK to change to Data View.", MsgBoxStyle.OkCancel) = _
                MsgBoxResult.Ok Then
                mxDoc.ActiveView = mxDoc.FocusMap
            Else
                retVal = False
            End If
        End If
        Return retVal
    End Function

    ''' <summary>
    ''' Clears the graphics container
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub clearAll()
        Dim mxDoc As IMxDocument = _app.document
        Dim graphicsContainer As IGraphicsContainer = mxDoc.ActiveView.GraphicsContainer
        graphicsContainer.DeleteAllElements()
        mxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
        _pointNumber = 1
    End Sub

    ''' <summary>
    ''' Write the the XML settings file
    ''' </summary>
    ''' <remarks>Not used in this version</remarks>
    Private Sub WriteXML()
        Dim XMLobj As Xml.XmlTextWriter
        Dim enc As New System.[Text].UnicodeEncoding()
        XMLobj = New Xml.XmlTextWriter("C:\temp\dimensionArrows.xml", enc)
        XMLobj.Formatting = Xml.Formatting.Indented
        XMLobj.Indentation = 5
        XMLobj.WriteStartDocument()
        XMLobj.WriteStartElement("landHook")
        XMLobj.WriteStartElement("flipped")
        XMLobj.WriteAttributeString("x0", "0")
        XMLobj.WriteAttributeString("y0", "0")
        XMLobj.WriteAttributeString("x1", "-10")
        XMLobj.WriteAttributeString("y1", "0")
        XMLobj.WriteAttributeString("x2", "-8")
        XMLobj.WriteAttributeString("y2", "2")
        XMLobj.WriteEndElement()
        XMLobj.WriteStartElement("notflipped")
        XMLobj.WriteAttributeString("x0", "0")
        XMLobj.WriteAttributeString("y0", "0")
        XMLobj.WriteAttributeString("x1", "-10")
        XMLobj.WriteAttributeString("y1", "0")
        XMLobj.WriteAttributeString("x2", "-8")
        XMLobj.WriteAttributeString("y2", "-2")
        XMLobj.WriteEndElement()
        XMLobj.WriteEndElement()
        XMLobj.Close()
    End Sub

    ''' <summary>
    ''' Reads the XML settings file
    ''' </summary>
    ''' <param name="arrowType">The arrow geometry being extracted from the file</param>
    ''' <returns>A formatted string with coordinate pairs</returns>
    ''' <remarks>The XML file is in the program installation folder</remarks>
    Private Function ReadXML(ByVal arrowType As Integer) As String
        ReadXML = ""
        Dim xmlDoc As XmlDocument = New XmlDocument
        Dim xmlNodeList As XmlNodeList
        Dim xmlNode As XmlNode

        xmlDoc.Load(_installationFolder & "\dimensionArrowGeometry.xml")

        Dim nodeName As String = ""

        Select Case arrowType
            Case arrowCategories.Straight
                nodeName = "/arrowDef/straight"
            Case arrowCategories.LandHook
                If _flipArrows Then
                    nodeName = "/arrowDef/landHookFlipped"
                Else
                    nodeName = "/arrowDef/landHook"
                End If
            Case arrowCategories.NoDashes
                nodeName = "/arrowDef/curved0"
            Case arrowCategories.OneDash
                nodeName = "/arrowDef/curved1"
            Case arrowCategories.TwoDashes
                nodeName = "/arrowDef/curved2"
            Case arrowCategories.ThreeDashes
                nodeName = "/arrowDef/curved3"
            Case arrowCategories.FourDashes
                nodeName = "/arrowDef/curved4"
            Case Else
                Return ""
                Exit Function
        End Select

        'Loop through the nodes
        Dim count As Integer

        xmlNodeList = xmlDoc.SelectNodes(nodeName)
        Dim xmlChild As XmlNode
        For Each xmlNode In xmlNodeList
            For Each xmlChild In xmlNode
                For count = 0 To xmlChild.Attributes.Count - 1
                    ReadXML = ReadXML & xmlChild.Attributes.Item(count).InnerText & ","
                Next
            Next
        Next
        ReadXML = Left(ReadXML, Len(ReadXML) - 1)
    End Function
End Module
