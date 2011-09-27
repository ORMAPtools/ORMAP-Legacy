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
Imports ESRI.ArcGIS.SystemUI
Imports stdole


Module SpiralUtilities
    Dim _editor As IEditor3 = CType(My.ArcMap.Editor, IEditor3)
    Friend Function IsEnable() As Boolean
        Dim IsEditing As Boolean
        If My.ArcMap.Editor.EditState = esriEditState.esriStateNotEditing Then
            IsEditing = False
        Else
            IsEditing = True
        End If
        Return IsEditing
    End Function
    Friend Function getDataFrameCoords(ByVal X As Integer, ByVal Y As Integer) As IPoint
        'Dim displayTransformation As ESRI.ArcGIS.Display.IDisplayTransformation
        'displayTransformation = _app.Display.DisplayTransformation
        Dim theDisplayTransformation As IDisplayTransformation = My.ThisApplication.Display.DisplayTransformation

        Return theDisplayTransformation.ToMapPoint(X, Y)
    End Function
    Function getSnapPoint(ByVal point As IPoint) As IPoint
        Dim snapEnv As ISnapEnvironment = CType(_editor, ISnapEnvironment)
        snapEnv.SnapPoint(point)
        Return point
    End Function
    Public Sub ConstructSCSbyLength(ByVal theFromPoint As IPoint, ByVal theTangentPoint As IPoint, ByVal theToPoint As IPoint, ByVal theSpiralLengths As Double, ByVal theRadius As Double, ByVal isCCW As Boolean)
        If My.ArcMap.Editor.EditState = esriEditState.esriStateNotEditing Then
            Exit Sub
        End If
        Try
            Dim toCurvature As Double = 1 / theRadius
            Dim DensifyParameter As Double = 0.5

            'Constructs the spiral curves
            Dim theFirstSpiralPolyLine As IPolyline6 = Construct_Spiral_by_length(theFromPoint, theTangentPoint, 0, toCurvature, isCCW, theSpiralLengths)

            If isCCW Then
                isCCW = False
            Else
                isCCW = True
            End If

            Dim theSecondSpiralPolyLine As IPolyline6 = Construct_Spiral_by_length(theToPoint, theTangentPoint, 0, toCurvature, isCCW, theSpiralLengths)

            'Constructs the Central Curve
            Dim TheCentralCurveConstruction As IConstructCircularArc2 = New CircularArc
            TheCentralCurveConstruction.ConstructEndPointsRadius(theFirstSpiralPolyLine.ToPoint, theSecondSpiralPolyLine.ToPoint, isCCW, theRadius, True)
            Dim theCentralCurve As ICurve3 = TryCast(TheCentralCurveConstruction, ICurve3)
            Dim TheCurvePolyline As ISegmentCollection = New PolylineClass()
            TheCurvePolyline.AddSegment(TryCast(TheCentralCurveConstruction, ISegment))

            Dim theFeatureclass As IFeatureClass = CType(My.ArcMap.Editor.Map.Layer(0), IFeatureLayer2).FeatureClass
            Dim theFirstSpiralFeature As IFeature = theFeatureclass.CreateFeature
            Dim theSecondSpiralFeature As IFeature = theFeatureclass.CreateFeature
            Dim theCenterCircularFeature As IFeature = theFeatureclass.CreateFeature

            'Add the new features to the feature Class
            My.ArcMap.Editor.StartOperation()
            theFirstSpiralFeature.Shape = CType(theFirstSpiralPolyLine, IGeometry)
            theFirstSpiralFeature.Store()
            theCenterCircularFeature.Shape = CType(TheCurvePolyline, IGeometry)
            theCenterCircularFeature.Store()
            theSecondSpiralFeature.Shape = CType(theSecondSpiralPolyLine, IGeometry)
            theSecondSpiralFeature.Store()
            My.ArcMap.Editor.StopOperation("Finished Construction")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub
    
    Public Function Create_Snap_Marker() As IMarkerElement
        Dim TheMarkerElement As IMarkerElement = New MarkerElement
        Dim theMarkerSymbol As ICharacterMarkerSymbol = New CharacterMarkerSymbol
        Dim theSnapFont As stdole.IFontDisp = CType(New stdole.StdFont, stdole.IFontDisp)

        With theSnapFont
            .Name = "ESRI Default Marker"
            .Size = My.ArcMap.Document.SearchTolerancePixels
        End With

        With theMarkerSymbol
            .Font = theSnapFont
            .CharacterIndex = 40
        End With

        TheMarkerElement.Symbol = theMarkerSymbol

        Return TheMarkerElement
    End Function
    Private Function Construct_Spiral_by_length(ByVal theFromPoint As IPoint, ByVal theTangentpoint As IPoint, ByRef theFromCurvature As Double, ByRef theToCurvature As Double, ByVal isCCW As Boolean, ByVal theSpiralLength As Double) As IPolyline6
        Dim thePolyLine As IPolyline6 = CType(New Polyline, IPolyline6)

        Try
            Dim theGeometryEnvironment As IGeometryEnvironment4 = New GeometryEnvironment
            Dim TheSpiralConstruction As IConstructClothoid = CType(theGeometryEnvironment, IConstructClothoid)
            thePolyLine = CType(TheSpiralConstruction.ConstructClothoidByLength(theFromPoint, theTangentpoint, isCCW, theFromCurvature, theToCurvature, theSpiralLength, esriCurveDensifyMethod.esriCurveDensifyByLength, densifyParameter:=0.5), IPolyline6)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Return thePolyLine
    End Function
End Module

