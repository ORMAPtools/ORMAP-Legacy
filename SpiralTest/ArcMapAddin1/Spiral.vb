Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.DisplayUI
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.DataSourcesFile

Imports System

Public Class Spiral
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool

    Public Sub New()
        'My.ArcMap.Application.OpenDocument("C:\SpiralTest\Test.mxd")

        Try
            Dim theEditorUID As IUID = New UID
            theEditorUID.Value = "esriEditor.Editor"


            Dim theEditor As IEditor = My.ArcMap.Application.FindExtensionByCLSID(theEditorUID)

            'Dim factoryType As Type = Type.GetTypeFromProgID("esriDataSourcesGDB.FileGDBWorkspaceFactory")
            'Dim workspaceFactory As IWorkspaceFactory = CType(Activator.CreateInstance(factoryType), IWorkspaceFactory)
            'Dim theWorkspace As IWorkspace = CType(workspaceFactory.OpenFromFile("U:\ArcGIS 10\SpiralCurve.gdb", 0), IWorkspace)

            ''Dim theWorkSpaceFactory As IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesGDB.FileGDBWorkspaceFactory
            ''Dim theWorkspace As IWorkspace = CType(theWorkSpaceFactory.OpenFromFile("U:\ArcGIS 10\SpiralCurve.gdb", 0), IWorkspace)
            'theEditor.StartEditing(theWorkspace)

            Dim theMxd As IMxDocument = My.ArcMap.Document


            Dim theGeometryEnvironment As IGeometryEnvironment4 = New GeometryEnvironment


            Dim SpiralCurve As IConstructClothoid = theGeometryEnvironment

            Dim TheCurveLength As Double = 400


            'ConstructClothoidbyAngle Method properties
            Dim DeflectionAngle As Double = 0.12930951

            'Universal to all methods 
            Dim TheFromPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
            Dim TheToPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
            Dim TheTangentPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
            Dim FromCurvature As Double = 0 'user argument
            Dim toCurvature As Double = 0.0001745328 'user argument This value is determined by 1/radius.
            Dim DensifyParameter As Double = 0.5 'may become a menu variable.
            Dim IsCurveRight As Boolean = True 'User argument

            TheFromPoint.PutCoords(7668316.7435, 600358.534)
            TheToPoint.PutCoords(7671214.385, 600741.731)
            TheTangentPoint.PutCoords(7669722.1515, 600878.41025)


            Dim theMaps As IMaps2 = theMxd.Maps
            Dim theMap As IMap = theMaps.Item(0)
            Dim theFeaturelayer As IFeatureLayer2 = theMap.Layer(0)
            Dim theFeatureClass As IFeatureClass = theFeaturelayer.FeatureClass
            Dim theFirstSpiralFeature As IFeature = theFeatureClass.CreateFeature()
            Dim theSecondSpiralFeature As IFeature = theFeatureClass.CreateFeature()

            Dim theFirstSpiralPolyline As IPolyline6 = New Polyline
            Dim theSecondSpiralPolyline As IPolyline6 = New Polyline

            'By(Angle)
            'thePolyline = SpiralCurve.ConstructClothoidByAngle(TheFromPoint, TheTangentPoint, IsCurveRight, FromCurvature, toCurvature, DeflectionAngle, esriCurveDensifyMethod.esriCurveDensifyByAngle, DensifyParameter)
            'By Length
            theFirstSpiralPolyline = SpiralCurve.ConstructClothoidByLength(TheFromPoint, TheTangentPoint, IsCurveRight, FromCurvature, toCurvature, TheCurveLength, esriCurveDensifyMethod.esriCurveDensifyByLength, DensifyParameter)

            theEditor.StartOperation()
            theFirstSpiralFeature.Shape = theFirstSpiralPolyline
            theFirstSpiralFeature.Store()
            theEditor.StopOperation("Done")

            IsCurveRight = False
            theSecondSpiralPolyline = SpiralCurve.ConstructClothoidByLength(TheToPoint, TheTangentPoint, IsCurveRight, FromCurvature, toCurvature, TheCurveLength, esriCurveDensifyMethod.esriCurveDensifyByLength, DensifyParameter)
            theEditor.StartOperation()
            theSecondSpiralFeature.Shape = theSecondSpiralPolyline
            theSecondSpiralFeature.Store()
            theEditor.StopOperation("Done 2")

            Dim theRadius As Double = 5729.58
            Dim TheCentralCurveConstruction As IConstructCircularArc2 = New CircularArc
            TheCentralCurveConstruction.ConstructEndPointsRadius(theFirstSpiralPolyline.ToPoint, theSecondSpiralPolyline.ToPoint, False, theRadius, True)
            Dim theCentralCurve As ICurve3 = TryCast(TheCentralCurveConstruction, ICurve3)
            Dim TheCurvePolyline As ISegmentCollection = New PolylineClass()
            TheCurvePolyline.AddSegment(TryCast(TheCentralCurveConstruction, ISegment))


            Dim TheCurveFeature As IFeature = theFeatureClass.CreateFeature
            theEditor.StartOperation()
            TheCurveFeature.Shape = TryCast(TheCurvePolyline, IGeometry)
            TheCurveFeature.Store()
            theEditor.StopOperation("Done 3")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        My.ArcMap.Document.ActiveView.Refresh()
       
    End Sub

    Protected Overrides Sub OnUpdate()
        Enabled = My.ArcMap.Application IsNot Nothing
    End Sub
End Class
