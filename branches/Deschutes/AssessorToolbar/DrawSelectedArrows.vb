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
Imports ESRI.ArcGIS.esriSystem
Imports System.Math

<ComClass(DrawSelectedArrows.ClassId, DrawSelectedArrows.InterfaceId, DrawSelectedArrows.EventsId), _
 ProgId("AssessorToolbar.DrawSelectedArrows")> _
Public NotInheritable Class DrawSelectedArrows
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "43dced6d-8935-4402-8329-22559853e50c"
    Public Const InterfaceId As String = "961b9d15-9a74-4bc0-923b-2c0fae13b21e"
    Public Const EventsId As String = "bb951c0e-05b1-4d83-bfad-cef007c95ed5"
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
    Private WithEvents _activeViewEvents As Map

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "AssessorToolbar"  'localizable text 
        MyBase.m_caption = "DrawSelectedArrows"   'localizable text 
        MyBase.m_message = "Draws Direction Arrows for Selected Features."   'localizable text 
        MyBase.m_toolTip = "Draws Direction Arrows for Selected Features." 'localizable text 
        MyBase.m_name = MyBase.m_category & "_DrawSelectedArrows"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")


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

        Dim theMxDoc As IMxDocument = _application.Document
        Dim theActiveView As IActiveView = theMxDoc.FocusMap

        _buttonChecked = Not _buttonChecked
        If _buttonChecked Then
            _activeViewEvents = theMxDoc.FocusMap
        End If

        theActiveView.Refresh()

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


    'Public Overrides ReadOnly Property Checked() As Boolean
    '    Get
    '        Return _buttonChecked
    '    End Get
    'End Property


    Private Sub _pActiveViewEvents_AfterDraw(ByVal theDisplay As IDisplay, ByVal phase As esriViewDrawPhase) Handles _activeViewEvents.AfterDraw

        ' Only draw in the selection phase

        If phase = esriViewDrawPhase.esriViewGeoSelection And _buttonChecked And _editor.EditState = esriEditState.esriStateEditing Then

            Dim theFeatLayer As IFeatureLayer
            Dim arrayList As New ArrayList()

            Dim theMxDocument As IMxDocument = _application.Document
            Dim theDispTransformation As IDisplayTransformation = theDisplay.DisplayTransformation

            ' Get the cancel tracker to stop is required
            Dim theTrackCancel As ITrackCancel = Nothing
            If TypeOf theDisplay Is IScreenDisplay Then
                Dim theScreenDisplay As IScreenDisplay = theDisplay
                theTrackCancel = theMxDocument.ActiveView.ScreenDisplay.CancelTracker
            End If

            Dim theMap As IMap = theMxDocument.FocusMap
            Dim theEnumLayer As IEnumLayer = theMap.Layers
            Dim theLayer As ILayer = theEnumLayer.Next

            Do Until theLayer Is Nothing
                If TypeOf theLayer Is IFeatureLayer And theLayer.Valid Then
                    theFeatLayer = theLayer
                    arrayList.Add(theFeatLayer.FeatureClass.AliasName)
                    If theLayer.Visible Then
                        arrayList.Add("1")
                    Else
                        arrayList.Add("0")
                    End If
                End If
                theLayer = theEnumLayer.Next
            Loop

            ' Create a symbol. Assume the symbol for the selected polyline is cyan
            Dim theRGBColor As IRgbColor = New RgbColor
            theRGBColor.Red = 255
            theRGBColor.Blue = 0
            theRGBColor.Green = 0

            Dim theArrowMrkrSymbol As IArrowMarkerSymbol = New ArrowMarkerSymbol
            theArrowMrkrSymbol.Style = esriArrowMarkerStyle.esriAMSPlain

            ' Make sure the symbol size is large enough to draw something
            Dim theArrowSize As Double = 12

            If theMxDocument.FocusMap.ReferenceScale > 0 Then

                ' Get the device frame which will give us the number of pixels in the X direction
                Dim theDeviceRECT As tagRECT = theDispTransformation.DeviceFrame
                Dim thePixelExtent As Long = theDeviceRECT.right - theDeviceRECT.left
                Dim theVisibleBoundsEnvelope As IEnvelope = theDispTransformation.VisibleBounds

                ' Calculate the size of one pixel
                Dim theRealWorldDisplayExtent As Double = theVisibleBoundsEnvelope.Width
                Dim theSizeOfOnePixel As Double = theRealWorldDisplayExtent / thePixelExtent
                If theArrowSize < (theSizeOfOnePixel * 2) Then
                    theArrowSize = theSizeOfOnePixel * 2
                End If
            End If

            theArrowMrkrSymbol.Size = theArrowSize
            theArrowMrkrSymbol.Color = theRGBColor
            Dim theMrkrSymbol As IMarkerSymbol = theArrowMrkrSymbol

            Dim theLine As ILine = New Line

            ' Get the selected features (in an enumeration)
            Dim theEnumFeature As IEnumFeature = theMxDocument.FocusMap.FeatureSelection
            theEnumFeature.Reset()

            ' Get the displayed features from the line target layer
            'Dim theEditLayers As IEditLayers
            'theEditLayers = _editor
            'Dim theTargetLayer As IFeatureLayer
            'theTargetLayer = theEditLayers.CurrentLayer

            Dim thisFeature As IFeature
            Dim thisGeomColl As IGeometryCollection
            Dim thisSegColl As ISegmentCollection
            Dim thisSegment As ISegment
            Dim thisGeometry As Long
            Dim thisAngle As Double
            Dim thisLength As Double
            Dim thisObjClass As IObjectClass
            Dim thisFeatureClassName As String
            Dim thisFeatClassIsVisible As Boolean


            ' Progress through the selection the given number of steps
            For lLoop As Long = 0 To theMxDocument.FocusMap.SelectionCount - 1
                ' Get the geometry collection
                thisFeature = theEnumFeature.Next
                If Not thisFeature Is Nothing Then
                    If thisFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then 'And theTargetLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                        ' Work out if this feature class is being displayed
                        thisObjClass = thisFeature.Class
                        thisFeatureClassName = thisObjClass.AliasName
                        thisFeatClassIsVisible = False
                        For c As Integer = 0 To arrayList.Count - 1 Step 2
                            If arrayList(c) = thisFeatureClassName AndAlso arrayList(c + 1) = 1 Then
                                thisFeatClassIsVisible = True
                            End If
                        Next

                        If thisFeatClassIsVisible Then
                            thisGeomColl = thisFeature.Shape
                            ' Look at the multipart geometries
                            For thisGeometry = 0 To thisGeomColl.GeometryCount - 1
                                thisSegColl = thisGeomColl.Geometry(thisGeometry)
                                ' Get the last segment
                                thisSegment = thisSegColl.Segment(thisSegColl.SegmentCount - 1)
                                ' Calculate the location of the arrow marker symbol. The position
                                ' is 7 pixels from the end of the line and the arrow is rotated tangent
                                ' to the last line segment.
                                thisLength = thisSegment.Length - ConvertPixelsToRW(7)
                                If thisLength <= 0 Then
                                    thisLength = thisSegment.Length
                                End If
                                If thisLength > 0 Then
                                    thisSegment.QueryTangent(esriSegmentExtension.esriExtendTangentAtTo, thisLength, False, 1, theLine)
                                    thisAngle = AngleFromCoords(theLine.FromPoint.X, theLine.FromPoint.Y, _
                                      theLine.ToPoint.X, theLine.ToPoint.Y)
                                    ' Draw the arrow
                                    theArrowMrkrSymbol.Angle = thisAngle
                                    theDisplay.SetSymbol(theMrkrSymbol)
                                    theDisplay.DrawPoint(theLine.FromPoint)
                                End If
                            Next thisGeometry
                        End If
                    End If
                    If Not theTrackCancel Is Nothing Then
                        If Not theTrackCancel.Continue Then
                            Exit For
                        End If
                    End If
                End If
            Next
        End If


    End Sub

    Private Function AngleFromCoords(ByVal FromX As Double, ByVal FromY As Double, ByVal ToX As Double, ByVal ToY As Double) As Double

        ' Simple algorithm to return the mathematical angle between two points

        Dim PI As Double = 4.0# * Atan(1.0#)
        Dim dX As Double = ToX - FromX
        Dim dY As Double = ToY - FromY

        Dim dBrg As Double
        Dim dDist As Double
        Dim dAngle As Double
        Dim dValue As Double

        If dX = 0 Then
            If dY >= 0 Then
                dBrg = 0
            Else
                dBrg = PI
            End If
        ElseIf dY = 0 Then
            If dX >= 0 Then
                dBrg = PI / 2
            Else
                dBrg = 3 * PI / 2
            End If
        Else
            dDist = ((ToX - FromX) ^ 2 + (ToY - FromY) ^ 2) ^ 0.5
            If dDist = 0 Then
                dBrg = 0
            Else
                dValue = Abs(dX) / dDist
                dAngle = Atan(dValue / Sqrt(-dValue * dValue + 1))

                If dX > 0 Then
                    If dY >= 0 Then
                        dBrg = dAngle
                    Else
                        dBrg = PI - dAngle
                    End If
                Else
                    If dY >= 0 Then
                        dBrg = PI * 2 - dAngle
                    Else
                        dBrg = dAngle + PI
                    End If
                End If
            End If
        End If

        Return 90 - dBrg * 180 / PI

    End Function

    Private Function ConvertPixelsToRW(ByVal pixelUnits As Double) As Double

        Dim theMxDocument As IMxDocument = _application.Document
        Dim theActiveView As IActiveView = theMxDocument.FocusMap
        Dim theDisplayTrans As IDisplayTransformation = theActiveView.ScreenDisplay.DisplayTransformation
        Dim theDeviceRECT As tagRECT = theDisplayTrans.DeviceFrame
        Dim pixelExtent As Long = theDeviceRECT.right - theDeviceRECT.left
        Dim theVisibleBoundsEnv As IEnvelope = theDisplayTrans.VisibleBounds
        Dim realWorldDisplayExtent As Double = theVisibleBoundsEnv.Width
        Dim sizeOfOnePixel As Double = realWorldDisplayExtent / pixelExtent
        Return pixelUnits * sizeOfOnePixel

    End Function






End Class



