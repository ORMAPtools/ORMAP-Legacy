Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display

Imports System.Windows.Forms
Imports AssessorToolbar.Utilities

<ComClass(DrawSectionGraphic.ClassId, DrawSectionGraphic.InterfaceId, DrawSectionGraphic.EventsId), _
 ProgId("AssessorToolbar.DrawSectionGraphic")> _
Public NotInheritable Class DrawSectionGraphic
    Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "6c8b8f2f-f2cf-4917-83cb-d2343d4a7905"
    Public Const InterfaceId As String = "9e2eda5c-b73d-4b87-a860-dff9cf500c7d"
    Public Const EventsId As String = "76457e65-b428-4fc3-a86e-337b664712c5"
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

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "AssessorToolbar"  'localizable text 
        MyBase.m_caption = "DrawSectionGraphic"   'localizable text 
        MyBase.m_message = "Draw a 1"" graphic box around mouse click"   'localizable text 
        MyBase.m_toolTip = "Draw a 1"" graphic box around mouse click" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_DrawSectionGraphic"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")


        Try
            'TODO: change resource name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
            MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), Me.GetType().Name + ".cur")
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

        '-- Check for necessary items before proceeding...
        If GetFeatureLayerByName("SeeMaps", _application.Document) Is Nothing Then
            MessageBox.Show("Unable to find the SeeMaps feature class.  Please ensure it's loaded into your project", "Error", MessageBoxButtons.OK)
            _application.CurrentTool = Nothing
        End If

    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add DrawSectionGraphic.OnMouseDown implementation
    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add DrawSectionGraphic.OnMouseMove implementation
    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)

        Dim theSelectMapIndexDialog As SelectMapindexDialog = MakeSelectMapIndexDialog(_application.Document)
        If theSelectMapIndexDialog.ShowDialog = DialogResult.Cancel Then Exit Sub

        Dim theMxDoc As IMxDocument = _application.Document
        Dim pActiveView As IActiveView = theMxDoc.FocusMap

        Dim theSeeMapsFC As IFeatureClass = GetFeatureLayerByName("SeeMaps", _application.Document).FeatureClass
        Dim theQueryFilter As IQueryFilter = New QueryFilter
        theQueryFilter.WhereClause = "MapNumber = '" & theSelectMapIndexDialog.MapNumber & "'"

        Dim theFeatureCursor As IFeatureCursor = theSeeMapsFC.Search(theQueryFilter, True)
        Dim theFeature As IFeature = theFeatureCursor.NextFeature
        Dim mapScale As Long = theFeature.Value(theFeatureCursor.FindField("MapScale"))

        Dim theOffSet As Double = (mapScale / 12) / 2

        Dim pPoint As IPoint
        pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)

        Dim theEnv As IEnvelope = New Envelope
        theEnv.XMax = pPoint.X + theOffSet
        theEnv.XMin = pPoint.X - theOffSet
        theEnv.YMax = pPoint.Y + theOffSet
        theEnv.YMin = pPoint.Y - theOffSet

        Dim pGraphicsContainer As IGraphicsContainer = pActiveView.GraphicsContainer
        pGraphicsContainer.DeleteAllElements()

        Dim pRect As IRectangleElement = New RectangleElement
        Dim pElement As IElement = pRect
        pElement.Geometry = theEnv

        'Set the Neatline symbology
        Dim pRectSym As IFillShapeElement = pRect

        'Set the style of the fill symbol
        Dim pFillSymbol As ISimpleFillSymbol = New SimpleFillSymbol
        pFillSymbol.Style = esriSimpleFillStyle.esriSFSHollow

        'Set the outline symbol
        Dim pLineSymbol As ISimpleLineSymbol = pFillSymbol.Outline
        pLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid
        pLineSymbol.Width = 1
        pLineSymbol.Color = GetRGBColor(Color.Red)
        pFillSymbol.Outline = pLineSymbol

        'Set it all into the element
        pRectSym.Symbol = pFillSymbol

        'Draw the neat line
        pGraphicsContainer.AddElement(pRect, 0)
        pActiveView.Refresh()

        _application.CurrentTool = Nothing

    End Sub
End Class

