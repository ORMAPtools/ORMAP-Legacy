Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display

Imports AssessorToolbar.Utilities
Imports System.Windows.forms

<ComClass(DrawNeatLine.ClassId, DrawNeatLine.InterfaceId, DrawNeatLine.EventsId), _
 ProgId("AssessorToolbar.DrawNeatLine")> _
Public NotInheritable Class DrawNeatLine
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "f7e8bc91-9b1a-4e62-9cf7-a08bd445a03d"
    Public Const InterfaceId As String = "01b3a6d0-f117-4150-b3af-dea5672f8eab"
    Public Const EventsId As String = "f78ce169-df33-4549-9457-58858db49c4f"
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
        MyBase.m_caption = "DrawNeatLine"   'localizable text 
        MyBase.m_message = "Draws the map neat Line."   'localizable text 
        MyBase.m_toolTip = "Draws the map neat Line." 'localizable text 
        MyBase.m_name = MyBase.m_category & "_DrawNeatLine"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
            Exit Sub
        End If
        If GetFeatureLayerByName("SectionLines", _application.Document) Is Nothing Then
            MessageBox.Show("Unable to find the SectionLines feature class.  Please ensure it's loaded into your project", "Error", MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim theSelectMapIndexDialog As SelectMapindexDialog = MakeSelectMapIndexDialog(_application.Document)
        If theSelectMapIndexDialog.ShowDialog = DialogResult.Cancel Then Exit Sub

        DrawNeatLine(theSelectMapIndexDialog.MapNumber)

    End Sub


    Private Sub DrawNeatLine(ByVal theMapNumber As String)

        Dim theSeeMapsFC As IFeatureClass = GetFeatureLayerByName("SeeMaps", _application.Document).FeatureClass
        Dim theQueryFilter As IQueryFilter = New QueryFilter
        theQueryFilter.WhereClause = "MapNumber = '" & theMapNumber & "'"

        Dim theFeatureCursor As IFeatureCursor = theSeeMapsFC.Search(theQueryFilter, True)
        Dim theFeature As IFeature = theFeatureCursor.NextFeature

        Dim mapScale As Long = theFeature.Value(theFeatureCursor.FindField("MapScale"))



        Dim theEnvelope As IEnvelope = theFeature.Shape.Envelope
        Dim theNeatLineEnvelope As IEnvelope = AddTRSEnvelope(theEnvelope, theMapNumber, mapScale)
        '-- Buffer the neatline
        Select Case mapScale
            Case 240000
                theNeatLineEnvelope.Expand(1.05, 1.05, True) ' 5%
            Case Else
                theNeatLineEnvelope.Expand(1.15, 1.15, True) ' 15%
        End Select

        '-- Check to make sure width is not over 18"... the neatline on the map is cut at 18"
        Dim theWidth As Double = (theNeatLineEnvelope.XMax - theNeatLineEnvelope.XMin) / (mapScale / 12)
        Dim theHeight As Double = (theNeatLineEnvelope.YMax - theNeatLineEnvelope.YMin) / (mapScale / 12)
        Dim theOffset As Double = 0
        If theWidth > 18 Then
            theOffset = ((theWidth - 17.99) * (mapScale / 12)) / 2
            theNeatLineEnvelope.XMax -= theOffset
            theNeatLineEnvelope.XMin += theOffset
        End If
        If theHeight > 18 Then
            theOffset = ((theHeight - 17.99) * (mapScale / 12)) / 2
            theNeatLineEnvelope.YMax -= theOffset
            theNeatLineEnvelope.YMin += theOffset
        End If

        Dim theMxDoc As IMxDocument = _application.Document
        Dim pActiveView As IActiveView = theMxDoc.FocusMap
        Dim pGraphicsContainer As IGraphicsContainer = pActiveView.GraphicsContainer
        pGraphicsContainer.DeleteAllElements()

        Dim pRect As IRectangleElement = New RectangleElement
        Dim pElement As IElement = pRect
        pElement.Geometry = theNeatLineEnvelope

        'Set the Neatline symbology
        Dim pRectSym As IFillShapeElement = pRect

        'Set the style of the fill symbol
        Dim pFillSymbol As ISimpleFillSymbol = New SimpleFillSymbol
        pFillSymbol.Style = esriSimpleFillStyle.esriSFSHollow

        'Set the outline symbol
        Dim pLineSymbol As ISimpleLineSymbol = pFillSymbol.Outline
        pLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid
        pLineSymbol.Width = 3
        pLineSymbol.Color = GetRGBColor(Color.Red)
        pFillSymbol.Outline = pLineSymbol

        'Set it all into the element
        pRectSym.Symbol = pFillSymbol

        'Draw the neat line
        pGraphicsContainer.AddElement(pRect, 0)
        pActiveView.Refresh()

    End Sub


    Private Function AddTRSEnvelope(ByVal theEnvelope As IEnvelope, ByVal mapNumber As String, ByVal mapScale As Long) As IEnvelope

        Dim theQueryFilter As IQueryFilter = New QueryFilter
        Dim theNewEnvelope As IEnvelope = theEnvelope

        Select Case mapScale

            Case 1200
                theQueryFilter.WhereClause = "TRSQX = '" & Left(mapNumber, 8) & "'"
            Case 2400
                theQueryFilter.WhereClause = "TRSQ = '" & Left(mapNumber, 7) & "'"
            Case 4800
                theQueryFilter.WhereClause = "TRS = '" & Left(mapNumber, 6) & "'"
            Case 24000
                theQueryFilter.WhereClause = "TR = '" & Left(mapNumber, 4) & "'"
            Case Else
                Return theNewEnvelope
                Exit Function

        End Select

        Dim theFeatureClass As IFeatureClass = GetFeatureLayerByName("Sixteenth_Section", _application.Document).FeatureClass

        Dim theFeatureCursor As IFeatureCursor = theFeatureClass.Search(theQueryFilter, True)
        Dim thisFeature As IFeature = theFeatureCursor.NextFeature  'Get the first feature

        Do Until thisFeature Is Nothing
            theNewEnvelope.Union(thisFeature.Shape.Envelope)
            thisFeature = theFeatureCursor.NextFeature
        Loop

        Return theNewEnvelope

    End Function


End Class



