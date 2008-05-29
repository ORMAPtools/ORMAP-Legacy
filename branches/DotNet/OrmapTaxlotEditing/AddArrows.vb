#Region "Copyright 2008 ORMAP Tech Group"

' File:  AddArrows.vb
'
' Original Author:  OPET.NET Migration Team (Shad Campbell, James Moore, 
'                   Nick Seigal)
'
' Date Created:  January 8, 2008
'
' Copyright Holder:  ORMAP Tech Group  
' Contact Info:  ORMAP Tech Group (a.k.a. opet developers) may be reached at 
' opet-developers@lists.sourceforge.net
'
' This file is part of the ORMAP Taxlot Editing Toolbar.
'
' ORMAP Taxlot Editing Toolbar is free software; you can redistribute it and/or
' modify it under the terms of the Lesser GNU General Public License as 
' published by the Free Software Foundation; either version 3 of the License, 
' or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT 
' ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or 
' FITNESS FOR A PARTICULAR PURPOSE.  See the Lesser GNU General Public License 
' located in the COPYING.LESSER.txt file for more details.
'
' You should have received a copy of the Lesser GNU General Public License 
' along with the ORMAP Taxlot Editing Toolbar; if not, write to the Free 
' Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 
' 02110-1301 USA.

#End Region

#Region "Subversion Keyword Expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision$
'Date of Last Change: $Date$
#End Region

#Region "Imported Namespaces"
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto

Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework

Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase

Imports OrmapTaxlotEditing.DataMonitor
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.StringUtilities
Imports OrmapTaxlotEditing.Utilities

Imports System.Text

#End Region


<ComVisible(True)> _
<ComClass(AddArrows.ClassId, AddArrows.InterfaceId, AddArrows.EventsId), _
ProgId("ORMAPTaxlotEditing.AddArrows")> _
Public NotInheritable Class AddArrows
    Inherits BaseTool
    Implements IDisposable

#Region "Class-Level Constants And Enumerations"
    Private Const _ignoreCase As StringComparison = StringComparison.CurrentCultureIgnoreCase
#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' Define protected instance field values for the public properties
        MyBase.m_category = "OrmapToolbar"  'localizable text 
        MyBase.m_caption = "AddArrows"   'localizable text 
        MyBase.m_message = "Add arrow features to the cartographic lines feature class."   'localizable text 
        MyBase.m_toolTip = "Add Arrows" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_AddArrows"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            ' Set the bitmap based on the name of the class.
            _bitmapResourceName = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), _bitmapResourceName)
        Catch ex As ArgumentException
            EditorExtension.ProcessUnhandledException(ex)
        End Try

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication
    Private _bitmapResourceName As String

    Private _theMxDoc As IMxDocument
    Private _theMap As IMap
    Private _theLineSymbol As ILineSymbol
    Private _theArrowPt1 As IPoint
    Private _theArrowPt2 As IPoint
    Private _theArrowPt3 As IPoint
    Private _theArrowPt4 As IPoint
    Private _theFromBreakPoint As IPoint ' hooks
    Private _theStartPoint As IPoint ' hooks
    Private _theTextPoint As IPoint
    Private _theToBreakPoint As IPoint ' hooks
    Private _theLinePolyline As IPolyline
    Private _theTextSymbol As ITextSymbol
    Private _theHookAngle As Double
    Private _theDoOnce As Boolean ' testing
    Private _theInUse As Boolean

    Private _theArrowPtTemp As IPoint
    Private _theArrowPtTemp2 As IPoint
    Private _thePt As IPoint
    Private _theSnapAgent As IFeatureSnapAgent
    Private _theMouseHasMoved As Boolean
    Private _theInTol As Boolean

    Private _theMxApp As IMxApplication
    Private _theActiveView As IActiveView

    Private WithEvents m_pEditorEvents As Editor 'Deal with this...
    Private m_pEditor As IEditor2 'Deal with this...

    Private _theToolJustCompletedTask As Boolean
    Private _theolDimensionChanged As Boolean
    Dim _thelArrowPointsCollection As Collection



#End Region

#Region "Properties"

    Private WithEvents _partnerAddArrowsForm As AddArrowsForm

    Friend ReadOnly Property PartnerAddArrowsForm() As AddArrowsForm
        Get
            If _partnerAddArrowsForm Is Nothing OrElse _partnerAddArrowsForm.IsDisposed Then
                setPartnerAddArrowsForm(New AddArrowsForm())
            End If
            Return _partnerAddArrowsForm
        End Get
    End Property

    Private Sub setPartnerAddArrowsForm(ByVal value As AddArrowsForm)
        If value IsNot Nothing Then
            _partnerAddArrowsForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerAddArrowsForm.Load, AddressOf PartnerAddArrowsForm_Load
            AddHandler _partnerAddArrowsForm.uxQuit.Click, AddressOf uxQuit_Click
            AddHandler _partnerAddArrowsForm.uxHelp.Click, AddressOf uxHelp_Click
            AddHandler _partnerAddArrowsForm.uxAddStandard.Click, AddressOf uxAddStandard_Click
            AddHandler _partnerAddArrowsForm.uxAddDimension.Click, AddressOf uxAddDimension_Click
        Else
            ' Unsubscribe to partner form events.
            RemoveHandler _partnerAddArrowsForm.Load, AddressOf PartnerAddArrowsForm_Load
            RemoveHandler _partnerAddArrowsForm.uxQuit.Click, AddressOf uxQuit_Click
            RemoveHandler _partnerAddArrowsForm.uxHelp.Click, AddressOf uxHelp_Click
            RemoveHandler _partnerAddArrowsForm.uxAddStandard.Click, AddressOf uxAddStandard_Click
            RemoveHandler _partnerAddArrowsForm.uxAddDimension.Click, AddressOf uxAddDimension_Click
        End If

    End Sub

    Private WithEvents _partnerDimensionArrowsForm As DimensionArrowsForm

    Friend ReadOnly Property PartnerDimensionArrowsForm() As DimensionArrowsForm
        Get
            If _partnerDimensionArrowsForm Is Nothing OrElse _partnerDimensionArrowsForm.IsDisposed Then
                setPartnerDimensionArrowsForm(New DimensionArrowsForm())
            End If
            Return _partnerDimensionArrowsForm
        End Get
    End Property

    Private Sub setPartnerDimensionArrowsForm(ByVal value As DimensionArrowsForm)
        If value IsNot Nothing Then
            _partnerDimensionArrowsForm = value
            ' Subscribe to partner form events.
        Else
            ' Unsubscribe to partner form events.
        End If
    End Sub

    Private _ratioLine As Double
    Friend Property RatioLine() As Double
        Get
            _ratioLine = CInt(PartnerDimensionArrowsForm.uxRatioOfLine.Text)
            Return _ratioLine
        End Get
        Set(ByVal value As Double)
            _ratioLine = value
            PartnerDimensionArrowsForm.uxRatioOfLine.Text = CStr(_ratioLine)
        End Set
    End Property

    Private _ratioCurve As Double
    Friend Property RatioCurve() As Double
        Get
            _ratioCurve = CInt(PartnerDimensionArrowsForm.uxRatioOfCurve.Text)
            Return _ratioCurve
        End Get
        Set(ByVal value As Double)
            _ratioCurve = value
            PartnerDimensionArrowsForm.uxRatioOfCurve.Text = CStr(_ratioCurve)
        End Set
    End Property

    Private _smoothRatio As Double
    Friend Property SmoothRatio() As Double
        Get
            _smoothRatio = CInt(PartnerDimensionArrowsForm.uxSmoothRatio.Text)
            Return _smoothRatio
        End Get
        Set(ByVal value As Double)
            _smoothRatio = value
            PartnerDimensionArrowsForm.uxSmoothRatio.Text = CStr(_smoothRatio)
        End Set
    End Property

    Private _addManually As Boolean
    Friend Property AddManually() As Boolean
        Get
            _addManually = PartnerDimensionArrowsForm.uxManuallyAddArrow.Checked
            Return _addManually
        End Get
        Set(ByVal value As Boolean)
            _addManually = value
            PartnerDimensionArrowsForm.uxManuallyAddArrow.Checked = _addManually
        End Set
    End Property

    Private _arrowType As String
    Friend Property arrowType() As String
        Get
            Return _arrowType
        End Get
        Set(ByVal value As String)
            _arrowType = value.ToUpper
        End Set
    End Property

    Private _arrowLineStyle As Integer
    Friend Property arrowLineStyle() As Integer
        Get
            '_arrowLineStyle = CInt(PartnerAddArrowsForm.uxArrowLineStyle.Text.Substring(0, 3))
            Return _arrowLineStyle
        End Get
        Set(ByVal value As Integer)
            _arrowLineStyle = value
        End Set
    End Property


#End Region

#Region "Event Handlers"

    Private Sub PartnerAddArrowsForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles PartnerAddArrowsForm.Load

        With PartnerAddArrowsForm
            'Populate multi-value controls
            If .uxArrowLineStyle.Items.Count = 0 Then '-- Only load the text box the first time the tool is run.
                .uxArrowLineStyle.Items.Add("100 - Anno Arrow")
                .uxArrowLineStyle.Items.Add("101 - Hooks")
                .uxArrowLineStyle.Items.Add("102 - Radius Line")
                .uxArrowLineStyle.Items.Add("120 - Station Reference")
                .uxArrowLineStyle.Items.Add("125 - River Arrow")
                .uxArrowLineStyle.Items.Add("134 - Bearing/Distance Arrow")
                .uxArrowLineStyle.Items.Add("136 - Reference Notes")
                .uxArrowLineStyle.Items.Add("137 - Taxlot Arrow")
                .uxArrowLineStyle.Items.Add("141 - Subdivision Arrow")
                .uxArrowLineStyle.Items.Add("147 - DLC Arrow")
                .uxArrowLineStyle.Items.Add("154 - Code Arrow")
                .uxArrowLineStyle.Items.Add("162 - See Map Arrow")
            End If
            ' Set control defaults
            .uxArrowLineStyle.SelectedIndex = 0

        End With

    End Sub

    Private Sub uxQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerTaxlotAssignmentForm.uxFind.Click

        Dim mapIndexFClass As IFeatureClass = DataMonitor.MapIndexFeatureLayer.FeatureClass
        MsgBox(mapIndexFClass.ShapeFieldName.ToString)

        PartnerAddArrowsForm.Close()
    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerTaxlotAssignmentForm.uxHelp.Click
        ' TODO [SC] Evaluate help systems and implement.
        MessageBox.Show("uxHelp clicked")
    End Sub


    Private Sub uxAddStandard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerTaxlotAssignmentForm.uxAddStandard.Click
        arrowType = "ARROW"
        arrowLineStyle = CInt(PartnerAddArrowsForm.uxArrowLineStyle.Text.Substring(0, 3))
        PartnerAddArrowsForm.Close()
    End Sub

    Private Sub uxAddDimension_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerTaxlotAssignmentForm.uxAddDimension.Click
        arrowType = "DIMENSION"
        PartnerAddArrowsForm.Close()
    End Sub

#End Region

#Region "Methods - Shad needs to Review"

    Friend Sub DoButtonOperation()

        Try
            ' Check for valid data.
            CheckValidMapIndexDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            End If

            ' HACK:

            If _thelArrowPointsCollection Is Nothing Then
                _thelArrowPointsCollection = New Collection
            End If

            PartnerAddArrowsForm.ShowDialog()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Sub

    Friend Function GetSmashedLine(ByVal theDisplay As IScreenDisplay, ByVal theTextSymbol As ISymbol, _
        ByVal thePoint As IPoint, ByVal thePolyline As IPolyline) As IPolyline

        Try

            Dim theBoundary As IPolygon = New Polygon
            theTextSymbol.QueryBoundary(theDisplay.hDC, theDisplay.DisplayTransformation, thePoint, theBoundary)

            Dim theTopoOperator As ITopologicalOperator = DirectCast(theBoundary, ITopologicalOperator)
            Dim pIntersect As IPolyline = DirectCast(theTopoOperator.Intersect(thePolyline, esriGeometryDimension.esriGeometry1Dimension), IPolyline)

            ' Returns the difference between the polyline and the intersection
            theTopoOperator = DirectCast(thePolyline, ITopologicalOperator)

            Return DirectCast(theTopoOperator.Difference(pIntersect), IPolyline)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try

    End Function

    Friend Sub GenerateHooks(ByRef pSketch As IGeometry)

        Try

            ' Initialize the hook angle
            _theHookAngle = 20

            'Make sure the edit sketch is a polyline
            If Not TypeOf pSketch Is IPolyline Then Exit Sub 'SC Perhaps throw an exception instead of exiting.
            Dim pCurve As ICurve = DirectCast(pSketch, ICurve)

            ' Retrieve a reference to the Map Index layer
            Dim pMIFlayer As IFeatureLayer = MapIndexFeatureLayer

            Dim pMIFclass As IFeatureClass = pMIFlayer.FeatureClass

            ' Retrieve the map scale from the overlaying Map Index layer
            Dim vMapScale1 As Object = GetValueViaOverlay(pCurve.FromPoint, pMIFclass, EditorExtension.MapIndexSettings.MapScaleField)
            Dim vMapScale2 As Object = GetValueViaOverlay(pCurve.ToPoint, pMIFclass, EditorExtension.MapIndexSettings.MapScaleField)

            ' Insure that the map scales exist and that they are equal
            If IsDBNull(vMapScale1) Or IsDBNull(vMapScale2) Then
                MsgBox("No mapscale for current MapIndex.  Unable to create hooks", MsgBoxStyle.OkOnly)
                Exit Sub 'SC perhaps throw and error instead.
            End If
            If Not vMapScale1 Is vMapScale2 Then
                MsgBox("Hook can not span Mapindex polygons with different scale", MsgBoxStyle.Critical)
                Exit Sub 'SC perhaps throw and error instead.
            End If

            ' Insures that the map scale is supported -- Not all scales are defined (Issue)
            Dim lLineLength As Integer
            If vMapScale1.Equals(600) Then
                lLineLength = 20
            ElseIf vMapScale1.Equals(1200) Then
                lLineLength = 20
            ElseIf vMapScale1.Equals(2400) Then
                lLineLength = 40
            ElseIf vMapScale1.Equals(4800) Then
                lLineLength = 80
            ElseIf vMapScale1.Equals(24000) Then
                lLineLength = 400
            Else
                MsgBox("Not a valid mapscale.  Unable to create hooks", MsgBoxStyle.OkOnly)
                Exit Sub 'SC perhaps throw and error instead.
            End If
            Dim dHookLength As Double = lLineLength * 0.1

            ' Insures that the polyline only has two vertices (Starting & Ending points only)
            If Not IsSketcha2PointLine(pSketch) Then
                Exit Sub 'SC perhaps throw and error instead.
            End If

            'Get the hook layer
            Dim pHookLayer As IFeatureLayer = FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.CartographicLinesFC)
            If pHookLayer Is Nothing Then
                MsgBox("The layer, " & EditorExtension.TableNamesSettings.CartographicLinesFC & ", is not in the map.", MsgBoxStyle.Exclamation, "Layer not found")
                Exit Sub 'SC perhaps throw and error instead.
            End If
            Dim pHookFC As IFeatureClass = pHookLayer.FeatureClass
            Dim pDSet As IDataset = DirectCast(pHookFC, IDataset)
            Dim pWSEdit As IWorkspaceEdit = DirectCast(pDSet.Workspace, IWorkspaceEdit)

            ' Locate the line type field
            Dim lLineTypeFld As Integer = LocateFields(pHookFC, (EditorExtension.CartographicLinesSettings.LineTypeField))
            If lLineTypeFld = -1 Then Exit Sub 'SC perhaps throw and error instead.

            ' Initialize line objects and collections to create a new line
            Dim pNewPointColl As IPointCollection = New Polyline
            Dim pNormal As ILine = New Line
            Dim pPointColl As IPointCollection = DirectCast(pSketch, IPointCollection)

            ' Adds the head of the hook based on the specified angle and hook length
            Dim dSideA As Double = (dHookLength * System.Math.Sin((360 - _theHookAngle) * (3.14 / 180)))
            Dim dSideC As Double = dHookLength
            Dim dSideB As Double = System.Math.Sqrt((dSideC * dSideC) - (dSideA * dSideA))
            pCurve.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, dSideB, False, dSideA, pNormal)
            pNewPointColl.AddPoint(pNormal.ToPoint)

            ' Adds the line points
            pNewPointColl.AddPoint(pPointColl.Point(0))
            pNewPointColl.AddPoint(pPointColl.Point(1))

            ' Adds the tail of the hook based on the specified angle and hook length
            dSideA = (dHookLength * System.Math.Sin(_theHookAngle * (3.14 / 180)))
            dSideC = dHookLength
            dSideB = System.Math.Sqrt((dSideC * dSideC) - (dSideA * dSideA))
            pCurve.QueryNormal(ESRI.ArcGIS.Geometry.esriSegmentExtension.esriNoExtension, (pCurve.Length - dSideB), False, dSideA, pNormal)
            pNewPointColl.AddPoint(pNormal.ToPoint)

            'Now get rid of the line between the start and end points (where user clicked)
            Dim pWholeLine As IPolyline4 = DirectCast(pNewPointColl, IPolyline4)
            Dim bBool As Boolean
            Dim bSplitHappened As Boolean
            Dim lNewPartIndex As Integer
            Dim lNewSegIndex As Integer
            pWholeLine.SplitAtPoint(_theFromBreakPoint, True, bBool, bSplitHappened, lNewPartIndex, lNewSegIndex)
            pWholeLine.SplitAtPoint(_theToBreakPoint, True, bBool, bSplitHappened, lNewPartIndex, lNewSegIndex)

            ' Initialize new path objects and collections to create a new polyline
            Dim pPath1 As ISegmentCollection = New Path
            Dim pPath2 As ISegmentCollection = New Path
            Dim pPath3 As ISegmentCollection = New Path

            ' QI to get the segment collection of the landhook
            Dim pSegCollection As ISegmentCollection = DirectCast(pWholeLine, ISegmentCollection)

            ' Retreive an enumeration of the segments
            Dim pEnumSeg As IEnumSegment = pSegCollection.EnumSegments

            ' Add segments to the paths that will make the final land hook
            Dim pSeg As ISegment = Nothing
            Dim lPartIndex As Integer
            Dim lSegIndex As Integer
            pEnumSeg.Next(pSeg, lPartIndex, lSegIndex)
            Do While Not pSeg Is Nothing
                If lSegIndex < 1 Then
                    pPath1.AddSegment(pSeg)
                ElseIf lSegIndex = 1 Then
                    pPath2.AddSegment(pSeg)
                ElseIf lSegIndex = 2 Then
                    pPath3.AddSegment(pSeg)
                End If
                pEnumSeg.Next(pSeg, lPartIndex, lSegIndex)
            Loop

            ' Add the component paths to the final land hook
            Dim pGeomColl As IGeometryCollection = New Polyline
            pGeomColl.AddGeometry(DirectCast(pPath1, IGeometry))
            pGeomColl.AddGeometry(DirectCast(pPath2, IGeometry))
            pGeomColl.AddGeometry(DirectCast(pPath3, IGeometry))
            pGeomColl.GeometriesChanged()

            'Store the new land hook feature
            pWSEdit.StartEditOperation()
            Dim pFeature As IFeature = pHookFC.CreateFeature
            pFeature.Shape = DirectCast(pGeomColl, IGeometry)
            pFeature.Value(lLineTypeFld) = 101

            'Set the AutoMethod Field
            lLineTypeFld = LocateFields(pHookFC, (EditorExtension.AllTablesSettings.AutoMethodField))
            If lLineTypeFld = -1 Then Exit Sub
            pFeature.Value(lLineTypeFld) = "UNK"

            'Set the AutoWho Field
            lLineTypeFld = LocateFields(pHookFC, (EditorExtension.AllTablesSettings.AutoWhoField))
            If lLineTypeFld = -1 Then Exit Sub
            pFeature.Value(lLineTypeFld) = UserName

            'Set the AutoDate Field
            lLineTypeFld = LocateFields(pHookFC, (EditorExtension.AllTablesSettings.AutoDateField))
            If lLineTypeFld = -1 Then Exit Sub
            pFeature.Value(lLineTypeFld) = Format(Today, "MM/dd/yyyy")

            'Set the MapScale Field
            lLineTypeFld = LocateFields(pHookFC, (EditorExtension.MapIndexSettings.MapScaleField))
            If lLineTypeFld = -1 Then Exit Sub
            pFeature.Value(lLineTypeFld) = vMapScale1

            'Set the MapNumber Field
            Dim sCurMapNum As String = GetValueViaOverlay(pFeature.Shape, pMIFclass, EditorExtension.MapIndexSettings.MapNumberField)
            lLineTypeFld = LocateFields(pHookFC, (EditorExtension.MapIndexSettings.MapNumberField))
            If lLineTypeFld = -1 Then Exit Sub
            pFeature.Value(lLineTypeFld) = sCurMapNum

            pFeature.Store()
            pWSEdit.StopEditOperation()

            Dim theArcMapDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
            Dim pActiveView As IActiveView = theArcMapDoc.ActiveView
            pActiveView.PartialRefresh(esriViewDrawPhase.esriViewBackground, Nothing, pFeature.Extent.Envelope)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Public Function IsSketcha2PointLine(ByRef pGeom As IGeometry) As Boolean
        Try
            ' Validate the passed geometry
            If Not TypeOf pGeom Is IPointCollection Then Exit Function

            ' Validate the number of points in the collection
            Dim pPointColl As IPointCollection = DirectCast(pGeom, IPointCollection)
            If pPointColl.PointCount <> 2 Then
                MessageBox.Show("When creating the parcel hook only digitize two points", "Parcel Hook Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False
        End Try

    End Function

    Private Function ReturnExtended(ByRef pExt As esriSegmentExtension, ByRef pPolyline As IPolyline, ByRef lLength As Integer) As IPolyline

        Try
            Dim pCurve As ICurve = pPolyline
            Dim pLine As ILine = New ESRI.ArcGIS.Geometry.Line
            Dim pPLine As IPolyline = New Polyline
            pCurve.QueryTangent(pExt, 1, False, lLength, pLine)

            'Convert ILine to an IPolyline
            pPLine.FromPoint = pLine.FromPoint
            pPLine.ToPoint = pLine.ToPoint
            Return pPLine

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try

    End Function

    Private Sub DrawArrows()

        Try
            'Set up line symbol to display temporary line
            _theLineSymbol = New SimpleLineSymbol
            _theLineSymbol.Width = 2
            Dim pRGBColor As IRgbColor = New RgbColor
            With pRGBColor
                .Red = 223
                .Green = 223
                .Blue = 223
            End With

            'UPGRADE_WARNING: Couldn't resolve default property of object pRGBColor. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            _theLineSymbol.Color = pRGBColor
            Dim pSymbol As ISymbol = DirectCast(_theLineSymbol, ISymbol)
            pSymbol.ROP2 = esriRasterOpCode.esriROPXOrPen

            ' Create the polyline from a point collection
            Dim pArrowLine As IPointCollection4 = New Polyline


            If arrowType.Equals("Arrow", _ignoreCase) Then
                If _thelArrowPointsCollection.Count() > 1 Then
                    For lngIndex As Integer = 1 To _thelArrowPointsCollection.Count()
                        pArrowLine.AddPoint(DirectCast(_thelArrowPointsCollection.Item(lngIndex), IPoint))
                    Next
                End If
            Else
                If Not _theArrowPt1 Is Nothing Then pArrowLine.AddPoint(_theArrowPt1)
                If Not _theArrowPt2 Is Nothing Then pArrowLine.AddPoint(_theArrowPt2)
                If Not _theArrowPt3 Is Nothing Then pArrowLine.AddPoint(_theArrowPt3)
                If Not _theArrowPt4 Is Nothing Then pArrowLine.AddPoint(_theArrowPt4)
            End If

            Dim pArrowLine2 As IPolyline4 = DirectCast(pArrowLine, IPolyline4)

            ' Draw the temporary line
            Dim theArcMapDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
            Dim pActiveView As IActiveView = theArcMapDoc.ActiveView

            pActiveView.ScreenDisplay.SetSymbol(DirectCast(_theLineSymbol, ISymbol))
            If (pArrowLine2.Length > 0) Then pActiveView.ScreenDisplay.DrawPolyline(pArrowLine2)
            pActiveView.ScreenDisplay.FinishDrawing()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub GetMousePoint(ByRef X As Integer, ByRef Y As Integer)

        Try
            '+++ Get the current map point (and invert the agent at that location)
            _thePt = _theActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)

            '+++ get the snap agent, if it is being used
            _theSnapAgent = Nothing

            Dim pSnapenv As ISnapEnvironment = DirectCast(EditorExtension.Editor, ISnapEnvironment)

            Dim pSnapAgent As ISnapAgent
            Dim pFSnapAgent As IFeatureSnapAgent
            Dim pEdLyrs As IEditLayers
            Dim pLayer As ILayer
            Dim ht As esriGeometryHitPartType

            For i As Integer = 0 To pSnapenv.SnapAgentCount - 1
                pSnapAgent = pSnapenv.SnapAgent(i)
                If TypeOf pSnapAgent Is ESRI.ArcGIS.Editor.IFeatureSnapAgent Then
                    pFSnapAgent = DirectCast(pSnapAgent, IFeatureSnapAgent)
                    pEdLyrs = DirectCast(EditorExtension.Editor, IEditLayers)
                    pLayer = pEdLyrs.CurrentLayer
                    ht = pFSnapAgent.HitType
                    If ht <> 0 Then
                        _theSnapAgent = pFSnapAgent
                        Exit For
                    End If
                    pLayer = Nothing
                    pEdLyrs = Nothing
                    pFSnapAgent = Nothing
                End If
            Next i

            pSnapAgent = Nothing
            pSnapenv = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Sub

    Private Function GetCurrentMapScale(ByRef pMIFC As IFeatureClass) As String

        Try
            Dim pDimensionArrowLayerTemp As IFeatureLayer = FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.CartographicLinesFC)
            If pDimensionArrowLayerTemp Is Nothing Then
                MsgBox("The layer, " & EditorExtension.TableNamesSettings.CartographicLinesFC & ", is not in the map.", MsgBoxStyle.Exclamation, "Layer not found")
                m_pEditor.AbortOperation()
                Return Nothing
                Exit Function
            End If

            Dim pDimensionArrowFCTemp As IFeatureClass = pDimensionArrowLayerTemp.FeatureClass
            Dim pDimensionDSetTemp As IDataset = DirectCast(pDimensionArrowFCTemp, IDataset)
            Dim pDimensionWSEditTemp As IWorkspaceEdit = DirectCast(pDimensionDSetTemp.Workspace, IWorkspaceEdit)

            'create the arrow feature
            pDimensionWSEditTemp.StartEditOperation()
            Dim pDimensionFeatureTemp As IFeature = pDimensionArrowFCTemp.CreateFeature
            Dim pDimensionpointsTemp As ESRI.ArcGIS.Geometry.IPointCollection4
            pDimensionpointsTemp = New ESRI.ArcGIS.Geometry.Polyline
            pDimensionpointsTemp.AddPoint(_theArrowPt1)
            pDimensionpointsTemp.AddPoint(_theArrowPt2)

            Dim pDimensionLineTemp As ESRI.ArcGIS.Geometry.IPolyline
            pDimensionLineTemp = DirectCast(pDimensionpointsTemp, IPolyline)
            pDimensionFeatureTemp.Shape = pDimensionLineTemp


            'Get the current MapNumber
            Dim sCurrentMapScale As String
            sCurrentMapScale = GetValueViaOverlay((pDimensionFeatureTemp.Shape), pMIFC, EditorExtension.MapIndexSettings.MapScaleField)
            Return CStr(CDbl(sCurrentMapScale) / 12)

            pDimensionWSEditTemp.AbortEditOperation()
            pDimensionWSEditTemp.StopEditOperation()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing

        End Try

    End Function

    Private Function GetDimensionArrowSide() As String

        Try
            'Determine point location is on the left or right by Dean Anderson, help of Nate Anderson
            Dim slope As Double = 0
            If _theArrowPt2.Y <> _theArrowPt1.Y Then
                slope = (_theArrowPt1.Y - _theArrowPt2.Y) / (_theArrowPt1.X - _theArrowPt2.X)
            End If

            Dim yint As Double = _theArrowPt1.Y - (slope * _theArrowPt1.X)
            Dim z As Double = (slope * _theArrowPt3.X) + yint - _theArrowPt3.Y

            If _theArrowPt1.X = _theArrowPt2.X Then 'vertical
                If _theArrowPt3.X = _theArrowPt1.X Then z = 0
                If _theArrowPt1.Y < _theArrowPt2.Y Then 'going up
                    If _theArrowPt3.X > _theArrowPt1.X Then z = 1
                    If _theArrowPt3.X < _theArrowPt1.X Then z = -1
                End If
                If _theArrowPt1.Y > _theArrowPt2.Y Then 'going down
                    If _theArrowPt3.X < _theArrowPt1.X Then z = 1
                    If _theArrowPt3.X > _theArrowPt1.X Then z = -1
                End If
            End If
            If _theArrowPt1.Y = _theArrowPt2.Y Then 'horizontal
                If _theArrowPt3.Y = _theArrowPt1.Y Then z = 0
                If _theArrowPt1.X < _theArrowPt2.X Then 'going right
                    If _theArrowPt3.Y > _theArrowPt1.Y Then z = -1
                    If _theArrowPt3.Y < _theArrowPt1.Y Then z = 1
                End If
                If _theArrowPt1.X > _theArrowPt2.X Then 'going left
                    If _theArrowPt3.Y < _theArrowPt1.Y Then z = -1
                    If _theArrowPt3.Y > _theArrowPt1.Y Then z = 1
                End If
            End If

            Dim dimensionArrowSide As String = ""
            If z < 0 Then dimensionArrowSide = "left"
            If z > 0 Then dimensionArrowSide = "right"
            If z = 0 Then dimensionArrowSide = "left" '"online"

            If (_theArrowPt1.X > _theArrowPt2.X And _theArrowPt1.Y > _theArrowPt2.Y) Or (_theArrowPt1.X > _theArrowPt2.X And _theArrowPt2.Y > _theArrowPt1.Y) Then
                If z > 0 Then dimensionArrowSide = "left"
                If z < 0 Then dimensionArrowSide = "right"
                If z = 0 Then dimensionArrowSide = "left" '"online"
            End If

            Return dimensionArrowSide

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing

        End Try

    End Function

    Private Function GetChange(ByRef sCurrentMapScale As String, ByVal Shift As Integer) As Short

        Try

            Dim theChange As Short = Nothing

            If sCurrentMapScale = "100" Then
                theChange = 15
            ElseIf sCurrentMapScale = "200" Then
                theChange = 30
            ElseIf sCurrentMapScale = "400" Then
                theChange = 60
            ElseIf sCurrentMapScale = "2000" Then
                theChange = 300
            End If

            Return theChange

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing

        End Try

    End Function


    Public Function ConvertToDescription(ByVal pFlds As IFields, ByVal sFldName As String, ByVal vVal As String) As String

        Try

            Dim lFld As Integer = pFlds.FindField(sFldName)
            If lFld > -1 Then
                'Determine if domain field
                Dim pField As IField = pFlds.Field(lFld)
                Dim pDomain As IDomain = pField.Domain
                If pDomain Is Nothing Then
                    Return vVal
                    Exit Function
                Else
                    'Determine type of domain  -If Coded Value, get the description
                    If TypeOf pDomain Is ICodedValueDomain Then
                        Dim pCVDomain As ICodedValueDomain = DirectCast(pDomain, ICodedValueDomain)
                        'Given the description, search the domain for the code
                        For i As Integer = 0 To pCVDomain.CodeCount - 1
                            If pCVDomain.Value(i).Equals(vVal) Then
                                ConvertToDescription = pCVDomain.Name(i) 'Return the code value
                                Exit Function
                            End If
                        Next i
                    Else ' If range domain, return the numeric value
                        Return vVal
                        Exit Function
                    End If
                End If  'If pDomain is nothing/Else
                Return vVal
            Else
                'Field not found
                Return ""
            End If 'If lFld > -1/Else

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing

        End Try

        '++ END JWalton 2/7/2007
    End Function

#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"

    ''' <summary>
    ''' Called by ArcMap once per second to check if the command is enabled.
    ''' </summary>
    ''' <remarks>WARNING: Do not put computation-intensive code here.</remarks>
    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Try
                Dim canEnable As Boolean
                canEnable = EditorExtension.CanEnableExtendedEditing
                Return canEnable
            Catch ex As Exception
                EditorExtension.ProcessUnhandledException(ex)
            End Try
        End Get
    End Property

#End Region

#Region "Methods"

    Public Overrides Sub OnCreate(ByVal hook As Object)
        Try
            If Not hook Is Nothing Then

                'Disable tool if parent application is not ArcMap
                If TypeOf hook Is IMxApplication Then
                    _application = DirectCast(hook, IApplication)
                    setPartnerAddArrowsForm(New AddArrowsForm())
                    MyBase.m_enabled = True
                Else
                    MyBase.m_enabled = False
                    'Disable tool if parent application is not ArcMap
                    If TypeOf hook Is IMxApplication Then
                        _application = DirectCast(hook, IApplication)
                        ' TODO: [SC] Create the form, property and set__ procedure. (Nick)
                        'setPartnerAddArrowsForm(New AddArrowsForm())
                        MyBase.m_enabled = True
                    Else
                        MyBase.m_enabled = False
                    End If
                End If

            End If
            ' NOTE: Add other initialization code here...

        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try

    End Sub

    Public Overrides Sub OnClick()
        Try
            DoButtonOperation()
        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)

        If PartnerAddArrowsForm.Visible Then Exit Sub

        Try

            Dim lLineTypeFld As Integer
            Dim sCurrentMapScale As String
            Dim dSmoothRatio As Double
            Dim bolFinish As Boolean
            Dim lngIndex As Long

            ' Set the in use flag
            _theInUse = True

            Dim theArcMapDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
            Dim pActiveView As IActiveView = theArcMapDoc.ActiveView

            ' Retrieve a reference to the Map Index layer
            'Dim pMIFL As IFeatureLayer = MapIndexFeatureLayer
            Dim pMIFC As IFeatureClass = MapIndexFeatureLayer.FeatureClass

            Select Case arrowType

                Case "HOOK"  '"Hook" 'If drawing hooks
                    ' Get point to measure distance from
                    _theStartPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                    _theDoOnce = False 'testing
                    _theFromBreakPoint = _theStartPoint

                    ' Get the scale of the current mapindex
                    Dim sScale As String = GetValueViaOverlay(_theStartPoint, pMIFC, EditorExtension.MapIndexSettings.MapScaleField)
                    sScale = ConvertToDescription(pMIFC.Fields, EditorExtension.MapIndexSettings.MapScaleField, sScale)


                Case "ARROW" 'If drawing annotation arrows
                    'If m_bToolJustCompletedTask = False Then
                    If bolFinish = False Then
                        'Right mouse click
                        If Button = 1 And Shift = 1 Then
                            'Save the first point
                            If _thelArrowPointsCollection.Count = 0 Then
                                If _theArrowPt1 Is Nothing Then
                                    _theArrowPt1 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                                    _thelArrowPointsCollection.Add(_theArrowPt1)

                                    'Clear existing point
                                    _theArrowPt1 = Nothing
                                End If

                                'Save the last point
                            Else
                                If _theArrowPt1 Is Nothing Then
                                    _theArrowPt1 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                                    _thelArrowPointsCollection.Add(_theArrowPt1)
                                    DrawArrows()

                                    'Clear existing point
                                    _theArrowPt1 = Nothing
                                    bolFinish = True
                                End If
                            End If

                            'Left mouse click
                        ElseIf Button = 1 And Shift = 0 Then
                            'Add vertex
                            If _theArrowPt1 Is Nothing Then
                                _theArrowPt1 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                                _thelArrowPointsCollection.Add(_theArrowPt1)

                                'Clear existing point
                                _theArrowPt1 = Nothing
                            End If
                        End If
                    End If


                    If _thelArrowPointsCollection.Count > 1 And bolFinish = True Then

                        ' Creates a new polygon from the points and smoothes it
                        Dim pArrowPoints As IPointCollection4 = New Polyline

                        For lngIndex = 1 To _thelArrowPointsCollection.Count
                            pArrowPoints.AddPoint(DirectCast(_thelArrowPointsCollection.Item(lngIndex), IPoint))
                        Next

                        Dim pArrowLine As IPolyline4 = DirectCast(pArrowPoints, IPolyline4)
                        pArrowLine.Smooth(pArrowLine.Length / 10)

                        ' Get a reference to the Cartographic Lines feature class
                        Dim pArrowLayer As IFeatureLayer = FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.CartographicLinesFC)
                        If pArrowLayer Is Nothing Then
                            MsgBox("The layer, " & EditorExtension.TableNamesSettings.CartographicLinesFC & ", is not in the map.", vbExclamation, "Layer not found")
                            Exit Sub 'SC better method?
                        End If
                        Dim pArrowFC As IFeatureClass = pArrowLayer.FeatureClass
                        Dim pDSet As IDataset = DirectCast(pArrowFC, IDataset)
                        Dim pWSEdit As IWorkspaceEdit = DirectCast(pDSet.Workspace, IWorkspaceEdit)

                        ' Start an edit operation to encompass the creation of the feature
                        pWSEdit.StartEditOperation()

                        ' Create the arrow feature
                        Dim pFeature As IFeature = pArrowFC.CreateFeature
                        pFeature.Shape = pArrowLine

                        ' Locates fields in the feature's dataset
                        Dim lCLMNfld As Integer = LocateFields(pArrowFC, EditorExtension.MapIndexSettings.MapNumberField)
                        lLineTypeFld = LocateFields(pArrowFC, EditorExtension.CartographicLinesSettings.LineTypeField)

                        ' Insure that the feature's fields are found
                        If lLineTypeFld = -1 Then Exit Sub 'SC better method?

                        ' Populate field values in the feature
                        Dim sCurMapNum As String = GetValueViaOverlay(pFeature.Shape, pMIFC, EditorExtension.MapIndexSettings.MapNumberField)
                        pFeature.Value(lCLMNfld) = sCurMapNum
                        pFeature.Value(lLineTypeFld) = arrowLineStyle
                        pFeature.Store()

                        'Set the AutoMethod Field
                        lLineTypeFld = LocateFields(pArrowFC, EditorExtension.AllTablesSettings.AutoMethodField)
                        If lLineTypeFld = -1 Then Exit Sub
                        pFeature.Value(lLineTypeFld) = "UNK"

                        'Set the AutoWho Field
                        lLineTypeFld = LocateFields(pArrowFC, EditorExtension.AllTablesSettings.AutoWhoField)
                        If lLineTypeFld = -1 Then Exit Sub
                        pFeature.Value(lLineTypeFld) = UserName

                        'Set the AutoDate Field
                        lLineTypeFld = LocateFields(pArrowFC, EditorExtension.AllTablesSettings.AutoDateField)
                        If lLineTypeFld = -1 Then Exit Sub
                        pFeature.Value(lLineTypeFld) = Format(Today, "MM/dd/yyyy")

                        'Set the MapScale Field
                        sCurrentMapScale = GetValueViaOverlay(pFeature.Shape, pMIFC, EditorExtension.MapIndexSettings.MapScaleField)
                        lLineTypeFld = LocateFields(pArrowFC, EditorExtension.MapIndexSettings.MapScaleField)
                        If lLineTypeFld = -1 Then Exit Sub
                        pFeature.Value(lLineTypeFld) = sCurrentMapScale

                        ' Finalize the edit operation
                        pWSEdit.StopEditOperation()

                        ' Refresh the display
                        pActiveView.PartialRefresh(esriViewDrawPhase.esriViewBackground, Nothing, pFeature.Extent.Envelope)

                        _thelArrowPointsCollection = Nothing
                        _thelArrowPointsCollection = New Collection


                    End If

                Case "DIMENSION"

                    If _theArrowPt1 Is Nothing Then
                        _theArrowPt1 = _thePt
                    ElseIf _theArrowPt2 Is Nothing Then
                        _theArrowPt2 = _thePt
                        DrawArrows()
                    ElseIf _theArrowPt3 Is Nothing Then
                        _theArrowPt3 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                        DrawArrows()

                        'Check to be sure the 2 points are not equal
                        If (_theArrowPt1.X = _theArrowPt2.X) And (_theArrowPt1.Y = _theArrowPt2.Y) Then
                            MsgBox("Two input points can't be equal, dimension arrows terminated.", vbInformation, "Invalid Input")
                            OnKeyDown(Keys.Q, 1)
                            Exit Sub
                        End If

                        'Get the mapscale
                        sCurrentMapScale = GetCurrentMapScale(pMIFC)

                        'Get the side the dimension arrows should be on based upon the 3rd input point
                        Dim sSide As String = GetDimensionArrowSide()

                        'Create two dimension arrows, one on the left and one on the right
                        'Dim lngIndex As Long
                        For lngIndex = 1 To 2

                            Dim pDimensionLine As IPolyline4
                            Dim pDimensionPoints As IPointCollection4
                            pDimensionPoints = New Polyline

                            If _addManually = False Then

                                'Add starting point
                                pDimensionPoints.AddPoint(_theArrowPt1)

                                'Instanciate the temp points
                                _theArrowPtTemp = New ESRI.ArcGIS.Geometry.Point
                                _theArrowPtTemp2 = New ESRI.ArcGIS.Geometry.Point

                                'Set the distance of change for the dimension arrows based upon the mapscale
                                Dim iChange As Integer = GetChange(sCurrentMapScale, Shift)

                                'Get 3 calculated points for a hook
                                'Create the line from input
                                pDimensionPoints.AddPoint(_theArrowPt1)
                                pDimensionPoints.AddPoint(_theArrowPt2)
                                pDimensionLine = DirectCast(pDimensionPoints, IPolyline4)

                                'Check the ratio changes
                                Dim dRatioLine As Double
                                Dim dRatioCurve As Double

                                If _theolDimensionChanged = True Then
                                    If _ratioLine > 0 Then
                                        dRatioLine = _ratioLine
                                    Else
                                        dRatioLine = 1.75
                                    End If

                                    If _ratioCurve > 0 Then
                                        dRatioCurve = _ratioCurve
                                    Else
                                        dRatioCurve = 1.35
                                    End If

                                    If _smoothRatio >= 0 Then
                                        dSmoothRatio = _smoothRatio
                                    Else
                                        dSmoothRatio = 10
                                    End If
                                Else
                                    dRatioLine = 1.75
                                    dRatioCurve = 1.35
                                    dSmoothRatio = 10
                                End If

                                'Create a line iChange from the beginning
                                Dim pLine As ILine = New Line
                                If lngIndex = 1 Then
                                    If Shift = 0 Then 'shift not pressed
                                        If Trim$(UCase$(sSide)) = "RIGHT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, iChange, False, ((iChange / dRatioLine) / 2), pLine)
                                        ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, iChange, False, -((iChange / dRatioLine) / 2), pLine)
                                        End If
                                    ElseIf Shift = 1 Then 'shift pressed
                                        If Trim$(UCase$(sSide)) = "RIGHT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, iChange, False, (((iChange * 2) / dRatioLine) / 2), pLine)
                                        ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, iChange, False, -(((iChange * 2) / dRatioLine) / 2), pLine)
                                        End If
                                    End If
                                ElseIf lngIndex = 2 Then
                                    If Shift = 0 Then 'shift not pressed
                                        If Trim$(UCase$(sSide)) = "RIGHT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (pDimensionLine.Length - iChange), False, ((iChange / dRatioLine) / 2), pLine)
                                        ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (pDimensionLine.Length - iChange), False, -((iChange / dRatioLine) / 2), pLine)
                                        End If
                                    ElseIf Shift = 1 Then 'shift pressed
                                        If Trim$(UCase$(sSide)) = "RIGHT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (pDimensionLine.Length - iChange), False, (((iChange * 2) / dRatioLine) / 2), pLine)
                                        ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (pDimensionLine.Length - iChange), False, -(((iChange * 2) / dRatioLine) / 2), pLine)

                                        End If
                                    End If
                                End If

                                'Save the to and from points of the line
                                _theArrowPtTemp.X = pLine.ToPoint.X
                                _theArrowPtTemp.Y = pLine.ToPoint.Y
                                pLine = Nothing

                                'Create a line (iChange/.25) from the beginning
                                pLine = New Line
                                If lngIndex = 1 Then
                                    If Shift = 0 Then 'shift not pressed
                                        If Trim$(UCase$(sSide)) = "RIGHT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (iChange / dRatioCurve), False, ((iChange / dRatioLine) / 1.75), pLine)
                                        ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (iChange / dRatioCurve), False, -((iChange / dRatioLine) / 1.75), pLine)
                                        End If
                                    ElseIf Shift = 1 Then 'shift pressed
                                        If Trim$(UCase$(sSide)) = "RIGHT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (iChange / dRatioCurve), False, (((iChange * 2) / dRatioLine) / 1.75), pLine)
                                        ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (iChange / dRatioCurve), False, -(((iChange * 2) / dRatioLine) / 1.75), pLine)
                                        End If
                                    End If
                                ElseIf lngIndex = 2 Then
                                    If Shift = 0 Then 'shift not pressed
                                        If Trim$(UCase$(sSide)) = "RIGHT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, ((iChange / dRatioLine) / 1.75), pLine)
                                        ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, -((iChange / dRatioLine) / 1.75), pLine)
                                        End If
                                    ElseIf Shift = 1 Then 'shift pressed
                                        If Trim$(UCase$(sSide)) = "RIGHT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, (((iChange * 2) / dRatioLine) / 1.75), pLine)
                                        ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                            pDimensionLine.QueryNormal(esriSegmentExtension.esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, -(((iChange * 2) / dRatioLine) / 1.75), pLine)
                                        End If
                                    End If
                                End If

                                'Save the to and from points of the line
                                _theArrowPtTemp2.X = pLine.ToPoint.X
                                _theArrowPtTemp2.Y = pLine.ToPoint.Y
                                pLine = Nothing

                                pDimensionPoints = Nothing
                                pDimensionPoints = New Polyline

                                'Add the points to the line/hook to be created
                                If lngIndex = 1 Then
                                    pDimensionPoints.AddPoint(_theArrowPt1)
                                ElseIf lngIndex = 2 Then
                                    pDimensionPoints.AddPoint(_theArrowPt2)
                                End If
                                pDimensionPoints.AddPoint(_theArrowPtTemp2)
                                pDimensionPoints.AddPoint(_theArrowPtTemp)

                            Else 'Adding dimension arrow manually
                                lngIndex = 3
                                pDimensionPoints.AddPoint(_theArrowPt1)
                                pDimensionPoints.AddPoint(_theArrowPt2)
                                pDimensionPoints.AddPoint(_theArrowPt3)

                                If _smoothRatio >= 0 Then
                                    dSmoothRatio = _smoothRatio
                                Else
                                    dSmoothRatio = 10
                                End If
                            End If

                            'Create the dimension arrow from the collection of 3 points (1 input, 2 calculated)
                            pDimensionLine = DirectCast(pDimensionPoints, IPolyline4)

                            If dSmoothRatio > 0 Then
                                pDimensionLine.Smooth(pDimensionLine.Length / dSmoothRatio)
                            End If

                            Dim pDimensionFeature As IFeature
                            Dim pDimensionWSEdit As IWorkspaceEdit
                            Dim pDimensionDSet As IDataset
                            Dim pDimensionArrowLayer As IFeatureLayer
                            Dim pDimensionArrowFC As IFeatureClass
                            pDimensionArrowLayer = FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.CartographicLinesFC)
                            If pDimensionArrowLayer Is Nothing Then
                                MsgBox("The layer, " & EditorExtension.TableNamesSettings.CartographicLinesFC & ", is not in the map.", vbExclamation, "Layer not found")
                                m_pEditor.AbortOperation()
                                Exit Sub
                            End If
                            pDimensionArrowFC = pDimensionArrowLayer.FeatureClass
                            pDimensionDSet = DirectCast(pDimensionArrowFC, IDataset)
                            pDimensionWSEdit = DirectCast(pDimensionDSet.Workspace, IWorkspaceEdit)

                            'create the arrow feature
                            pDimensionWSEdit.StartEditOperation()
                            pDimensionFeature = pDimensionArrowFC.CreateFeature
                            pDimensionFeature.Shape = pDimensionLine
                            lLineTypeFld = LocateFields(pDimensionArrowFC, EditorExtension.CartographicLinesSettings.LineTypeField)
                            If lLineTypeFld = -1 Then Exit Sub
                            pDimensionFeature.Value(lLineTypeFld) = 134 'Bearing Distance Arrow

                            'Get the current MapNumber
                            Dim sCurrentMapNums As String
                            sCurrentMapNums = GetValueViaOverlay(pDimensionFeature.Shape, pMIFC, EditorExtension.MapIndexSettings.MapNumberField)
                            Dim lCLMNSfld As Integer
                            lCLMNSfld = LocateFields(pDimensionArrowFC, EditorExtension.MapIndexSettings.MapNumberField)
                            pDimensionFeature.Value(lCLMNSfld) = sCurrentMapNums

                            'Set the AutoMethod Field
                            lLineTypeFld = LocateFields(pDimensionArrowFC, EditorExtension.AllTablesSettings.AutoMethodField)
                            If lLineTypeFld = -1 Then Exit Sub
                            pDimensionFeature.Value(lLineTypeFld) = "UNK"

                            'Set the AutoWho Field
                            lLineTypeFld = LocateFields(pDimensionArrowFC, EditorExtension.AllTablesSettings.AutoWhoField)
                            If lLineTypeFld = -1 Then Exit Sub
                            pDimensionFeature.Value(lLineTypeFld) = UserName

                            'Set the AutoDate Field
                            lLineTypeFld = LocateFields(pDimensionArrowFC, EditorExtension.AllTablesSettings.AutoDateField)
                            If lLineTypeFld = -1 Then Exit Sub
                            pDimensionFeature.Value(lLineTypeFld) = Format(Today, "MM/dd/yyyy")

                            'Set the MapScale Field
                            lLineTypeFld = LocateFields(pDimensionArrowFC, EditorExtension.MapIndexSettings.MapScaleField)
                            If lLineTypeFld = -1 Then Exit Sub
                            pDimensionFeature.Value(lLineTypeFld) = CInt(sCurrentMapScale) * 12

                            pDimensionFeature.Store()
                            pDimensionWSEdit.StopEditOperation()

                            'Refresh the display
                            pActiveView.PartialRefresh(esriViewDrawPhase.esriViewBackground, Nothing, pDimensionFeature.Extent.Envelope)

                            _theArrowPtTemp = Nothing
                            _theArrowPtTemp2 = Nothing

                            _theTextSymbol = Nothing
                            _theTextPoint = Nothing
                            _theLinePolyline = Nothing
                            _theLineSymbol = Nothing

                        Next

                        _theArrowPt1 = Nothing
                        _theArrowPt2 = Nothing
                        _theArrowPt3 = Nothing
                        _theArrowPt4 = Nothing

                    Else 'Reset everything
                        _theArrowPt1 = Nothing
                        _theArrowPt2 = Nothing
                        _theArrowPt3 = Nothing
                        _theArrowPt4 = Nothing
                        _theArrowPtTemp = Nothing
                        _theArrowPtTemp2 = Nothing
                        _theSnapAgent = Nothing
                        _theToolJustCompletedTask = True

                        'Deactivate the tool
                        _application.CurrentTool = Nothing

                    End If
            End Select


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        MyBase.OnMouseMove(Button, Shift, X, Y)


        Try

            Dim theArcMapDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
            Dim pActiveView As IActiveView = theArcMapDoc.ActiveView
            Dim bfirstTime As Boolean
            Dim lLineLength As Long = 0 'TODO - SC This does not get set.

            ' Draw the temporary line the user sees while moving the mouse
            If arrowType.Equals("Hook", _ignoreCase) Then

                If (Not _theInUse) Then Exit Sub ' SC perhaps exit try instead??

                ' Checks to see if the line symbol is defined
                If (_theLineSymbol Is Nothing) Then bfirstTime = True

                ' Get current point
                Dim pPoint As IPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                _theToBreakPoint = pPoint 'the unextended to point used to break the hooks

                'Check to be sure there is a start point for the hook to prevent an error
                If _theStartPoint Is Nothing Then Exit Sub

                ' Draw a virtual line that represents the extended line
                Dim pPLine As IPolyline = New Polyline
                pPLine.FromPoint = _theStartPoint
                pPLine.ToPoint = pPoint
                Dim pCv As ICurve = pPLine

                Dim pCPoint As IConstructPoint = New ESRI.ArcGIS.Geometry.Point
                pCPoint.ConstructAlong(pCv, esriSegmentExtension.esriExtendAtTo, pCv.Length + CDbl(lLineLength), False)
                pPoint = DirectCast(pCPoint, IPoint)

                If Not _theDoOnce Then
                    Dim pFCPoint As IConstructPoint = New ESRI.ArcGIS.Geometry.Point
                    pFCPoint.ConstructAlong(pCv, esriSegmentExtension.esriExtendAtFrom, -(CDbl(lLineLength)), False)
                    _theStartPoint = DirectCast(pFCPoint, IPoint)
                    _theDoOnce = True
                End If

                'Draw the line
                pActiveView.ScreenDisplay.StartDrawing(pActiveView.ScreenDisplay.hDC, -1)

                ' Initialize or draw the temporary line
                If bfirstTime Then
                    'Line Symbol
                    _theLineSymbol = New SimpleLineSymbol
                    _theLineSymbol.Width = 2
                    Dim pRGBColor As IRgbColor = New RgbColor
                    With pRGBColor
                        .Red = 223
                        .Green = 223
                        .Blue = 223
                    End With
                    _theLineSymbol.Color = pRGBColor
                    Dim pSymbol As ISymbol = DirectCast(_theLineSymbol, ISymbol)
                    pSymbol.ROP2 = esriRasterOpCode.esriROPXOrPen

                    'Text Symbol
                    _theTextSymbol = New ESRI.ArcGIS.Display.TextSymbol
                    _theTextSymbol.HorizontalAlignment = esriTextHorizontalAlignment.esriTHACenter
                    _theTextSymbol.VerticalAlignment = esriTextVerticalAlignment.esriTVACenter
                    _theTextSymbol.Size = 16
                    _theTextSymbol.Font.Name = "Arial"

                    pSymbol = DirectCast(_theTextSymbol, ISymbol)
                    Dim pFont As Font = DirectCast(_theTextSymbol.Font, Font)
                    pSymbol.ROP2 = esriRasterOpCode.esriROPXOrPen

                    'Create point to draw text in
                    _theTextPoint = New ESRI.ArcGIS.Geometry.Point
                Else
                    'Use existing symbols and draw existing text and polyline
                    pActiveView.ScreenDisplay.SetSymbol(DirectCast(_theTextSymbol, ISymbol))
                    pActiveView.ScreenDisplay.DrawText(_theTextPoint, _theTextSymbol.Text)
                    pActiveView.ScreenDisplay.SetSymbol(DirectCast(_theLineSymbol, ISymbol))
                    If (_theLinePolyline.Length > 0) Then _
                      pActiveView.ScreenDisplay.DrawPolyline(_theLinePolyline)
                End If

                'Get line between from and to points, and angle for text
                Dim pLine As ILine = New ESRI.ArcGIS.Geometry.Line
                pLine.PutCoords(_theStartPoint, pPoint)
                Dim angle As Double = pLine.Angle
                angle = angle * (180.0# / 3.14159)
                If ((angle > 90.0#) And (angle < 180.0#)) Then
                    angle = angle + 180.0#
                ElseIf ((angle < 0.0#) And (angle < -90.0#)) Then
                    angle = angle - 180.0#
                ElseIf ((angle < -90.0#) And (angle > -180)) Then
                    angle = angle - 180.0#
                ElseIf (angle > 180) Then
                    angle = angle - 180.0#
                End If

                'For drawing text, get text(distance), angle, and point
                Dim deltaX As Double = pPoint.X - _theStartPoint.X
                Dim deltaY As Double = pPoint.Y - _theStartPoint.Y
                _theTextPoint.X = _theStartPoint.X + deltaX / 2.0#
                _theTextPoint.Y = _theStartPoint.Y + deltaY / 2.0#
                _theTextSymbol.Angle = angle
                _theTextSymbol.Text = ""

                'Draw text
                pActiveView.ScreenDisplay.SetSymbol(DirectCast(_theTextSymbol, ISymbol))
                pActiveView.ScreenDisplay.DrawText(_theTextPoint, _theTextSymbol.Text)

                'Get polyline with blank space for text
                Dim pPolyline As IPolyline = New ESRI.ArcGIS.Geometry.Polyline
                Dim pSegColl As ISegmentCollection = DirectCast(pPolyline, ISegmentCollection)
                pSegColl.AddSegment(DirectCast(pLine, ISegment))
                _theLinePolyline = GetSmashedLine(pActiveView.ScreenDisplay, DirectCast(_theTextSymbol, ISymbol), _theTextPoint, pPolyline)

                'Draw polyline
                pActiveView.ScreenDisplay.SetSymbol(DirectCast(_theLineSymbol, ISymbol))
                If (_theLinePolyline.Length > 0) Then _
                  pActiveView.ScreenDisplay.DrawPolyline(_theLinePolyline)
                pActiveView.ScreenDisplay.FinishDrawing()
            End If

            '++ START Added by Laura Gordon, 02/20/2007
            If arrowType.Equals("Dimension", _ignoreCase) Then

                m_pEditor = EditorExtension.Editor
                m_pEditorEvents = DirectCast(m_pEditor, Editor)

                If _theToolJustCompletedTask Then 'gets rid of a stray editor agent
                    _theActiveView.Refresh()
                    _theToolJustCompletedTask = False
                    Exit Sub
                End If

                If _theMouseHasMoved Then
                    'Check to be sure m_pPt has a value, prevents an error if called after other tools
                    If _thePt Is Nothing Then
                        _theMouseHasMoved = False
                        Exit Sub
                    End If
                    'erase the old agent
                    m_pEditor.InvertAgent(_thePt, 0)
                    'get the new point
                    GetMousePoint(X, Y)
                Else
                    _thePt = _theActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                End If
                _theMouseHasMoved = True

                If Not _theSnapAgent Is Nothing Then
                    Dim pSnapenv As ISnapEnvironment
                    pSnapenv = DirectCast(m_pEditor, ISnapEnvironment)
                    Dim pTmpPt As IPoint = New ESRI.ArcGIS.Geometry.Point
                    pTmpPt.PutCoords(_thePt.X, _thePt.Y)
                    Dim dTol As Double
                    dTol = pSnapenv.SnapTolerance
                    _theInTol = pSnapenv.SnapPoint(_thePt)
                    If Not _theInTol Then
                        _thePt = New ESRI.ArcGIS.Geometry.Point
                        _thePt.PutCoords(pTmpPt.X, pTmpPt.Y)  '+++ set the point back b/c it was not in the snap tol
                    End If
                Else
                    _theInTol = False
                End If

                m_pEditor.InvertAgent(_thePt, 0) 'draw the agent

            End If


        Catch ex As Exception

        End Try


    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        MyBase.OnMouseUp(Button, Shift, X, Y)

        If arrowType.Equals("Hook", _ignoreCase) Then
            If (Not _theInUse) Then Exit Sub 'SC ??

            If (_theLineSymbol Is Nothing) Then Exit Sub 'SC ??

            Dim theArcMapDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
            Dim pActiveView As IActiveView = theArcMapDoc.ActiveView

            ' Draws a temporary line on the screen
            pActiveView.ScreenDisplay.StartDrawing(pActiveView.ScreenDisplay.hDC, -1)
            pActiveView.ScreenDisplay.SetSymbol(DirectCast(_theTextSymbol, ISymbol))
            pActiveView.ScreenDisplay.DrawText(_theTextPoint, _theTextSymbol.Text)
            pActiveView.ScreenDisplay.SetSymbol(DirectCast(_theLineSymbol, ISymbol))
            If (_theLinePolyline.Length > 0) Then pActiveView.ScreenDisplay.DrawPolyline(_theLinePolyline)
            pActiveView.ScreenDisplay.FinishDrawing()

            ' Generate hooks based on the graphic polyline
            GenerateHooks(DirectCast(_theLinePolyline, IPolyline))


            ' Records that the tool is no longer in use
            _theInUse = False

            ' Clean up
            _theTextSymbol = Nothing
            _theTextPoint = Nothing
            _theLinePolyline = Nothing
            _theLineSymbol = Nothing

        End If

    End Sub

    Public Overrides Sub OnKeyDown(ByVal keyCode As Integer, ByVal Shift As Integer)
        MyBase.OnKeyDown(keyCode, Shift)

        '++ START Added by Laura Gordon 02/20/2007
        'End the dimension arrow tool if the "q" key is pressed
        If keyCode = System.Windows.Forms.Keys.Q Then
            If arrowType.Equals("Dimension", _ignoreCase) Then
                'Deactivate the tool
                _theArrowPt1 = Nothing
                _theArrowPt2 = Nothing
                _theArrowPt3 = Nothing
                _theArrowPt4 = Nothing
                _thePt = Nothing
                _theArrowPtTemp = Nothing
                _theArrowPtTemp2 = Nothing
                _application.CurrentTool = Nothing
                _application.RefreshWindow()
                _theToolJustCompletedTask = True
                If arrowType.Equals("Dimension", _ignoreCase) Then

                End If
            ElseIf arrowType.Equals("Arrow", _ignoreCase) Then
                'Deactivate the tool
                _theArrowPt1 = Nothing
                _theArrowPt2 = Nothing
                _theArrowPt3 = Nothing
                _theArrowPt4 = Nothing
                _thelArrowPointsCollection = Nothing
                _application.CurrentTool = Nothing
                _application.RefreshWindow()
                _theToolJustCompletedTask = True
            Else 'Hooks
                'Deactivate the tool
                _application.CurrentTool = Nothing
                _application.RefreshWindow()
                _theToolJustCompletedTask = True
            End If
        End If
        If keyCode = System.Windows.Forms.Keys.D Then
            If arrowType.Equals("Dimension", _ignoreCase) Then
                _theolDimensionChanged = True
                MsgBox("show dialog of another form - sc fix")
                PartnerDimensionArrowsForm.ShowDialog()
            End If
        End If

    End Sub

#End Region



#End Region

#Region "Implemented Interface Members"

#Region "IDisposable Interface Implementation"

    Private _isDuringDispose As Boolean ' Used to track whether Dispose() has been called and is in progress.

    ''' <summary>
    ''' Dispose of managed and unmanaged resources.
    ''' </summary>
    ''' <param name="disposing">True or False.</param>
    ''' <remarks>
    ''' <para>Member of System::IDisposable.</para>
    ''' <para>Dispose executes in two distinct scenarios. 
    ''' If disposing equals true, the method has been called directly
    ''' or indirectly by a user's code. Managed and unmanaged resources
    ''' can be disposed.</para>
    ''' <para>If disposing equals false, the method has been called by the 
    ''' runtime from inside the finalizer and you should not reference 
    ''' other objects. Only unmanaged resources can be disposed.</para>
    ''' </remarks>
    Friend Sub Dispose(ByVal disposing As Boolean)
        Try
            ' Check to see if Dispose has already been called.
            If Not Me._isDuringDispose Then

                ' Flag that disposing is in progress.
                Me._isDuringDispose = True

                If disposing Then
                    ' Free managed resources when explicitly called.

                    ' Dispose managed resources here.
                    '   e.g. component.Dispose()

                End If

                ' Free "native" (shared unmanaged) resources, whether 
                ' explicitly called or called by the runtime.

                ' Call the appropriate methods to clean up 
                ' unmanaged resources here.
                _bitmapResourceName = Nothing
                MyBase.m_bitmap = Nothing

                ' Flag that disposing has been finished.
                _isDuringDispose = False

            End If
        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try
    End Sub

#Region " IDisposable Support "

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

#End Region

#End Region

#Region "Other Members"

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "952aa746-4886-42e9-bd85-0b3b08fa1a95"
    Public Const InterfaceId As String = "b8b0c98f-df9e-4078-8736-993c7a1e0d1d"
    Public Const EventsId As String = "81c6ba05-d21b-480e-aea7-9990483d3cc1"
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
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub

    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

#End Region

End Class



