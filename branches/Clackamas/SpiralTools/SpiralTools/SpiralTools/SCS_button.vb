#Region "Imported Namespaces"
Imports System.Drawing
Imports System.Environment
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Desktop.AddIns
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Editor
Imports SpiralTools.SpiralUtilities



#End Region

Public Class SCS_button
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool

#Region "constructors"
    Public Sub New()
        Try

            Dim windowID As UID = New UIDClass
            windowID.Value = "ORMAP_SpiralTools_SpiralCurveSpiralDockWindow"
            _partnerSCSDockWindow = My.ArcMap.DockableWindowManager.GetDockableWindow(windowID)

            If _partnerSCSDockWindow.IsVisible Then
                _partnerSCSDockWindow.Show(False)
            End If
        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub
#End Region
#Region "Properties"

    Private _IsGettingTangentPoint As Boolean = False
    Private _IsGettingFromPoint As Boolean = False
    Private _IsGettingToPoint As Boolean = False
    Private _partnerSCSDockWindow As IDockableWindow

    Private WithEvents _partnerSCSDockWindowUI As SpiralCurveSpiralDockWindow


   
    Friend ReadOnly Property partnerSCSDockFormUI() As SpiralCurveSpiralDockWindow
        Get
            If _partnerSCSDockWindowUI Is Nothing Then
                setPartnerSCSDockFormUI(AddIn.FromID(Of SpiralCurveSpiralDockWindow.AddinImpl)(My.ThisAddIn.IDs.SpiralCurveSpiralDockWindow).UI)
            End If
            Return _partnerSCSDockWindowUI
        End Get
    End Property

    Private Sub setPartnerSCSDockFormUI(ByVal value As SpiralCurveSpiralDockWindow)
        If value IsNot Nothing Then
            _partnerSCSDockWindowUI = value
            'subscribe to partner form events
            AddHandler _partnerSCSDockWindowUI.uxCreate.Click, AddressOf uxCreate_Click
            AddHandler _partnerSCSDockWindowUI.uxHelp.Click, AddressOf uxHelp_Click
            AddHandler _partnerSCSDockWindowUI.uxGettoPoint.Click, AddressOf uxGettoPoint_Click
            AddHandler _partnerSCSDockWindowUI.uxGetTangentPoint.Click, AddressOf uxGetTangentPoint_Click
            AddHandler _partnerSCSDockWindowUI.uxGetFromPoint.Click, AddressOf uxGetFromPoint_Click
            AddHandler _partnerSCSDockWindowUI.uxCurveByRadius.CheckedChanged, AddressOf uxCurveByRadius_CheckedChanged
            AddHandler _partnerSCSDockWindowUI.uxCurvebyDegree.CheckedChanged, AddressOf uxCurvebyDegree_CheckedChanged
           
        Else
            'unSubscribe to partner form events
            RemoveHandler _partnerSCSDockWindowUI.uxCreate.Click, AddressOf uxCreate_Click
            RemoveHandler _partnerSCSDockWindowUI.uxHelp.Click, AddressOf uxHelp_Click
            RemoveHandler _partnerSCSDockWindowUI.uxGettoPoint.Click, AddressOf uxGettoPoint_Click
            RemoveHandler _partnerSCSDockWindowUI.uxGetTangentPoint.Click, AddressOf uxGetTangentPoint_Click
            RemoveHandler _partnerSCSDockWindowUI.uxGetFromPoint.Click, AddressOf uxGetFromPoint_Click
            RemoveHandler _partnerSCSDockWindowUI.uxCurveByRadius.CheckedChanged, AddressOf uxCurveByRadius_CheckedChanged
            RemoveHandler _partnerSCSDockWindowUI.uxCurvebyDegree.CheckedChanged, AddressOf uxCurvebyDegree_CheckedChanged
            
        End If
    End Sub
#End Region

#Region "Event Handler"
    
    Private Sub uxCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim theFromPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
        Dim theTangent As IPoint = New ESRI.ArcGIS.Geometry.Point
        Dim theToPoint As IPoint = New ESRI.ArcGIS.Geometry.Point

        With _partnerSCSDockWindowUI
            If .uxCurveByRadius.Checked And .uxCurveByRadiusValue.TextLength = 0 Then
                MessageBox.Show("Please enter a radius value")
                Exit Sub
            ElseIf .uxCurveByRadius.Checked And Not IsNumeric(.uxCurveByRadiusValue.Text) Then
                MessageBox.Show("Please us a numeric value in the radius value box")
                Exit Sub
            End If
            If .uxCurvebyDegree.Checked And .uxCurveDegreeValue.TextLength = 0 Then
                MessageBox.Show("Please eneter a degree value")
                Exit Sub
            ElseIf .uxCurvebyDegree.Checked And Not IsNumeric(.uxCurveDegreeValue.Text) Then
                MessageBox.Show("Please us a numeric value in the degree value box")
            End If

            theFromPoint.PutCoords(CDbl(.uxFromPointXValue.Text), CDbl(.uxFromPointYValue.Text))
            theTangent.PutCoords(CDbl(.uxTangentPointXValue.Text), CDbl(.uxTangentPointYValue.Text))
            theToPoint.PutCoords(CDbl(.uxToPointXValue.Text), CDbl(.uxToPointYValue.Text))

            If .uxCurveByRadius.Checked Then
                ConstructSCSbyLength(theFromPoint, theTangent, theToPoint, CDbl(.uxArcLengthValue.Text), CDbl(.uxCurveByRadiusValue.Text), .uxCurvetotheRight.Checked)
            ElseIf .uxCurvebyDegree.Checked Then
                Dim theCurveRadius As Double = 5729.578 / CDbl(.uxCurveDegreeValue.Text)
                ConstructSCSbyLength(theFromPoint, theTangent, theToPoint, CDbl(.uxArcLengthValue.Text), theCurveRadius, .uxCurvetotheRight.Checked)
            End If
        End With

        My.ArcMap.Document.ActiveView.Refresh()

    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxGettoPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _IsGettingToPoint = True
        _IsCircleActive = True
        MyBase.Cursor = Cursors.Cross

    End Sub
    Private Sub uxGetTangentPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _IsGettingTangentPoint = True
        _IsCircleActive = True
        MyBase.Cursor = Cursors.Cross

    End Sub
    Private Sub uxGetFromPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _IsGettingFromPoint = True
        _IsCircleActive = True
        MyBase.Cursor = Cursors.Cross

    End Sub
    Private Sub uxCurveByRadius_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If _partnerSCSDockWindowUI.uxCurveByRadius.Checked Then
            _partnerSCSDockWindowUI.uxCurveByRadiusValue.Enabled = True
        Else
            _partnerSCSDockWindowUI.uxCurveByRadiusValue.Enabled = False
        End If

    End Sub
    Private Sub uxCurvebyDegree_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If _partnerSCSDockWindowUI.uxCurvebyDegree.Checked Then
            _partnerSCSDockWindowUI.uxCurveDegreeValue.Enabled = True
        Else
            _partnerSCSDockWindowUI.uxCurveDegreeValue.Enabled = False
        End If

    End Sub
   
    Private Sub partnerSCSDockWindow_load()
        Try
            Dim theEnumLayer As IEnumLayer = My.ArcMap.Editor.Map.Layers
            theEnumLayer.Reset()
            Dim thisLayer As ILayer = CType(theEnumLayer.Next, ILayer)
            Do While Not (thisLayer Is Nothing)
                If TypeOf thisLayer Is FeatureLayer Then
                    Dim thisFeatureLayer As IFeatureLayer = CType(thisLayer, IFeatureLayer)
                    If thisFeatureLayer.FeatureClass.ShapeType = 3 Then
                        _partnerSCSDockWindowUI.uxTargetTemplate.AutoCompleteCustomSource.Add(thisFeatureLayer.Name)
                    End If
                End If
                thisLayer = CType(theEnumLayer.Next, ILayer)
            Loop
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
#End Region

#Region "Methods"


    Friend Sub DoButtonOperation()

        With partnerSCSDockFormUI
            .uxCreate.Enabled = True
        End With
        If _partnerSCSDockWindow.IsVisible AndAlso _partnerSCSDockWindowUI.uxTargetTemplate.AutoCompleteCustomSource.Count = 0 Then partnerSCSDockWindow_load()

        _partnerSCSDockWindow.Show(Not _partnerSCSDockWindow.IsVisible)

        If _partnerSCSDockWindow.IsVisible AndAlso _partnerSCSDockWindowUI.uxTargetTemplate.AutoCompleteCustomSource.Count = 0 Then partnerSCSDockWindow_load()

    End Sub

    Protected Overrides Sub OnMouseDown(ByVal arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseDown(arg)
        If arg.Button = MouseButtons.Left And arg.Shift = True Then
            DoButtonOperation()
        ElseIf arg.Button = MouseButtons.Left And _IsGettingToPoint Then
            Dim theToPoint As IPoint = getSnapPoint(getDataFrameCoords(arg.X, arg.Y))
            _partnerSCSDockWindowUI.uxToPointXValue.Text = theToPoint.X.ToString
            _partnerSCSDockWindowUI.uxToPointYValue.Text = theToPoint.Y.ToString
            _IsGettingToPoint = False
            MyBase.Cursor = Cursors.Arrow
        ElseIf arg.Button = MouseButtons.Left And _IsGettingFromPoint Then
            Dim theFromPoint As IPoint = getSnapPoint(getDataFrameCoords(arg.X, arg.Y))
            _partnerSCSDockWindowUI.uxFromPointXValue.Text = theFromPoint.X.ToString
            _partnerSCSDockWindowUI.uxFromPointYValue.Text = theFromPoint.Y.ToString
            _IsGettingFromPoint = False
            _IsCircleActive = False
            MyBase.Cursor = Cursors.Arrow
        ElseIf arg.Button = MouseButtons.Left And _IsGettingTangentPoint Then
            Dim theTangentPoint As IPoint = getSnapPoint(getDataFrameCoords(arg.X, arg.Y))
            _partnerSCSDockWindowUI.uxTangentPointXValue.Text = theTangentPoint.X.ToString
            _partnerSCSDockWindowUI.uxTangentPointYValue.Text = theTangentPoint.Y.ToString
            _IsGettingTangentPoint = False
            MyBase.Cursor = Cursors.Arrow
        End If
        _IsCircleActive = False

    End Sub
    'Show Snap Point
    Private _SnappingMarkerElement As IElement = Nothing
    Private _IsCircleActive As Boolean = False
    Private _TheSnapPoint As IPoint

    'Move below sub routine to Spiral Utilities
    Protected Overrides Sub OnMouseMove(ByVal arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseMove(arg)
        Try
            If _IsCircleActive And _SnappingMarkerElement Is Nothing Then
                _SnappingMarkerElement = CType(Create_Snap_Marker(), IElement)
                Dim theGraphicsContainer As IGraphicsContainer = CType(My.ArcMap.Document.ActiveView, IGraphicsContainer)
                theGraphicsContainer.AddElement(_SnappingMarkerElement, 0)
                Dim theCurrentPoint As IPoint = getSnapPoint(getDataFrameCoords(arg.X, arg.Y))
                _SnappingMarkerElement.Geometry = theCurrentPoint
                My.ArcMap.Document.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
            ElseIf _IsCircleActive And Not _SnappingMarkerElement Is Nothing Then
                _SnappingMarkerElement.Geometry = CType(getSnapPoint(getDataFrameCoords(arg.X, arg.Y)), IPoint)
                My.ArcMap.Document.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
            ElseIf Not _IsCircleActive And Not _SnappingMarkerElement Is Nothing Then
                Dim theGraphicsContainer As IGraphicsContainer = CType(My.ArcMap.Document.ActiveView, IGraphicsContainer)
                theGraphicsContainer.DeleteElement(_SnappingMarkerElement)
                _SnappingMarkerElement.Geometry.SetEmpty()
                _SnappingMarkerElement = Nothing
                My.ArcMap.Document.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewAll, Nothing, Nothing)
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Protected Overrides Sub OnUpdate()
        Me.Enabled = SpiralUtilities.IsEnable
        If Not Me.Enabled Then
            _partnerSCSDockWindow.Show(False)
        End If

    End Sub

#End Region

    Private Function featureclass() As Object
        Throw New NotImplementedException
    End Function



End Class


