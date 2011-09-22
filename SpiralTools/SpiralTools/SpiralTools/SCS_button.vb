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

#End Region

Public Class SCS_button
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
#Region "constructors"
    Public Sub New()
        Try

            Dim windowID As UID = New UIDClass
            windowID.Value = "ORMAP_SpiralTools_SpiralCurveSpiralDockWindow"
            _partnerSCSDockWindow = My.ArcMap.DockableWindowManager.GetDockableWindow(windowID)

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub
#End Region
#Region "Properties"

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
            AddHandler _partnerSCSDockWindowUI.uxSpiralsbyArclength.CheckedChanged, AddressOf uxSpiralsbyArclength_CheckedChanged
            AddHandler _partnerSCSDockWindowUI.uxSpiralsbyDelta.CheckedChanged, AddressOf uxSpiralsbyDelta_CheckedChanged
            AddHandler _partnerSCSDockWindowUI.uxCurvetotheRight.CheckedChanged, AddressOf uxCurvetotheRight_CheckedChanged
            AddHandler _partnerSCSDockWindowUI.uxCurvetotheLeft.CheckedChanged, AddressOf uxCurvetotheLeft_CheckedChanged
        Else
            'unSubscribe to partner form events
            RemoveHandler _partnerSCSDockWindowUI.uxCreate.Click, AddressOf uxCreate_Click
            RemoveHandler _partnerSCSDockWindowUI.uxHelp.Click, AddressOf uxHelp_Click
            RemoveHandler _partnerSCSDockWindowUI.uxGettoPoint.Click, AddressOf uxGettoPoint_Click
            RemoveHandler _partnerSCSDockWindowUI.uxGetTangentPoint.Click, AddressOf uxGetTangentPoint_Click
            RemoveHandler _partnerSCSDockWindowUI.uxGetFromPoint.Click, AddressOf uxGetFromPoint_Click
            RemoveHandler _partnerSCSDockWindowUI.uxCurveByRadius.CheckedChanged, AddressOf uxCurveByRadius_CheckedChanged
            RemoveHandler _partnerSCSDockWindowUI.uxCurvebyDegree.CheckedChanged, AddressOf uxCurvebyDegree_CheckedChanged
            RemoveHandler _partnerSCSDockWindowUI.uxSpiralsbyArclength.CheckedChanged, AddressOf uxSpiralsbyArclength_CheckedChanged
            RemoveHandler _partnerSCSDockWindowUI.uxSpiralsbyDelta.CheckedChanged, AddressOf uxSpiralsbyDelta_CheckedChanged
            RemoveHandler _partnerSCSDockWindowUI.uxCurveByRadius.CheckedChanged, AddressOf uxCurvetotheRight_CheckedChanged
            RemoveHandler _partnerSCSDockWindowUI.uxCurvetotheLeft.CheckedChanged, AddressOf uxCurvetotheLeft_CheckedChanged
        End If
    End Sub
#End Region

#Region "Event Handler"

    Private Sub uxCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("This Works")
    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub uxGettoPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim theFromPoint As IPoint = GetPointFromMouseClick(My.ArcMap.Document.ActiveView)

        _partnerSCSDockWindowUI.uxToPointXValue.Text = theFromPoint.X.ToString
        _partnerSCSDockWindowUI.uxToPointYValue.Text = theFromPoint.Y.ToString
    End Sub
    Private Sub uxGetTangentPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxGetFromPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxCurveByRadius_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxCurvebyDegree_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxSpiralsbyArclength_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxSpiralsbyDelta_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxCurvetotheRight_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxCurvetotheLeft_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
#End Region

#Region "Methods"
    Public Function GetPointFromMouseClick(ByVal theActiveView As IActiveView) As IPoint

        'Dim screenDisplay As ESRI.ArcGIS.Display.IScreenDisplay = activeView.ScreenDisplay

        'Dim rubberBand As ESRI.ArcGIS.Display.IRubberBand = New ESRI.ArcGIS.Display.RubberLineClass
        'Dim geometry As ESRI.ArcGIS.Geometry.IGeometry = rubberBand.TrackNew(screenDisplay, Nothing)

        'Dim polyline As ESRI.ArcGIS.Geometry.IPolyline = CType(geometry, ESRI.ArcGIS.Geometry.IPolyline)
        Dim theScreenDisplay As IScreenDisplay2 = CType(theActiveView.ScreenDisplay, IScreenDisplay2)

        Dim theRubberBand As IRubberBand2 = New ESRI.ArcGIS.Display.RubberPointClass
        Dim thePointGeometry As IGeometry5 = CType(theRubberBand.TrackNew(CType(theScreenDisplay, IScreenDisplay), Nothing), IGeometry5)


        Dim ThePoint As IPoint = CType(thePointGeometry, IPoint)

        Return ThePoint
    End Function
    Friend Sub DoButtonOperation()

        _partnerSCSDockWindow.Show(Not _partnerSCSDockWindow.IsVisible)

    End Sub

    Protected Overrides Sub OnClick()
        Try
            DoButtonOperation()
        Catch ex As Exception

        End Try
    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub

#End Region

End Class
