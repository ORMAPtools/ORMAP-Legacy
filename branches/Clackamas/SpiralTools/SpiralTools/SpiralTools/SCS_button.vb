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

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub
#End Region
#Region "Properties"

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
        MessageBox.Show("This Works")
    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxGettoPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        _IsGettingToPoint = True
        MyBase.Cursor = Cursors.Cross
    End Sub
    Private Sub uxGetTangentPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim thePolyline As IPolyline5 = TestRubberBand(My.ArcMap.Document.ActiveView)

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
    Public Function TestRubberBand(ByVal theActiveView As IActiveView) As IPolyline5
        Dim screenDisplay As ESRI.ArcGIS.Display.IScreenDisplay = theActiveView.ScreenDisplay

        Dim rubberBand As ESRI.ArcGIS.Display.IRubberBand = New ESRI.ArcGIS.Display.RubberLineClass
        Dim geometry As ESRI.ArcGIS.Geometry.IGeometry = rubberBand.TrackNew(screenDisplay, Nothing)

        Dim polyline As ESRI.ArcGIS.Geometry.IPolyline = CType(geometry, ESRI.ArcGIS.Geometry.IPolyline)
        Return CType(polyline, IPolyline5)

    End Function
    'Public Function GetPointFromMouseClick(ByVal theActiveView As IActiveView) As IPoint

    '    Dim theScreenDisplay As IScreenDisplay2 = CType(theActiveView.ScreenDisplay, IScreenDisplay2)
    '    Dim theRubberBand As IRubberBand2 = New ESRI.ArcGIS.Display.RubberPoint
    '    Dim thePointGeometry As IGeometry5 = CType(theRubberBand.TrackNew(CType(theScreenDisplay, IScreenDisplay), Nothing), IGeometry5)

    '    Dim ThePoint As IPoint = CType(thePointGeometry, IPoint)

    '    Return ThePoint
    'End Function

    Friend Sub DoButtonOperation()

        With partnerSCSDockFormUI
            .uxTargetTemplate.Text = "Construction Lines"
        End With
        _partnerSCSDockWindow.Show(Not _partnerSCSDockWindow.IsVisible)

    End Sub

    'Protected Overrides Sub OnClick()
    '    Try
    '        DoButtonOperation()
    '    Catch ex As Exception

    '    End Try
    'End Sub
    Protected Overrides Sub OnMouseDown(ByVal arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
        MyBase.OnMouseDown(arg)
        If arg.Button = MouseButtons.Left And arg.Shift = True Then
            DoButtonOperation()
        ElseIf arg.Button = MouseButtons.Left And _IsGettingToPoint Then
            Dim TheToPoint As IPoint = getSnapPoint(getDataFrameCoords(arg.X, arg.Y))
            _partnerSCSDockWindowUI.uxToPointXValue.Text = TheToPoint.X.ToString
            _partnerSCSDockWindowUI.uxToPointYValue.Text = TheToPoint.Y.ToString
            _IsGettingToPoint = False
            MyBase.Cursor = Cursors.Arrow
        End If
    End Sub
    'Protected Overrides Sub OnMouseDown(ByVal arg As ESRI.ArcGIS.Desktop.AddIns.Tool.MouseEventArgs)
    '    MyBase.OnMouseDown(arg)
    '    MsgBox(arg.X & " " & arg.Y)
    'End Sub
    Protected Overrides Sub OnUpdate()
        Me.Enabled = SpiralUtilities.IsEnable
    End Sub

#End Region

    

End Class


