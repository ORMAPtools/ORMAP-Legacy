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
                'setPartnerSCSDockFormUI()
            End If
            Return _partnerSCSDockWindowUI
        End Get
    End Property

    Private Sub setPartnerSCSDockFormUI(ByVal value As SpiralCurveSpiralDockWindow)
        If value IsNot Nothing Then
            _partnerSCSDockWindowUI = value
            'subscribe to partner form events
            AddHandler _partnerSCSDockWindowUI.uxCreate.Click, AddressOf uxCreate_Click
            AddHandler _partnerSCSDockWindowUI.uxCancel.Click, AddressOf uxCancel_Click
            AddHandler _partnerSCSDockWindowUI.uxHelp.Click, AddressOf uxHelp_Click
            AddHandler _partnerSCSDockWindowUI.uxGettoPoint.Click, AddressOf uxGettoPoint_Click
            AddHandler _partnerSCSDockWindowUI.uxGetTangentPoint.Click, AddressOf uxGetTangentPoint_Click
        Else
            'unSubscribe to partner form events
            RemoveHandler _partnerSCSDockWindowUI.uxCreate.Click, AddressOf uxCreate_Click
            RemoveHandler _partnerSCSDockWindowUI.uxCancel.Click, AddressOf uxCancel_Click
            RemoveHandler _partnerSCSDockWindowUI.uxHelp.Click, AddressOf uxHelp_Click
            RemoveHandler _partnerSCSDockWindowUI.uxGettoPoint.Click, AddressOf uxGettoPoint_Click
            RemoveHandler _partnerSCSDockWindowUI.uxGetTangentPoint.Click, AddressOf uxGetTangentPoint_Click
        End If
    End Sub
#End Region

#Region "Event Handler"

    Private Sub uxCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub uxCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxGettoPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub uxGetTangentPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

#End Region

#Region "Methods"
    Friend Sub DoButtonOperation()

    End Sub

    Protected Overrides Sub OnClick()
        Try

        Catch ex As Exception

        End Try
    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub

#End Region

End Class
