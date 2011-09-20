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


Public Class SpiralConstruction
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
#Region "Constructors"
    Public Sub New()
        Try

            Dim windowID As UID = New UIDClass
            windowID.Value = "ORMAP_SpiralTools_SpiralCurveSpiralDockWindow"
            _partnerSpiralDockWindow = My.ArcMap.DockableWindowManager.GetDockableWindow(windowID)

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try
    End Sub
#End Region

#Region "Properties"

    Private _partnerSpiralDockWindow As IDockableWindow
    Private WithEvents _partnerSpiralDockWindowUI As SpiralDockWindow

    Friend ReadOnly Property partnerSpiralDockWindowUI() As SpiralDockWindow
        Get
            If _partnerSpiralDockWindowUI Is Nothing Then
                setPartnerSpiralDockWindowUI(AddIn.FromID(Of SpiralDockWindow.AddinImpl)(My.ThisAddIn.IDs.SpiralDockWindow).UI)
            End If
            Return _partnerSpiralDockWindowUI
        End Get
    End Property
    Private Sub setPartnerSpiralDockWindowUI(ByVal value As SpiralDockWindow)

    End Sub
#End Region

#Region "Methods"

    Protected Overrides Sub OnClick()
        Try

        Catch ex As Exception

        End Try
    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub
#End Region
End Class
