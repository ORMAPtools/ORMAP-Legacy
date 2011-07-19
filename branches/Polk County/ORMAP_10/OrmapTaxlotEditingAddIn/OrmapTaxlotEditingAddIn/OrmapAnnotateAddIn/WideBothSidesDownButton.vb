Public Class WideBothSidesDownButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        Dim _wideBothSidesDown As WideBothSidesDown = New WideBothSidesDown
        _wideBothSidesDown.DoButtonOperation()
  End Sub

  Protected Overrides Sub OnUpdate()

  End Sub
End Class
