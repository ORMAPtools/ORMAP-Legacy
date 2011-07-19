Public Class WideBothSidesUpButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        Dim _wideBothSidesUp As WideBothSidesUp = New WideBothSidesUp
        _wideBothSidesUp.DoButtonOperation()
  End Sub

    Protected Overrides Sub OnUpdate()

    End Sub
End Class
