Public Class StandardBothSidesDownButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        Dim _standardBothSidesDown As StandardBothSidesDown = New StandardBothSidesDown
        _standardBothSidesDown.DoButtonOperation()
  End Sub

  Protected Overrides Sub OnUpdate()

  End Sub
End Class
