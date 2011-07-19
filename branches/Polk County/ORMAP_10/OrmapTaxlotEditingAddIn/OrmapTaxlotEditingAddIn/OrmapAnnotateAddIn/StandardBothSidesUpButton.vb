Public Class StandardBothSidesUpButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        Dim _standardBothSidesUp As StandardBothSidesUp = New StandardBothSidesUp
        _standardBothSidesUp.DoButtonOperation()
  End Sub

  Protected Overrides Sub OnUpdate()

  End Sub
End Class
