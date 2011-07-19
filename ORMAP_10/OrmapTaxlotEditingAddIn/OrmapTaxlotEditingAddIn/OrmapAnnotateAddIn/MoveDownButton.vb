Public Class MoveDownButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        Dim _moveDown As MoveDown = New MoveDown
        _moveDown.DoButtonOperation()
  End Sub

  Protected Overrides Sub OnUpdate()

  End Sub
End Class
