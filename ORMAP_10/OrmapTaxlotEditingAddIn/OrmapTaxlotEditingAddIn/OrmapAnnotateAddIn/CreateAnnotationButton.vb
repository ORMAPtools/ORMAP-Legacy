Public Class CreateAnnotationButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        Dim _createAnnotation As CreateAnnotation = New CreateAnnotation
        _createAnnotation.DoButtonOperation()
  End Sub

  Protected Overrides Sub OnUpdate()

  End Sub
End Class
