Public Class InvertAnnotationButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        Dim _invertAnnotation As InvertAnnotation = New InvertAnnotation
        _invertAnnotation.DoButtonOperation()
  End Sub

  Protected Overrides Sub OnUpdate()

  End Sub
End Class
