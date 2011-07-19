Public Class TransposeAnnotationButton
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        Dim _transposeAnnotation As TransposeAnnotation = New TransposeAnnotation
        _transposeAnnotation.DoButtonOperation()
  End Sub

  Protected Overrides Sub OnUpdate()

  End Sub
End Class
