''' <summary>
''' ESRI AddIn button events
''' </summary>
Public Class MoveDownButton
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()

    End Sub

    ''' <summary>
    ''' ESRI AddIn button OnClick event handler
    ''' </summary>
    Protected Overrides Sub OnClick()
        'Dim _moveDown As MoveDown = New MoveDown
        MoveDown.DoButtonOperation()
    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub
End Class
