''' <summary>
''' ESRI AddIn button events
''' </summary>
Public Class WideBothSidesDownButton
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()

    End Sub

    ''' <summary>
    ''' ESRI AddIn button OnClick event handler
    ''' </summary>
    Protected Overrides Sub OnClick()
        'Dim _wideBothSidesDown As WideBothSidesDown = New WideBothSidesDown
        WideBothSidesDown.DoButtonOperation()
    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub
End Class
