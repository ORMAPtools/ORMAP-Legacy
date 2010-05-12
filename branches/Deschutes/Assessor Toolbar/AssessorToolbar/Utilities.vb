
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Display


Public NotInheritable Class Utilities

    Public Shared Function GetMapNumberArray(ByVal theMxDoc As IMxDocument) As Array

        Dim theMapIndexFeatureLayer As IFeatureLayer = GetFeatureLayerByName("SeeMaps", theMxDoc)
        Dim mapNumberStringList As List(Of String) = Nothing

        If Not theMapIndexFeatureLayer Is Nothing Then
            mapNumberStringList = New List(Of String)
            Dim theMapIndexFClass As IFeatureClass = theMapIndexFeatureLayer.FeatureClass
            Dim theQueryFilter As IQueryFilter = New QueryFilter
            theQueryFilter.SubFields = "MapNumber"
            Dim theFeatCursor As IFeatureCursor = theMapIndexFClass.Search(theQueryFilter, True)
            Dim theFeature As IFeature = theFeatCursor.NextFeature
            Dim theFieldIdx As Integer = theFeature.Fields.FindField("MapNumber")
            Do Until theFeature Is Nothing
                Dim theMapNumberVal As String = theFeature.Value(theFieldIdx).ToString
                If Not mapNumberStringList.Contains(theMapNumberVal) Then
                    mapNumberStringList.Add(theMapNumberVal)
                End If
                theFeature = theFeatCursor.NextFeature
            Loop

        End If

        Return mapNumberStringList.ToArray

    End Function


    Public Shared Function GetFeatureLayerByName(ByVal featureClassName As String, ByVal theMxDoc As IMxDocument) As IFeatureLayer

        '-- Get reference to map object
        Dim theMap As IMap = theMxDoc.FocusMap

        Dim theFeatureLayer As IFeatureLayer = Nothing
        Dim theLayersEnum As IEnumLayer = theMap.Layers
        Dim theLayer As ILayer = theLayersEnum.Next
        Do Until theLayer Is Nothing
            If TypeOf theLayer Is IFeatureLayer AndAlso theLayer.Name.ToUpper = featureClassName.ToUpper Then
                theFeatureLayer = theLayer
                Exit Do
            End If
            theLayer = theLayersEnum.Next
        Loop

        Return theFeatureLayer

    End Function


    Public Shared Function GetRGBColor(ByVal theColor As System.Drawing.Color) As IRgbColor

        Dim theRGBColor As IRgbColor = New RgbColor
        With theRGBColor
            .Red = theColor.R
            .Green = theColor.G
            .Blue = theColor.B
            .Transparency = 1
        End With

        Return theRGBColor

    End Function


    Public Shared Function MakeInputDialog(ByVal Prompt As String, ByVal Title As String, ByVal Message As String) As InputDialog

        Dim theDialogForm As New InputDialog
        theDialogForm.uxLabel.Text = Prompt
        theDialogForm.Text = Title
        theDialogForm.Message = Message

        Return theDialogForm

    End Function

    Public Shared Function MakeSelectMapIndexDialog(ByVal theMxDoc As IMxDocument) As SelectMapindexDialog

        Dim theSelectmapIndexDialog As New SelectMapindexDialog
        theSelectmapIndexDialog.uxMapNumber.AutoCompleteCustomSource.Clear()
        theSelectmapIndexDialog.uxMapNumber.AutoCompleteCustomSource.AddRange(GetMapNumberArray(theMxDoc))
        theSelectmapIndexDialog.uxMapNumber.Text = String.Empty

        Return theSelectmapIndexDialog

    End Function


    Private Shared _referenceScale As Integer = Nothing
    Friend Shared Property ReferenceScale() As Integer
        Get
            Return _referenceScale
        End Get
        Set(ByVal value As Integer)
            _referenceScale = value
        End Set
    End Property

End Class

