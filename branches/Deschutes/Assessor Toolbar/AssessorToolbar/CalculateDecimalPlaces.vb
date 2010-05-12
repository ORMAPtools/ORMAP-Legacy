Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Editor

<ComClass(CalculateDecimalPlaces.ClassId, CalculateDecimalPlaces.InterfaceId, CalculateDecimalPlaces.EventsId), _
 ProgId("AssessorToolbar.CalculateDecimalPlaces")> _
Public NotInheritable Class CalculateDecimalPlaces
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "0a15656a-e189-4376-91dc-1cc453840d7b"
    Public Const InterfaceId As String = "62865f15-c856-4bb9-9e8d-0c7009cb9b95"
    Public Const EventsId As String = "4b27787c-4d0e-4112-afe8-9ab8982a8387"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region


    Private _application As IApplication
    Private _editor As IEditor

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "AssessorToolbar"  'localizable text 
        MyBase.m_caption = "CalculateDecimalPlaces"   'localizable text 
        MyBase.m_message = "Add the Dropped Decimal Places for Selected Features."   'localizable text 
        MyBase.m_toolTip = "Add the Dropped Decimal Places for Selected Features." 'localizable text 
        MyBase.m_name = MyBase.m_category & "_CalculateDecimalPlaces"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            'TODO: change bitmap name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

    End Sub


    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            _application = CType(hook, IApplication)
            _editor = _application.FindExtensionByName("ESRI Object Editor")

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If

        ' TODO:  Add other initialization code
    End Sub


    ''' <summary>
    ''' Called by ArcMap once per second to check if the command is enabled.
    ''' </summary>
    ''' <remarks>WARNING: Do not put computation-intensive code here.</remarks>
    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Dim canEnable As Boolean
            canEnable = _editor.EditState
            Return canEnable
        End Get
    End Property


    Public Overrides Sub OnClick()

        Dim updatedFeatureLayers As Integer

        Dim theMxDocument As IMxDocument = _application.Document
        Dim theMap As IMap = theMxDocument.FocusMap
        Dim theLayerEnum As IEnumLayer = theMap.Layers
        Dim thisLayer As ILayer = theLayerEnum.Next
        Dim theFeatureLayer As IFeatureLayer = Nothing
        Dim theFeatureSelection As IFeatureSelection = Nothing

        Do Until thisLayer Is Nothing
            If TypeOf thisLayer Is IFeatureLayer And thisLayer.Valid Then
                theFeatureLayer = thisLayer
                theFeatureSelection = theFeatureLayer
                If theFeatureSelection.SelectionSet.Count > 0 AndAlso (theFeatureLayer.FeatureClass.FindField("DISTANCE") + theFeatureLayer.FeatureClass.FindField("ARCLENGTH")) > 0 Then
                    UpdateFeatureLayer(theFeatureLayer)
                    updatedFeatureLayers += 1
                End If
            End If
            thisLayer = theLayerEnum.Next
        Loop

        If updatedFeatureLayers = 0 Then MessageBox.Show("No features were found to calculate.  Please make a selection on a layer that contains a ""DISTANCE"" or ""ARCLENGTH"" field.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)

    End Sub


    Sub UpdateFeatureLayer(ByVal theFeatureLayer As IFeatureLayer)


        Dim numberOfDecimals As String = InputBox("Enter the number of decimal places for the features in " & theFeatureLayer.Name & ":", "Calculate Missing Decimals", "2")
        If numberOfDecimals = String.Empty Then
            MessageBox.Show("Cancelling.  No updates were made.", "Operation Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        While Not IsNumeric(numberOfDecimals) OrElse CInt(numberOfDecimals) > 3 OrElse CInt(numberOfDecimals) < 0
            MessageBox.Show("Invalid entry.  Please enter a number between 0 and 3", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
            numberOfDecimals = InputBox("Enter the number of decimal places for the features in " & theFeatureLayer.Name & ":", "Calculate Missing Decimals", "2")
        End While

        If CInt(numberOfDecimals) <> 2 AndAlso MessageBox.Show("You have specified a number of decimals other than 2.  Are you sure this is correct?", "Decimal Confirmation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.Cancel Then Exit Sub

        Dim theFeatureSelection As IFeatureSelection = theFeatureLayer

        '-- Start edit operation (must start edit operation due to versioning)
        Dim theDataset As IDataset = theFeatureLayer.FeatureClass
        Dim theWorkspaceEdit As IWorkspaceEdit = theDataset.Workspace
        theWorkspaceEdit.StartEditOperation()

        Dim theFeatureCursor As IFeatureCursor = Nothing
        Dim theSelectionSet As ISelectionSet = theFeatureSelection.SelectionSet
        theSelectionSet.Search(Nothing, False, theFeatureCursor)

        Dim thisFeature As IFeature = theFeatureCursor.NextFeature
        Dim thisDistance As Object
        Dim thisArcLength As Object

        Do While Not thisFeature Is Nothing
            thisDistance = thisFeature.Value(thisFeature.Fields.FindField("DISTANCE"))
            thisArcLength = thisFeature.Value(thisFeature.Fields.FindField("ARCLENGTH"))

            If IsNumeric(thisDistance) Then
                thisFeature.Value(thisFeature.Fields.FindField("DISTANCE")) = FormatNumber(thisDistance, CInt(numberOfDecimals), vbFalse, vbFalse, vbFalse)
                thisFeature.Store()
            End If
            If IsNumeric(thisArcLength) Then
                thisFeature.Value(thisFeature.Fields.FindField("ARCLENGTH")) = FormatNumber(thisArcLength, CInt(numberOfDecimals), vbFalse, vbFalse, vbFalse)
                thisFeature.Store()
            End If

            thisFeature = theFeatureCursor.NextFeature
        Loop

        '-- Stop edit operation
        theWorkspaceEdit.StopEditOperation()

    End Sub


    'Private Function getFeatLayer(ByVal layerName As String) As IFeatureLayer

    '    '-- Get the document and map
    '    Dim theMxDocument As IMxDocument = _application.Document
    '    'If pDoc.FocusMap.LayerCount = 0 Then Exit Function

    '    Dim theMap As IMap = theMxDocument.FocusMap
    '    Dim theLayerEnum As IEnumLayer = theMap.Layers
    '    Dim thisLayer As ILayer = theLayerEnum.Next
    '    Dim theFeatureLayer As IFeatureLayer = Nothing

    '    Do Until thisLayer Is Nothing
    '        If TypeOf thisLayer Is IFeatureLayer And thisLayer.Valid Then
    '            If thisLayer.Name.ToUpper = layerName.ToUpper Then
    '                theFeatureLayer = thisLayer
    '                Exit Do
    '            End If
    '        End If
    '        thisLayer = theLayerEnum.Next
    '    Loop

    '    Return theFeatureLayer

    'End Function

End Class



