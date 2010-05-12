Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.carto
Imports System.Windows.forms
Imports AssessorToolbar.Utilities

<ComClass(DefinitionQuery.ClassId, DefinitionQuery.InterfaceId, DefinitionQuery.EventsId), _
 ProgId("AssessorToolbar.DefinitionQuery")> _
Public NotInheritable Class DefinitionQuery
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4f9a0151-f830-4e88-aeea-9594a848d8c0"
    Public Const InterfaceId As String = "585a1a12-df92-49ed-b019-000b18a45751"
    Public Const EventsId As String = "66f69b8f-a5a4-44f7-99e0-0f8eed51abb4"
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
    Private _buttonChecked As Boolean

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()


        ' TODO: Define values for the public properties
        MyBase.m_category = "AssessorToolbar"  'localizable text 
        MyBase.m_caption = "DefinitionQueryTool"   'localizable text 
        MyBase.m_message = "Sets and Clears Definition Queries."   'localizable text 
        MyBase.m_toolTip = "Sets and Clears Definition Queries." 'localizable text 
        MyBase.m_name = MyBase.m_category & "_DefinitionQueryTool"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If
        ' TODO:  Add other initialization code
    End Sub

    Public Overrides Sub OnClick()

        If GetFeatureLayerByName("SeeMaps", _application.Document) Is Nothing Then
            MessageBox.Show("Unable to find the SeeMaps feature class.  Please ensure it's loaded into your project", "Error", MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim theMxDoc As IMxDocument = _application.Document
        Dim theActiveView As IActiveView = theMxDoc.FocusMap

        _buttonChecked = Not _buttonChecked
        If _buttonChecked Then
            If setDefQuery() Then
                Me.m_checked = True
            Else
                _buttonChecked = Not _buttonChecked
            End If
        Else
            Me.m_checked = False
            clearDefQuery()
        End If

        theActiveView.Refresh()

    End Sub


    Private Function setDefQuery() As Boolean

        '-- Get the document and map
        Dim theMxDoc As IMxDocument = _application.Document
        If theMxDoc.FocusMap.LayerCount = 0 Then Exit Function

        '-- Get the mapnumber
        Dim theSelectMapIndexDialog As SelectMapindexDialog = MakeSelectMapIndexDialog(_application.Document)
        If theSelectMapIndexDialog.ShowDialog = DialogResult.Cancel Then Exit Function
        Dim theMapNumber As String = theSelectMapIndexDialog.MapNumber

        Dim theMap As IMap = theMxDoc.FocusMap
        Dim theEnumLayer As IEnumLayer = theMap.Layers

        Dim theLayer As ILayer
        Dim theFLayer As IFeatureLayer

        theLayer = theEnumLayer.Next
        Do Until theLayer Is Nothing
            If TypeOf theLayer Is IFeatureLayer And theLayer.Valid Then
                theFLayer = theLayer

                If theFLayer.FeatureClass.FeatureType = ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTAnnotation Or _
                    theFLayer.Name.ToUpper = "CORNER" Or theFLayer.Name.ToUpper = "REFERENCELINES" Or _
                    theFLayer.Name.ToUpper = "MAPSECLINES" Then

                    If theFLayer.FeatureClass.FindField("mapnumber") > 0 Then
                        setTheDefQuery(theFLayer, "mapnumber = '" & theMapNumber & "'")
                    End If
                End If

            End If

            theLayer = theEnumLayer.Next
        Loop

        Return True

    End Function

    Private Sub clearDefQuery()

        '-- Get the document and map
        Dim theMxDoc As IMxDocument = _application.Document

        If theMxDoc.FocusMap.LayerCount = 0 Then Exit Sub
        Dim theMap As IMap = theMxDoc.FocusMap

        Dim theEnumlayer As IEnumLayer = theMap.Layers

        Dim theLayer As ILayer
        Dim theFLayer As IFeatureLayer

        theLayer = theEnumlayer.Next
        Do Until theLayer Is Nothing
            If TypeOf theLayer Is IFeatureLayer And theLayer.Valid Then
                theFLayer = theLayer
                If theFLayer.FeatureClass.FeatureType = ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTAnnotation Or _
                    theFLayer.Name.ToUpper = "CORNER" Or theFLayer.Name.ToUpper = "REFERENCELINES" Or _
                    theFLayer.Name.ToUpper = "MAPSECLINES" Then
                    Call setTheDefQuery(theFLayer, "")
                End If
            End If

            theLayer = theEnumlayer.Next
        Loop


    End Sub


    Private Sub setTheDefQuery(ByVal theFeatureLayer As IFeatureLayer, ByVal theDefQuery As String)

        If Not theFeatureLayer Is Nothing Then
            Dim theFeatureLayDef As IFeatureLayerDefinition
            theFeatureLayDef = theFeatureLayer
            theFeatureLayDef.DefinitionExpression = theDefQuery
        End If

        '-- Cleanup
        theFeatureLayer = Nothing

    End Sub


End Class



