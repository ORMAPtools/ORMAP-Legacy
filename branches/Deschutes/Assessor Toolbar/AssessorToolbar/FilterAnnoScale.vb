Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports System.Windows.forms
Imports AssessorToolbar.Utilities
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.Display


<ComClass(FilterAnnoScale.ClassId, FilterAnnoScale.InterfaceId, FilterAnnoScale.EventsId), _
 ProgId("AssessorToolbar.FilterAnnoScale")> _
Public NotInheritable Class FilterAnnoScale
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "60627e9c-9731-4ba0-ba78-5ff6bbef2024"
    Public Const InterfaceId As String = "312143b8-6e0f-467f-8a7b-4eeed1123d5c"
    Public Const EventsId As String = "cd469290-a3e9-414e-995e-57da3e1642dc"
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
    Private _mapnumber As String

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "AssessorToolbar"  'localizable text 
        MyBase.m_caption = "FilterAnnoScale"   'localizable text 
        MyBase.m_message = "Loads the specified annotation and sets environment for editing scale."   'localizable text 
        MyBase.m_toolTip = "Loads the specified annotation and sets environment for editing scale." 'localizable text 
        MyBase.m_name = MyBase.m_category & "_FilterAnnoScale"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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


    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Dim canEnable As Boolean
            canEnable = Not _editor.EditState
            Return canEnable
        End Get
    End Property


    Public Overrides Sub OnClick()

        Dim theSelectMapScaleDialog As FilterAnnoScaleForm = New FilterAnnoScaleForm
        If theSelectMapScaleDialog.ShowDialog = DialogResult.Cancel Then Exit Sub
        Dim theMapScale As Integer = Convert.ToInt32(theSelectMapScaleDialog.uxMapScale.Text) * 12
        ReferenceScale = theMapScale

        Dim theAnnoNameScale As String = theSelectMapScaleDialog.uxMapScale.Text
        Do Until theAnnoNameScale.Length = 4
            theAnnoNameScale = "0" & theAnnoNameScale
        Loop

        Dim theAnnoOffset As String = String.Empty

        Select Case theSelectMapScaleDialog.uxMapScale.Text
            Case "10"
                theAnnoOffset = ".13"

            Case "20"
                theAnnoOffset = ".1"

            Case "30"
                theAnnoOffset = ".15"

            Case "40"
                theAnnoOffset = ".2"

            Case "50"
                theAnnoOffset = ".23"

            Case "60"
                theAnnoOffset = ".28"

            Case "100"
                theAnnoOffset = ".5"

            Case "200"
                theAnnoOffset = "1"

            Case "400"
                theAnnoOffset = "1.9"

            Case "2000"
                theAnnoOffset = "9"

        End Select

        Dim theMxDoc As IMxDocument = _application.Document
        Dim theMap As IMap = theMxDoc.FocusMap
        Dim theActiveView As IActiveView = theMxDoc.FocusMap

        If theMap.ReferenceScale <> 0 Then
            theMap.ReferenceScale = theMapScale
        End If

        '-- Remove existing anno feature classes
        Dim theLayer As ILayer = GetSDEAnnoFeature(theMap)
        Do While Not theLayer Is Nothing
            theMap.DeleteLayer(theLayer)
            theLayer = GetSDEAnnoFeature(theMap)
        Loop

        Call LoadSDEAnnoFeatures("Anno" & theAnnoNameScale & "scale", theMxDoc)
        Call LoadSDEAnnoFeatures("Arrow" & theAnnoNameScale & "scale", theMxDoc)

        theMxDoc.UpdateContents()
        theActiveView.Refresh()

        MessageBox.Show("Set your Annotation Offset to " & theAnnoOffset & " for " & theSelectMapScaleDialog.uxMapScale.Text & " scale maps.", "Annotation Offset", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub


    Sub LoadSDEAnnoFeatures(ByVal annoName As String, ByVal theMxDoc As IMxDocument)

        Dim theSeeMapsFlayer As IFeatureLayer = GetFeatureLayerByName("SeeMaps", theMxDoc)
        If theSeeMapsFlayer Is Nothing Then
            MessageBox.Show("Unable to find the SeeMaps feature class.  This feature class is required.  Please load it in and ensure it's name is 'SeeMaps'.", "Missing Data", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        Dim theDataset As IDataset = theSeeMapsFlayer.FeatureClass

        Dim theFDOGraphicsLayerFactory As IFDOGraphicsLayerFactory = New FDOGraphicsLayerFactory
        Dim theFLayer As IFeatureLayer = theFDOGraphicsLayerFactory.OpenGraphicsLayer(theDataset.Workspace, Nothing, "giscarto.CREATOR_ASR." + annoName)
        theFLayer.Name = annoName

        Dim theMap As IMap = theMxDoc.FocusMap
        Dim theGrouplayer As IGroupLayer = Nothing
        For i As Integer = 0 To theMap.LayerCount - 1
            If TypeOf theMap.Layer(i) Is IGroupLayer AndAlso theMap.Layer(i).Name = "Annotation" Then
                theGrouplayer = theMap.Layer(i)
                Exit For
            End If
        Next i

        If theGrouplayer Is Nothing Then
            MessageBox.Show("Could not find the Group Layer 'Annotation' in the Table of Contents", "Missing Group Layer", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        theGrouplayer.Add(theFLayer)

        Dim theCompositeLayer As ICompositeLayer2
        theCompositeLayer = theFLayer
        theCompositeLayer.Expanded = False

        Dim theRGBColorAnno As IRgbColor = New RgbColor
        With theRGBColorAnno
            .Red = 204
            .Green = 204
            .Blue = 204
        End With

        Dim theSymbolSubstitution As ISymbolSubstitution
        theSymbolSubstitution = theFLayer

        theSymbolSubstitution.SubstituteType = esriSymbolSubstituteType.esriSymbolSubstituteColor
        theSymbolSubstitution.MassColor = theRGBColorAnno

    End Sub

    Function GetSDEAnnoFeature(ByVal theMap As IMap) As ILayer

        Dim theLayer As ILayer = Nothing

        For i As Integer = 0 To theMap.LayerCount - 1
            If TypeOf theMap.Layer(i) Is ICompositeLayer And theMap.Layer(i).Name = "Annotation" Then
                Dim pCompositeLayer As ICompositeLayer = theMap.Layer(i)
                For j As Integer = 0 To pCompositeLayer.Count - 1
                    If pCompositeLayer.Layer(j).Name.Contains("0scale") Then
                        theLayer = pCompositeLayer.Layer(j)
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next

        Return theLayer

    End Function





End Class



