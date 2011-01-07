#Region "Imported Namespaces"

Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Environment
Imports System.Globalization
Imports System.Drawing.Text
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Display

Imports ESRI.ArcGIS.SystemUI
Imports OrmapTaxlotEditing.DataMonitor
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.EditorExtension
Imports OrmapTaxlotEditing.Utilities

#End Region

<ComClass(SaveCogo.ClassId, SaveCogo.InterfaceId, SaveCogo.EventsId), _
 ProgId("OrmapTaxlotEditing.SaveCogo")> _
Public NotInheritable Class SaveCogo
    Inherits BaseCommand
    Implements IDisposable

#Region "Class-Level Constants and Enumerations"

    Dim theDataCollection As Collection = New Collection
    Dim distanceIndex As Integer
    Dim oIDIndex As Integer
    Dim tangentIndex As Integer
    Dim directionIndex As Integer

    Structure cogoData
        Dim featureClassName As String
        Dim oid As Integer
        Dim distance As String
        Dim direction As String
        Dim shapeLength As Double
        Dim tangent2 As String
    End Structure


#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"
    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = ""  'localizable text 
        MyBase.m_caption = ""   'localizable text 
        MyBase.m_message = ""   'localizable text 
        MyBase.m_toolTip = "" 'localizable text 
        MyBase.m_name = ""  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            'TODO: change bitmap name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try
    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication
    Private _bitmapResourceName As String

#End Region

#Region "Properties"

    Private WithEvents _partnerSaveCogoForm As SaveCogoForm

    Friend ReadOnly Property PartnerSaveCogoForm() As SaveCogoForm
        Get
            If _partnerSaveCogoForm Is Nothing OrElse _partnerSaveCogoForm.IsDisposed Then
                setSaveCogoForm(New SaveCogoForm())
            End If
            Return _partnerSaveCogoForm
        End Get
    End Property

#End Region

#Region "Event Handlers"

    Private Sub setSaveCogoForm(ByVal value As SaveCogoForm)
        If value IsNot Nothing Then
            _partnerSaveCogoForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerSaveCogoForm.uxCogoSave.Click, AddressOf uxCogoSave_Click
            AddHandler _partnerSaveCogoForm.uxCogoProportion.Click, AddressOf uxCogoProportion_Click
            AddHandler _partnerSaveCogoForm.uxCogoHelp.Click, AddressOf uxCogoHelp_Click
            AddHandler _partnerSaveCogoForm.uxCogoQuit.Click, AddressOf uxCogoQuit_Click
        Else
            ' Unsubscribe to partner form events.
            RemoveHandler _partnerSaveCogoForm.uxCogoSave.Click, AddressOf uxCogoSave_Click
            RemoveHandler _partnerSaveCogoForm.uxCogoProportion.Click, AddressOf uxCogoProportion_Click
            RemoveHandler _partnerSaveCogoForm.uxCogoHelp.Click, AddressOf uxCogoHelp_Click
            RemoveHandler _partnerSaveCogoForm.uxCogoQuit.Click, AddressOf uxCogoQuit_Click
        End If
    End Sub

    Friend Sub DoButtonOperation()
        Try
            Dim theMxDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
            Dim theMap As IMap = theMxDoc.FocusMap
            'The EditorExtension.EditEvents_OnChangeFeature method checks on Taxlot, CancelledNumbersTable, and
            'MapIndex. Once transferred there, the edit operation started by uxCogoSave_Click cause a failure if the
            'user tries to use one of the EditorExtension's load feature operations. Therefore, check for these 
            'data now and abort the Save Cogo until the user loads them (even through they are not really needed). 

            'Check for:
            ' Taxlot status and layer properties
            ' MapIndex status and layer properties
            ' CancelledNumbersTable status and table properties
            ' TaxlotLines status and layer properties

            'DataMonitor.CheckValidTaxlotDataProperties()
            'If Not HasValidTaxlotData Then
            '    MessageBox.Show("Missing data: Valid ORMAP Taxlot layer not found in the map." & NewLine & _
            '                    "Please load this dataset into your map.", _
            '                    "Save Cogo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            '    Exit Sub
            'End If

            'DataMonitor.CheckValidMapIndexDataProperties()
            'If Not HasValidMapIndexData Then
            '    MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & NewLine & _
            '                    "Please load this dataset into your map.", _
            '                    "Save Cogo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            '    Exit Sub
            'End If

            'DataMonitor.CheckValidCancelledNumbersTableDataProperties()
            'If Not HasValidCancelledNumbersTableData Then
            '    MessageBox.Show("Missing data: Valid ORMAP CancelledNumbersTable not found in the map." & NewLine & _
            '                    "Please load this dataset into your map.", _
            '                    "Save Cogo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            '    Exit Sub
            'End If

            DataMonitor.CheckValidTaxlotLinesDataProperties()
            If Not HasValidTaxlotLinesData Then
                MessageBox.Show("Missing data: Valid ORMAP TaxlotLines layer not found in the map." & NewLine & _
                                "Please load this dataset into your map.", _
                                "Save Cogo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub 
            End If

            If theMap.SelectionCount < 1 Then
                MessageBox.Show("Missing data: No line features have been selected." & NewLine & _
                                "Please select at least one line feature which has." & NewLine & _
                                "Distance and Direction attributes.", _
                                "Save Cogo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            End If

            PartnerSaveCogoForm.Show()

        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try

    End Sub

    Private Sub uxCogoSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim theMxDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
        Dim theMap As IMap = theMxDoc.FocusMap
        Dim theActiveView As IActiveView = CType(theMap, IActiveView)
        Dim theMapExtent As IEnvelope = theActiveView.Extent

        Dim theEnumFeature As IEnumFeature = CType(theMap.FeatureSelection, IEnumFeature)
        Dim theEnumFeatureSetup As IEnumFeatureSetup = CType(theEnumFeature, IEnumFeatureSetup)
        theEnumFeatureSetup.AllFields = True
        theDataCollection.Clear()
        Dim thisFeature As IFeature = theEnumFeature.Next
        Dim theFeatureClass As IFeatureClass
        theFeatureClass = CType(thisFeature.Class, IFeatureClass)
        distanceIndex = theFeatureClass.FindField("Distance")
        oIDIndex = theFeatureClass.FindField("OBJECTID")
        tangentIndex = theFeatureClass.FindField("Tangent2")
        directionIndex = theFeatureClass.FindField("Direction")
        Dim geometryDistance As Integer = theFeatureClass.FindField(theFeatureClass.LengthField.Name)

        EditorExtension.Editor.StartOperation()

        Do While Not thisFeature Is Nothing
            If Not thisFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                MessageBox.Show("Wrong Type: A feature was selected which is NOT a polyline feature." & NewLine & _
                                "Feature " & thisFeature.OID & " will not be processed.", _
                                "Save Cogo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ElseIf thisFeature.Fields.FindField("Direction") < 0 Or thisFeature.Fields.FindField("Distance") < 0 _
                Or thisFeature.Fields.FindField("Tangent2") < 0 Then
                MessageBox.Show("Missing data: Tangent2, Direction and/or Distance attributes are missing" & NewLine & _
                                "from the selected feature. Feature " & thisFeature.OID & " will not be processed.", _
                                "Save Cogo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            ElseIf thisFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Then
                Try
                    Dim theCogoData As cogoData = New cogoData
                    theCogoData.oid = CInt(thisFeature.Value(oIDIndex))
                    theCogoData.shapeLength = CDbl(thisFeature.Value(geometryDistance))
                    theCogoData.featureClassName = thisFeature.Class.AliasName

                    If Not IsDBNull(thisFeature.Value(directionIndex)) Then
                        theCogoData.direction = CStr(thisFeature.Value(directionIndex))
                    ElseIf IsDBNull(thisFeature.Value(directionIndex)) Then
                        MessageBox.Show("Missing data: Direction attribute is <Null> or missing" & NewLine & _
                                        "from the selected feature. Feature " & thisFeature.OID & " will not be processed.", _
                                        "Save Cogo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        EditorExtension.Editor.AbortOperation()
                        Exit Do
                    End If

                    If Not IsDBNull(thisFeature.Value(distanceIndex)) Then
                        theCogoData.distance = CStr(thisFeature.Value(distanceIndex))
                    ElseIf IsDBNull(thisFeature.Value(distanceIndex)) Then
                        MessageBox.Show("Missing data: Distance attribute is <Null> or missing" & NewLine & _
                                        "from the selected feature. Feature " & thisFeature.OID & " will not be processed.", _
                                        "Save Cogo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        EditorExtension.Editor.AbortOperation()
                        Exit Do
                    End If

                    If Not IsDBNull(thisFeature.Value(tangentIndex)) Then
                        theCogoData.tangent2 = CStr(thisFeature.Value(tangentIndex))
                    ElseIf IsDBNull(thisFeature.Value(tangentIndex)) Then
                        theCogoData.tangent2 = Nothing
                    End If
                    theDataCollection.Add(theCogoData)
                    'Set the Tangent2 field to be the OID (will become parent ID for new features)
                    thisFeature.Value(tangentIndex) = thisFeature.Value(oIDIndex)
                    thisFeature.Store()
                Catch ex As Exception
                    EditorExtension.ProcessUnhandledException(ex)
                End Try
            End If
            thisFeature = theEnumFeature.Next
        Loop

        EditorExtension.Editor.StopOperation("Save COGO attributes")

        PartnerSaveCogoForm.uxCogoSave.Enabled = False
        PartnerSaveCogoForm.uxCogoProportion.Enabled = True
    End Sub

    Private Sub uxCogoProportion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim parentCogoData As cogoData
        Dim thisFeatureLayer As IFeatureLayer
        Dim thisFeatureClass As IFeatureClass
        Dim i As Integer = 1, n As Integer = 1

        For i = 1 To theDataCollection.Count
            parentCogoData = CType(theDataCollection.Item(i), cogoData)
            thisFeatureLayer = SpatialUtilities.FindFeatureLayerByDSName(parentCogoData.featureClassName)
            thisFeatureClass = thisFeatureLayer.FeatureClass
            Dim distanceIndex As Integer = thisFeatureClass.FindField("Distance")
            Dim oIDIndex As Integer = thisFeatureClass.FindField("OBJECTID")
            Dim tangentIndex As Integer = thisFeatureClass.FindField("Tangent2")
            Dim directionIndex As Integer = thisFeatureClass.FindField("Direction")
            Dim geometryDistance As Integer = thisFeatureClass.FindField(thisFeatureClass.LengthField.Name)

            Dim queryFilter As IQueryFilter = New QueryFilter
            queryFilter.WhereClause = SpatialUtilities.formatWhereClause(String.Concat("Tangent2 = '", parentCogoData.oid, "'"), thisFeatureClass)
            queryFilter.SubFields = "*"
            Dim featureCursor As IFeatureCursor = thisFeatureClass.Update(queryFilter, False)
            Dim count As Integer = thisFeatureClass.FeatureCount(queryFilter)
            Dim childFeature As IFeature

            If count > 1 Then
                'If more than one feature, then child has been created so lines must be proportioned
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Proportion using ratio formula based on:
                ' p = Parent feature
                ' c = Child feature
                ' L = Legal (COGO) distance
                ' G = Geometry (Map) distance
                '
                ' cL   pL      cL = (pL)(cG)
                ' -- = --  ==       --------
                ' cG   pG              pG
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                EditorExtension.Editor.StartOperation()
                For n = 1 To count
                    childFeature = featureCursor.NextFeature
                    'Set tangent2 value back to original value from parent
                    Try
                        If parentCogoData.tangent2 Is Nothing Then
                            childFeature.Value(tangentIndex) = System.DBNull.Value
                        Else
                            childFeature.Value(tangentIndex) = parentCogoData.tangent2
                        End If
                        childFeature.Value(directionIndex) = parentCogoData.direction
                        childFeature.Value(distanceIndex) = CStr(Math.Round((CDbl(childFeature.Value(geometryDistance)) * CDbl(parentCogoData.distance)) / parentCogoData.shapeLength, _
                                                                       2, MidpointRounding.AwayFromZero))
                        childFeature.Store()
                    Catch ex As Exception
                        EditorExtension.ProcessUnhandledException(ex)
                    End Try

                Next n
                EditorExtension.Editor.StopOperation("Proportion COGO Distance(s)")
            End If
            If count = 1 Then
                'If only one feature, check to see if Distance and Direction fields are NULL
                'NOTE=> Features which originally have no Distance and Direction are NOT 
                'added to the collection, so this means that an edit tool has nulled out the
                'original values. Need to reset them.
                childFeature = featureCursor.NextFeature
                If IsDBNull(childFeature.Value(directionIndex)) Then
                    EditorExtension.Editor.StartOperation()

                    Try
                        childFeature.Value(directionIndex) = parentCogoData.direction
                        childFeature.Value(distanceIndex) = parentCogoData.distance
                    Catch ex As Exception
                        EditorExtension.ProcessUnhandledException(ex)
                    End Try

                End If
                If parentCogoData.tangent2 Is Nothing Then
                    childFeature.Value(tangentIndex) = System.DBNull.Value
                Else
                    childFeature.Value(tangentIndex) = parentCogoData.tangent2
                End If
                childFeature.Store()
                EditorExtension.Editor.StopOperation("Reset COGO Attributes")
            End If
            'Release the update cursor to remove the lock on the input data.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(featureCursor)
        Next i
        theDataCollection.Clear()
        PartnerSaveCogoForm.uxCogoProportion.Enabled = False
        PartnerSaveCogoForm.uxCogoSave.Enabled = True
    End Sub

    Private Sub uxCogoHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim theRTFStream As System.IO.Stream = Me.GetType().Assembly.GetManifestResourceStream("OrmapTaxlotEditing.SaveCogo_help.rtf")
        OpenHelp("Save COGO Help", theRTFStream)
    End Sub

    Private Sub uxCogoQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If theDataCollection.Count > 0 Then
            Dim parentCogoData As cogoData
            Dim thisFeatureLayer As IFeatureLayer
            Dim thisFeatureClass As IFeatureClass
            Dim i As Integer = 1, n As Integer = 1

            For i = 1 To theDataCollection.Count
                parentCogoData = CType(theDataCollection.Item(i), cogoData)
                thisFeatureLayer = SpatialUtilities.FindFeatureLayerByDSName(parentCogoData.featureClassName)
                thisFeatureClass = thisFeatureLayer.FeatureClass
                Dim tangentIndex As Integer = thisFeatureClass.FindField("Tangent2")

                Dim queryFilter As IQueryFilter = New QueryFilter
                queryFilter.WhereClause = SpatialUtilities.formatWhereClause(String.Concat("Tangent2 = '", parentCogoData.oid, "'"), thisFeatureClass)
                queryFilter.SubFields = "*"
                Dim featureCursor As IFeatureCursor = thisFeatureClass.Update(queryFilter, False)
                Dim count As Integer = thisFeatureClass.FeatureCount(queryFilter)
                Dim childFeature As IFeature

                If count > 0 Then
                    EditorExtension.Editor.StartOperation()
                    Try
                        childFeature = featureCursor.NextFeature
                        If parentCogoData.tangent2 Is Nothing Then
                            childFeature.Value(tangentIndex) = System.DBNull.Value
                        Else
                            childFeature.Value(tangentIndex) = parentCogoData.tangent2
                        End If
                        childFeature.Store()
                    Catch ex As Exception
                        EditorExtension.ProcessUnhandledException(ex)
                    End Try
                    EditorExtension.Editor.StopOperation("Resetting COGO temp field")
                End If
            Next
        End If
        PartnerSaveCogoForm.Close()
    End Sub

#Region "Methods"

#End Region

#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"

        ''' <summary>
        ''' Called by ArcMap once per second to check if the command is enabled.
        ''' </summary>
        ''' <remarks>WARNING: Do not put computation-intensive code here.</remarks>
    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
        Dim canEnable As Boolean
        canEnable = EditorExtension.CanEnableExtendedEditing
        canEnable = canEnable AndAlso EditorExtension.Editor.EditState = esriEditState.esriStateEditing
        canEnable = canEnable AndAlso EditorExtension.IsValidWorkspace
        Return canEnable
        End Get
    End Property

#End Region

#Region "Methods"

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
        DoButtonOperation()
    End Sub
#End Region

#End Region

#Region "Implemented Interface Properties"

#Region "IDisposable Interface Implementation"

        Private _isDuringDispose As Boolean ' Used to track whether Dispose() has been called and is in progress.

        ''' <summary>
        ''' Dispose of managed and unmanaged resources.
        ''' </summary>
        ''' <param name="disposing">True or False.</param>
        ''' <remarks>
        ''' <para>Member of System::IDisposable.</para>
        ''' <para>Dispose executes in two distinct scenarios. 
        ''' If disposing equals true, the method has been called directly
        ''' or indirectly by a user's code. Managed and unmanaged resources
        ''' can be disposed.</para>
        ''' <para>If disposing equals false, the method has been called by the 
        ''' runtime from inside the finalizer and you should not reference 
        ''' other objects. Only unmanaged resources can be disposed.</para>
        ''' </remarks>
    Friend Sub Dispose(ByVal disposing As Boolean)
        ' Check to see if Dispose has already been called.
        If Not Me._isDuringDispose Then

            ' Flag that disposing is in progress.
            Me._isDuringDispose = True

            If disposing Then
                ' Free managed resources when explicitly called.

                ' Dispose managed resources here.
                '   e.g. component.Dispose()

            End If

            ' Free "native" (shared unmanaged) resources, whether 
            ' explicitly called or called by the runtime.

            ' Call the appropriate methods to clean up 
            ' unmanaged resources here.
            _bitmapResourceName = Nothing
            MyBase.m_bitmap = Nothing

            ' Flag that disposing has been finished.
            _isDuringDispose = False

        End If

    End Sub

#Region " IDisposable Support "

        ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

#End Region

#End Region

#Region "Other Members"

#Region "COM GUIDs"
        ' These  GUIDs provide the COM identity for this class 
        ' and its COM interfaces. If you change them, existing 
        ' clients will no longer be able to access the class.
        Public Const ClassId As String = "95726933-ca77-434a-b20c-1b5d4757fc0a"
        Public Const InterfaceId As String = "56db7053-87b9-496a-b532-e419d79f414c"
        Public Const EventsId As String = "3a7deb66-0f5d-47a5-a629-9b794bcc4871"
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

#End Region

End Class



