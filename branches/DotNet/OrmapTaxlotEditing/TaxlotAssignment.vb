#Region "Copyright 2008 ORMAP Tech Group"

' File:  TaxlotAssignment.vb
'
' Original Author:  OPET.NET Migration Team (Shad Campbell, James Moore, 
'                   Nick Seigal)
'
' Date Created:  January 8, 2008
'
' Copyright Holder:  ORMAP Tech Group  
' Contact Info:  ORMAP Tech Group (a.k.a. opet developers) may be reached at 
' opet-developers@lists.sourceforge.net
'
' This file is part of the ORMAP Taxlot Editing Toolbar.
'
' ORMAP Taxlot Editing Toolbar is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License as published by 
' the Free Software Foundation; either version 3 of the License, or (at your 
' option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT 
' ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or 
' FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License located
' in the COPYING.txt file for more details.
'
' You should have received a copy of the GNU General Public License along
' with the ORMAP Taxlot Editing Toolbar; if not, write to the Free Software 
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.

#End Region

#Region "Subversion Keyword Expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision$
'Date of Last Change: $Date$
#End Region

#Region "Imported Namespaces"
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports OrmapTaxlotEditing.DataMonitor
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.StringUtilities
Imports OrmapTaxlotEditing.Utilities
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

#End Region

<ComVisible(True)> _
<ComClass(TaxlotAssignment.ClassId, TaxlotAssignment.InterfaceId, TaxlotAssignment.EventsId), _
ProgId("ORMAPTaxlotEditing.TaxlotAssignment")> _
Public NotInheritable Class TaxlotAssignment
    Inherits BaseTool

#Region "Class-Level Constants And Enumerations"

    ' Taxlot number type constants
    ' NOTE: these must be exactly 5 characters long
    Private Const taxlotNumberTypeTaxlot As String = "TAXLOT" 'normal taxlot number
    Private Const taxlotNumberTypeRoads As String = "ROADS"
    Private Const taxlotNumberTypeWater As String = "WATER"
    Private Const taxlotNumberTypeRails As String = "RAILS"
    Private Const taxlotNumberTypeNontaxlot As String = "NONTL"

    Private Const defaultCommand As String = "esriArcMapUI.SelectTool"

    Private Enum StatePassageType As Integer
        Entering = 1
        Exiting = 2
    End Enum

    Public Enum CommandStateType As Integer
        Enabled = 11
        Disabled = 12
    End Enum

#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' Define protected instance field values for the public properties.
        MyBase.m_category = "OrmapToolbar"  'localizable text 
        MyBase.m_caption = "TaxlotAssignment"   'localizable text 
        MyBase.m_message = "Populate values in the Taxlots feature class based on a starting value and an increment value."   'localizable text 
        MyBase.m_toolTip = "Assign Taxlots" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_TaxlotAssignment"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            ' Set the bitmap based on the name of the class.
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New System.Drawing.Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As ArgumentException
            Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

        Try
            ' Set the (enabled) cursor based on the name of the class.
            Dim cursorResourceName As String = Me.GetType().Name + ".cur"
            MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), cursorResourceName)
        Catch ex As ArgumentException
            Trace.WriteLine(ex.Message, "Invalid Cursor")
        End Try

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication

#End Region

#Region "Properties"

    'Private _state As CommandStateType = CommandStateType.Disabled

    'Public ReadOnly Property State() As CommandStateType
    '    Get
    '        Return _state
    '    End Get
    'End Property

    'Private Sub setState(ByVal stateType As CommandStateType)
    '    _state = stateType
    'End Sub

    'Private _canCheckIfEnabled As Boolean

    'Public Property CanCheckIfEnabled() As Boolean
    '    Set(ByVal value As Boolean)
    '        _canCheckIfEnabled = value
    '    End Set
    '    Get
    '        Return _canCheckIfEnabled
    '    End Get
    'End Property

    Private _incrementNumber As Integer

    Public ReadOnly Property IncrementNumber() As Integer
        Get
            _incrementNumber = CInt(PartnerTaxlotAssignmentForm.uxType.SelectedItem.ToString)
            Return _incrementNumber
        End Get
    End Property

    Private _taxlotType As String

    Public ReadOnly Property TaxlotType() As String
        Get
            _taxlotType = PartnerTaxlotAssignmentForm.uxType.SelectedItem.ToString
            Return _taxlotType
        End Get
    End Property

    Private _numberStartingFrom As Integer

    Public Property NumberStartingFrom() As Integer
        Get
            _numberStartingFrom = CInt(PartnerTaxlotAssignmentForm.uxStartingFrom.Text)
            Return _numberStartingFrom
        End Get
        Set(ByVal value As Integer)
            _numberStartingFrom = value
            PartnerTaxlotAssignmentForm.uxStartingFrom.Text = CStr(_numberStartingFrom)
        End Set
    End Property

    Private WithEvents _partnerTaxlotAssignmentForm As TaxlotAssignmentForm  ' TODO: [NIS] Is WithEvents needed here?

    Friend ReadOnly Property PartnerTaxlotAssignmentForm() As TaxlotAssignmentForm
        Get
            If _partnerTaxlotAssignmentForm Is Nothing OrElse _partnerTaxlotAssignmentForm.IsDisposed Then
                setPartnerTaxlotAssignmentForm(New TaxlotAssignmentForm())
            End If
            Return _partnerTaxlotAssignmentForm
        End Get
    End Property

    Private Sub setPartnerTaxlotAssignmentForm(ByRef value As TaxlotAssignmentForm)
        If value IsNot Nothing Then
            _partnerTaxlotAssignmentForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerTaxlotAssignmentForm.Load, AddressOf PartnerTaxlotAssignmentForm_Load
            AddHandler _partnerTaxlotAssignmentForm.uxHelp.Click, AddressOf uxHelp_Click
            AddHandler _partnerTaxlotAssignmentForm.uxType.SelectedValueChanged, AddressOf uxType_SelectedValueChanged
        End If
    End Sub

#End Region

#Region "Event Handlers"

#Region "Partner Form Event Handlers"

    Private Sub PartnerTaxlotAssignmentForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles PartnerTaxlotAssignmentForm.Load

        With PartnerTaxlotAssignmentForm

            'Populate multi-value controls
            .uxType.Items.Add(taxlotNumberTypeTaxlot)
            .uxType.Items.Add(taxlotNumberTypeRoads)
            .uxType.Items.Add(taxlotNumberTypeWater)
            .uxType.Items.Add(taxlotNumberTypeRails)
            .uxType.Items.Add(taxlotNumberTypeNontaxlot)

            ' Set control defaults
            .uxType.Text = taxlotNumberTypeTaxlot
            .uxIncrementByNone.Checked = True
            .uxStartingFrom.Text = "0"

            ' Enable the numbering settings controls by enabling the group
            .uxTaxlotNumberingOptions.Enabled = True
            'With .uxStartingFrom
            '    '.BackColor = System.Drawing.Color.White
            '    .Enabled = True
            'End With
            '.uxIncrementByNone.Enabled = True
            '.uxIncrementBy1.Enabled = True
            '.uxIncrementBy10.Enabled = True
            '.uxIncrementBy100.Enabled = True
            '.uxIncrementBy1000.Enabled = True
        End With


    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles TaxlotAssignmentForm.uxHelp.Click
        ' TODO: [NIS] Could be replaced with new help mechanism.
        ' Open a custom help file.
        ' Note: Requires a specific file in the help subdirectory of the application directory.
        Dim filePath As String
        filePath = My.Application.Info.DirectoryPath & "\help\videos\TaxlotAssignment\TaxlotAssignment.html"
        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(filePath) Then
            ' Open the help form.
            Dim helpForm As New HelpForm
            helpForm.Text = "Taxlot Assignment Help"
            Dim theUri As New System.Uri("file:///" & filePath)
            helpForm.WebBrowser1.Url = theUri
            helpForm.Width = 1000
            helpForm.Height = 740
            helpForm.Show()
        Else
            MessageBox.Show("No help file available in the directory " & vbNewLine & _
                    My.Application.Info.DirectoryPath & "\help\videos\TaxlotAssignment" & ".")
        End If

    End Sub

    Private Sub uxType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles TaxlotAssignmentForm.uxType.SelectedValueChanged
        With PartnerTaxlotAssignmentForm
            Const NoSelectedIndex As Integer = -1
            If .uxType.SelectedIndex <> NoSelectedIndex Then
                If .uxType.SelectedItem.ToString = taxlotNumberTypeTaxlot Then
                    ' Enable the numbering settings controls by enabling the group
                    .uxTaxlotNumberingOptions.Enabled = True
                Else
                    ' Disable the numbering settings controls by disabling the group
                    .uxTaxlotNumberingOptions.Enabled = False
                End If
            End If
        End With

    End Sub

#End Region

    '#Region "EditEvents Event Handlers"
    '
    '    Private Sub EditEvents_OnStartEditing() 'Implements ESRI.ArcGIS.Editor.IEditEvents.OnStartEditing
    '        ' State Transistion E1
    '        TransitionE1()
    '    End Sub
    '
    '    Private Sub EditEvents_OnStopEditing(ByVal save As Boolean) 'Implements ESRI.ArcGIS.Editor.IEditEvents.OnStopEditing
    '        ' State Transistion E2
    '        TransitionE2()
    '    End Sub
    '
    '#End Region
    '
    '#Region "ActiveViewEvents Event Handlers"
    '
    '    Public Sub ActiveViewEvents_FocusMapChanged() 'Implements ESRI.ArcGIS.Carto.IActiveViewEvents.FocusMapChanged
    '        ' State Transistion E3
    '        TransitionE3()
    '    End Sub
    '
    '    Public Sub ActiveViewEvents_ItemAdded(ByVal Item As Object) 'Implements ESRI.ArcGIS.Carto.IActiveViewEvents.ItemAdded
    '        ' State Transistion E3
    '        TransitionE3()
    '    End Sub
    '
    '    Public Sub ActiveViewEvents_ItemDeleted(ByVal Item As Object) 'Implements ESRI.ArcGIS.Carto.IActiveViewEvents.ItemDeleted
    '        ' State Transistion E3
    '        TransitionE3()
    '    End Sub
    '
    '#End Region

#End Region

#Region "Methods"

    'Private Function hasRequiredData() As Boolean

    '    ' TEMPLATE: Const theFCName1 As String = "FeatureClassName1"  'TODO: Insert real fc name
    '    ' TEMPLATE: Const theFC1FieldName1 As String = "FieldName1"  'TODO: Insert real field name
    '    ' TEMPLATE: Const theFC1Name2 As String = "FieldName2"  'TODO: Insert real field name
    '    ' TEMPLATE: Dim theFC1FieldNames As New List(Of String)
    '    ' TEMPLATE: theFC1FieldNames.Add(theFC1FieldName1)
    '    ' TEMPLATE: colFC1FieldNames.Add(theFC1FieldName2)

    '    ' Set up to find the Taxlot feature class fields.
    '    Dim theFCName1 As String = EditorExtension.TableNamesSettings.TaxLotFC
    '    Dim theFC1FieldName1 As String = EditorExtension.TaxLotSettings.OrmapTaxlotField
    '    Dim theFC1FieldName2 As String = EditorExtension.TaxLotSettings.OrmapMapNumberField
    '    Dim theFC1FieldName3 As String = EditorExtension.TaxLotSettings.MapTaxlotField
    '    Dim theFC1FieldName4 As String = EditorExtension.TaxLotSettings.TaxlotField
    '    Dim theFC1FieldName5 As String = EditorExtension.TaxLotSettings.AnomalyField
    '    Dim theFC1FieldNamesList As New List(Of String)
    '    theFC1FieldNamesList.Add(theFC1FieldName1)
    '    theFC1FieldNamesList.Add(theFC1FieldName2)
    '    theFC1FieldNamesList.Add(theFC1FieldName3)
    '    theFC1FieldNamesList.Add(theFC1FieldName4)
    '    theFC1FieldNamesList.Add(theFC1FieldName5)

    '    ' Set up to find the MapIndex feature class fields.
    '    Dim theFCName2 As String = EditorExtension.TableNamesSettings.MapIndexFC
    '    Dim theFC2FieldName1 As String = EditorExtension.MapIndexSettings.MapNumberField
    '    Dim theFC2FieldName2 As String = EditorExtension.MapIndexSettings.MapScaleField  ' TODO: [NIS] Does the tool need this field?
    '    Dim theFC2FieldNamesList As New List(Of String)
    '    theFC2FieldNamesList.Add(theFC2FieldName1)
    '    theFC2FieldNamesList.Add(theFC2FieldName2)

    '    Dim foundAllFields As Boolean = True 'initial assumption
    '    Const canLoadData As Boolean = True

    '    'TEMPLATE: foundAllFields = foundAllFields AndAlso FeatureClassHasRequiredFields(fcName1, colFieldNames1, loadData)
    '    'TEMPLATE: foundAllFields = foundAllFields AndAlso FeatureClassHasRequiredFields(fcName2, colFieldNames2, loadData)

    '    foundAllFields = foundAllFields AndAlso FeatureClassHasRequiredFields(theFCName1, theFC1FieldNamesList, canLoadData)
    '    foundAllFields = foundAllFields AndAlso FeatureClassHasRequiredFields(theFCName2, theFC2FieldNamesList, canLoadData)

    '    Return foundAllFields

    'End Function


    Private Sub DoToolOperation(ByVal Button As ESRIMouseButtons, ByVal X As Integer, ByVal Y As Integer)

        Dim withinOperation As Boolean = False

        Try
            If (Button <> ESRIMouseButtons.Left) Then
                ' Exit silently.
                Exit Try
            End If

            ' Check for valid data
            CheckValidDataProperties()
            If Not HasValidTaxlotData Then
                MessageBox.Show("Unable to assign taxlot values to polygons." & vbNewLine & _
                                "Valid Taxlot layer not found in the map.", _
                                "Taxlot Assignment", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If
            If Not HasValidMapIndexData Then
                MessageBox.Show("Unable to assign taxlot values to polygons." & vbNewLine & _
                                "Valid MapIndex layer not found in the map.", _
                                "Taxlot Assignment", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If
            If HasValidTaxlotData OrElse HasValidMapIndexData Then
                Exit Try
            End If
            
            ' If taxlot numbering is selected, then make sure value is numeric.
            Dim isTaxlotType As Boolean = (StrComp(Me.TaxlotType, TaxlotAssignment.taxlotNumberTypeTaxlot, CompareMethod.Text) = 0)
            If isTaxlotType Then
                If Not IsNumeric(Me.NumberStartingFrom) Then
                    Throw New InvalidOperationException(String.Format("Expected a number for {0}, got {1}.", "me.NumberStartingFrom", Me.NumberStartingFrom)) ' TODO: [NIS] Find a better exception.
                End If
            End If

            ' Create a search shape out of the point that the user clicked.
            Dim thePoint As IPoint = Nothing
            Dim theGeometry As IGeometry = Nothing

            thePoint = EditorExtension.Editor.Display.DisplayTransformation.ToMapPoint(X, Y)
            theGeometry = thePoint 'QI

            ' Insure the presence of an underlying MapIndex polygon.
            Dim theSpatialFilter As ISpatialFilter
            Dim theShapeFieldName As String
            Dim theMIFCursor As IFeatureCursor
            Dim theMIFeature As IFeature

            theSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
            theSpatialFilter.Geometry = theGeometry
            theShapeFieldName = DataMonitor.MapIndexFeatureLayer.FeatureClass.ShapeFieldName
            theSpatialFilter.OutputSpatialReference(theShapeFieldName) = EditorExtension.Editor.Map.SpatialReference
            theSpatialFilter.GeometryField = DataMonitor.MapIndexFeatureLayer.FeatureClass.ShapeFieldName
            theSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
            theMIFCursor = DataMonitor.MapIndexFeatureLayer.FeatureClass.Search(theSpatialFilter, False)
            theMIFeature = theMIFCursor.NextFeature
            If theMIFeature Is Nothing Then
                MessageBox.Show("Unable to assign taxlot values to polygons" & vbNewLine & _
                                "that are not within a MapIndex polygon.", _
                                "Taxlot Assignment", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If

            ' Verify the uniqueness of the specified taxlot number (if taxlot type input).
            If isTaxlotType Then
                '[Taxlot value is a number...]
                If Not IsTaxlotNumberLocallyUnique(CStr(Me.NumberStartingFrom), theGeometry) Then  ' TODO: [NIS] Confirm this function with Jim.
                    If MessageBox.Show("The current Taxlot value (" & Me.NumberStartingFrom & ")" & vbNewLine & _
                                       "is not unique within this MapIndex." & vbNewLine & _
                                       "Attribute feature with value anyway?", _
                                       "Taxlot Assignment", MessageBoxButtons.YesNo, _
                                       MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Try
                    End If
                End If
            End If

            '=====================================
            ' The Update Operation Starts Here...
            '=====================================

            '------------------------------------------
            ' Get the taxlot feature to update.
            ' If found, start the feature update operation.
            '------------------------------------------
            Dim theTaxlotFCursor As IFeatureCursor = Nothing
            Dim theTaxlotFeature As IFeature = Nothing
            ' Select any feature under the given point in the target layer
            theSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
            theSpatialFilter.Geometry = theGeometry
            theShapeFieldName = DataMonitor.TaxlotFeatureLayer.FeatureClass.ShapeFieldName
            theSpatialFilter.OutputSpatialReference(theShapeFieldName) = EditorExtension.Editor.Map.SpatialReference
            theSpatialFilter.GeometryField = DataMonitor.TaxlotFeatureLayer.FeatureClass.ShapeFieldName
            theSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
            theTaxlotFCursor = DataMonitor.TaxlotFeatureLayer.FeatureClass.Search(theSpatialFilter, False)
            If theTaxlotFCursor IsNot Nothing Then
                theTaxlotFeature = theTaxlotFCursor.NextFeature
                ' Start the feature update operation
                If theTaxlotFeature IsNot Nothing Then
                    '[At least one taxlot feature is selected...]
                    EditorExtension.Editor.StartOperation()
                    withinOperation = True
                Else
                    '[No taxlot features are selected...]
                    MessageBox.Show("No Taxlot features have been selected.", _
                                    "Taxlot Assignment", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    Exit Try
                End If
            End If

            '------------------------------------------
            ' Define the new Taxlot string value.
            '------------------------------------------
            Dim theExistingTaxlotVal As String = String.Empty 'initialize
            Dim theTLTaxlotFldIdx As Integer = DataMonitor.TaxlotFeatureLayer.FeatureClass.FindField(EditorExtension.TaxLotSettings.TaxlotField)
            theExistingTaxlotVal = CStr(IIf(IsDBNull(theTaxlotFeature.Value(theTLTaxlotFldIdx)), "", theTaxlotFeature.Value(theTLTaxlotFldIdx)))

            ' Check with user before updating Taxlot field
            If Len(theExistingTaxlotVal) > 0 And theExistingTaxlotVal <> "0" Then
                If MessageBox.Show("Taxlot currently has a Taxlot value (" & theExistingTaxlotVal & ")." & vbNewLine & _
                          "Update it?", "Taxlot Assignment", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Try
                End If
            End If

            Dim theNewTLTaxlotNumVal As String = String.Empty 'initialize
            If isTaxlotType Then
                '[Taxlot value is a number...]
                theNewTLTaxlotNumVal = CStr(Me.NumberStartingFrom) 'User entered number
                ' Remove leading Zeros for taxlot number if any exist (CInt conversion will remove them)
                theNewTLTaxlotNumVal = CStr(CInt(theNewTLTaxlotNumVal))
            Else
                '[Taxlot value is a word...]
                theNewTLTaxlotNumVal = Me.TaxlotType 'Predefined text enum
            End If

            '------------------------------------------
            ' End the edit operation (store & stop)
            '------------------------------------------
            theTaxlotFeature.Store()
            EditorExtension.Editor.StopOperation("Assign Taxlot Number (AutoIncrement)")
            withinOperation = False

            '------------------------------------------
            ' AutoIncrement if taxlot number type
            '------------------------------------------
            If isTaxlotType AndAlso Me.IncrementNumber > 0 Then
                Me.NumberStartingFrom += Me.IncrementNumber
            End If

            ' Select the feature
            EditorExtension.Editor.Map.ClearSelection()
            EditorExtension.Editor.Map.SelectFeature(DataMonitor.TaxlotFeatureLayer, theTaxlotFeature)

            ' Check for stacked features in the same location
            theTaxlotFeature = theTaxlotFCursor.NextFeature
            If Not theTaxlotFeature Is Nothing Then
                MessageBox.Show("Multiple (""vertical"") features found at this location. This tool can only edit one.", _
                                "Taxlot Assignment", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

            ' Update the view
            Dim theActiveView As IActiveView
            theActiveView = DirectCast(EditorExtension.Editor.Map, IActiveView)

            ' Partially refresh the display
            theActiveView.PartialRefresh(esriViewDrawPhase.esriViewBackground Or _
                                         esriViewDrawPhase.esriViewGeography Or _
                                         esriViewDrawPhase.esriViewGraphics Or _
                                         esriViewDrawPhase.esriViewGraphicSelection, Nothing, Nothing)

        Catch ex As Exception
            If withinOperation Then
                ' Abort any ongoing edit operations
                EditorExtension.Editor.AbortOperation()
                withinOperation = False
            End If

            MessageBox.Show(ex.ToString)

        Finally
            ' Insure that this tool keeps the focus
            Dim theUID As New UID
            Dim theCmdItem As ICommandItem
            theUID.Value = Me.Name
            theCmdItem = EditorExtension.Application.Document.CommandBars.Find(theUID, True, False)
            EditorExtension.Application.CurrentTool = theCmdItem

        End Try
    End Sub

    'Private Sub initializeData()

    '    ' Obtain references to feature layer feature classes
    '    _theTaxlotFLayer = GetTaxlotFeatureLayer()
    '    _theTaxlotFClass = _theTaxlotFLayer.FeatureClass
    '    _theMapIndexFLayer = GetMapIndexFeatureLayer()
    '    _theMapIndexFClass = _theMapIndexFLayer.FeatureClass

    '    ' Get field indexes
    '    With EditorExtension.TaxLotSettings

    '        ' Find the ORMAP Taxlot field index
    '        _theTLOrmapTaxlotNumberFldIdx = _theTaxlotFClass.FindField(.OrmapTaxlotField)
    '        If _theTLOrmapTaxlotNumberFldIdx = FieldNotFoundIndex Then
    '            ' TODO: [NIS] Raise exception?
    '        End If

    '        ' Find the ORMAP Map Number field index
    '        _theTLOrmapMapNumberFldIdx = _theTaxlotFClass.FindField(.OrmapMapNumberField)
    '        If _theTLOrmapMapNumberFldIdx = FieldNotFoundIndex Then
    '            ' TODO: [NIS] Raise exception?
    '        End If

    '        ' Find the Map Taxlot field index
    '        _theTLMapTaxlotFldIdx = _theTaxlotFClass.FindField(.MapTaxlotField)
    '        If _theTLMapTaxlotFldIdx = FieldNotFoundIndex Then
    '            ' TODO: [NIS] Raise exception?
    '        End If

    '        ' Find the Taxlot field index
    '        _theTLTaxlotFldIdx = _theTaxlotFClass.FindField(.TaxlotField)
    '        If _theTLTaxlotFldIdx = FieldNotFoundIndex Then
    '            ' TODO: [NIS] Raise exception?
    '        End If

    '        ' Find the Anomaly field index
    '        _theTLAnomalyFldIdx = _theTaxlotFClass.FindField(.AnomalyField)
    '        If _theTLAnomalyFldIdx = FieldNotFoundIndex Then
    '            ' TODO: [NIS] Raise exception?
    '        End If

    '    End With

    'End Sub

#End Region

    '#Region "State Machine"
    '
    '    ' TODO: [NIS} Embed URL reference to statechart in the XML help for these methods.
    '
    '    Private Sub TransitionE1()
    '        StateS1_2(StatePassageType.Exiting)
    '        CondState1()
    '    End Sub
    '
    '    Private Sub TransitionE2()
    '        StateS1_1(StatePassageType.Exiting)
    '        StateS1_2(StatePassageType.Entering)
    '    End Sub
    '
    '    Private Sub TransitionE3()
    '        StateS1(StatePassageType.Exiting)
    '        CondState1()
    '    End Sub
    '
    '    Private Sub StateS1(ByVal statePassage As StatePassageType)
    '        Select Case statePassage
    '            Case StatePassageType.Entering
    '                ' Do actions
    '                ' (none)
    '                ' Do substate transitions
    '                StateS1_2(StatePassageType.Entering)
    '            Case StatePassageType.Exiting
    '                ' Do actions
    '                ' (none)
    '                ' Do substate transitions
    '                StateS1_1(StatePassageType.Exiting)
    '                StateS1_2(StatePassageType.Exiting)
    '        End Select
    '    End Sub
    '
    '    Private Sub StateS1_1(ByVal statePassage As StatePassageType)
    '        '[Tool Enabled...]
    '        Select Case statePassage
    '            Case StatePassageType.Entering
    '                setState(CommandStateType.Enabled)
    '                ' Do actions
    '                ' (none)
    '                ' Do substate transitions
    '                ' (none)
    '            Case StatePassageType.Exiting
    '                ' Do actions
    '                ' (none)
    '                ' Do substate transitions
    '                ' (none)
    '        End Select
    '    End Sub
    '
    '    Private Sub StateS1_2(ByVal statePassage As StatePassageType)
    '        '[Tool Disabled...]
    '        Select Case statePassage
    '            Case StatePassageType.Entering
    '                setState(CommandStateType.Disabled)
    '                ' Do actions
    '                ' (none)
    '                ' Do substate transitions
    '                ' (none)
    '            Case StatePassageType.Exiting
    '                ' Do actions
    '                ' (none)
    '                ' Do substate transitions
    '                ' (none)
    '        End Select
    '    End Sub
    '
    '    Private Sub CondState1()
    '        ' Evaluate condition
    '        If hasRequiredData() Then
    '            StateS1_1(StatePassageType.Entering)
    '        Else
    '            StateS1_2(StatePassageType.Entering)
    '        End If
    '    End Sub
    '
    '#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Dim canEnable As Boolean
            canEnable = EditorExtension.CanEnableExtendedEditing
            Return canEnable
        End Get
    End Property

#End Region

#Region "Methods"

    Public Overrides Sub OnClick()
        Try
            ' Show and activate the partner form.
            If PartnerTaxlotAssignmentForm.Visible Then
                PartnerTaxlotAssignmentForm.Activate()
            Else
                PartnerTaxlotAssignmentForm.Show()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Occurs when this command is created.
    ''' </summary>
    ''' <param name="hook">A generic <c>Object</c> hook to an instance of the application.</param>
    ''' <remarks>The application hook may not point to an <c>IMxApplication</c> object.</remarks>
    Public Overrides Sub OnCreate(ByVal hook As Object)
        Try
            If Not hook Is Nothing Then
                _application = DirectCast(hook, IApplication)

                'Disable tool if parent application is not ArcMap
                If TypeOf hook Is IMxApplication Then
                    MyBase.m_enabled = True
                Else
                    MyBase.m_enabled = False
                End If

                If MyBase.m_enabled Then
                    ' Set partner form.
                    setPartnerTaxlotAssignmentForm(New TaxlotAssignmentForm())
                End If
            End If

            ' NOTE: Add other initialization code here...

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' This method is called when a mouse button is pressed down, when this tool is active. 
    ''' </summary>
    ''' <param name="Button">Specifies which mouse button is pressed; 1 for the left mouse button, 2 for the right mouse button, and 4 for the middle mouse button.</param>
    ''' <param name="Shift">Specifies an integer corresponding to the state of the SHIFT (bit 0), CTRL (bit 1) and ALT (bit 2) keys. When none, some, or all of these keys are pressed none, some, or all the bits get set. These bits correspond to the values 1, 2, and 4, respectively. For example, if both SHIFT and ALT were pressed, Shift would be 5.</param>
    ''' <param name="X">The X coordinate, in device units, of the location of the mouse event. See the OnMouseDown Event for more details.</param>
    ''' <param name="Y">The Y coordinate, in device units, of the location of the mouse event. See the OnMouseDown Event for more details.</param>
    ''' <remarks></remarks>
    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        Try
            If Button = ESRIMouseButtons.Left Then
                '[Left button clicked...]
                DoToolOperation(DirectCast(Button, ESRIMouseButtons), X, Y)
            ElseIf Button = ESRIMouseButtons.Right Then
                '[Right button clicked...]
                ' Show and activate the partner form.
                If PartnerTaxlotAssignmentForm.Visible Then
                    PartnerTaxlotAssignmentForm.Activate()
                Else
                    PartnerTaxlotAssignmentForm.Show()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

#End Region

#End Region

#Region "Implemented Interface Members (none)"
#End Region

#Region "Other Members"

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "d091f7ea-0626-4d05-9a7c-533e0961f1cc"
    Public Const InterfaceId As String = "938bb5ab-a827-4731-be72-9db650fb8ed3"
    Public Const EventsId As String = "04e3b3ef-1fb3-46ed-b8d7-11dc19763f32"
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
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub

    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

#End Region

End Class



