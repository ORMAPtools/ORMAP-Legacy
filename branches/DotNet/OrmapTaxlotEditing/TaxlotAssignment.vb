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

#Region "Subversion Keyword expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision$
'Date of Last Change: $Date$
#End Region

Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.StringUtilities
Imports OrmapTaxlotEditing.Utilities

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
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

        Try
            ' Set the (enabled) cursor based on the name of the class.
            Dim cursorResourceName As String = Me.GetType().Name + ".cur"
            MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), cursorResourceName)
        Catch ex As ArgumentException
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Cursor")
        End Try

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication

    Private _disabledCursor As System.Drawing.Image

    ' Initialize document and map objects, and their events for tool reference only
    Private _doc As IDocument
    Private _focusMap As IMap

    ' Get to feature layers and feature classes
    Private _theTaxlotFLayer As IFeatureLayer
    Private _theTaxlotFClass As IFeatureClass
    Private _theMapIndexFLayer As IFeatureLayer
    Private _theMapIndexFClass As IFeatureClass

    ' Field indexes
    Private _theTLOrmapTaxlotNumberFldIdx As Integer
    Private _theTLOrmapMapNumberFldIdx As Integer
    Private _theTLMapTaxlotFldIdx As Integer
    Private _theTLTaxlotFldIdx As Integer
    Private _theTLAnomalyFldIdx As Integer

#End Region

#Region "Properties"

    Private _state As CommandStateType = CommandStateType.Disabled

    Public ReadOnly Property State() As CommandStateType
        Get
            Return _state
        End Get
    End Property

    Private Sub setState(ByVal stateType As CommandStateType)
        _state = stateType
    End Sub

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
        filePath = My.Application.Info.DirectoryPath & "\help\TaxlotAssignmentHelp.rtf"
        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(filePath) Then
            ' Open help file from the application directory.
            Dim helpForm As New HelpForm
            helpForm.uxContent.LoadFile(filePath, RichTextBoxStreamType.RichText)
            helpForm.Text = "Taxlot Assignment Help"
            helpForm.Show()
        Else
            MessageBox.Show("No help file available in the directory " & My.Application.Info.DirectoryPath & "\help" & ".")
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

#Region "EditEvents Event Handlers"

    Private Sub EditEvents_OnStartEditing() 'Implements ESRI.ArcGIS.Editor.IEditEvents.OnStartEditing
        ' State Transistion E1
        TransitionE1()
    End Sub

    Private Sub EditEvents_OnStopEditing(ByVal save As Boolean) 'Implements ESRI.ArcGIS.Editor.IEditEvents.OnStopEditing
        ' State Transistion E2
        TransitionE2()
    End Sub

#End Region

#Region "ActiveViewEvents Event Handlers"

    Public Sub ActiveViewEvents_FocusMapChanged() 'Implements ESRI.ArcGIS.Carto.IActiveViewEvents.FocusMapChanged
        ' State Transistion E3
        TransitionE3()
    End Sub

    Public Sub ActiveViewEvents_ItemAdded(ByVal Item As Object) 'Implements ESRI.ArcGIS.Carto.IActiveViewEvents.ItemAdded
        ' State Transistion E3
        TransitionE3()
    End Sub

    Public Sub ActiveViewEvents_ItemDeleted(ByVal Item As Object) 'Implements ESRI.ArcGIS.Carto.IActiveViewEvents.ItemDeleted
        ' State Transistion E3
        TransitionE3()
    End Sub

#End Region

#End Region

#Region "Methods"

    Public Function HasRequiredData() As Boolean

        ' TEMPLATE: Const fcName1 As String = "FeatureClassName1"  'TODO: Insert real fc name
        ' TEMPLATE: Const fieldName1FC1 As String = "FieldName1"  'TODO: Insert real field name
        ' TEMPLATE: Const fieldName2FC1 As String = "FieldName2"  'TODO: Insert real field name
        ' TEMPLATE: Dim colFieldNames1 As New Collection
        ' TEMPLATE: colFieldNames1.Add(fieldName1FC1)
        ' TEMPLATE: colFieldNames1.Add(fieldName2FC1)

        ' Set up to find the Taxlot feature class fields.
        Dim fcName1 As String = EditorExtension.TableNamesSettings.TaxLotFC
        Dim fieldName1FC1 As String = EditorExtension.TaxLotSettings.OrmapTaxlotField
        Dim fieldName2FC1 As String = EditorExtension.TaxLotSettings.OrmapMapNumberField
        Dim fieldName3FC1 As String = EditorExtension.TaxLotSettings.MapTaxlotField
        Dim fieldName4FC1 As String = EditorExtension.TaxLotSettings.TaxlotField
        Dim fieldName5FC1 As String = EditorExtension.TaxLotSettings.AnomalyField
        Dim colFieldNames1 As New Collection
        colFieldNames1.Add(fieldName1FC1)
        colFieldNames1.Add(fieldName2FC1)
        colFieldNames1.Add(fieldName3FC1)
        colFieldNames1.Add(fieldName4FC1)
        colFieldNames1.Add(fieldName5FC1)

        ' Set up to find the MapIndex feature class fields.
        Dim fcName2 As String = EditorExtension.TableNamesSettings.MapIndexFC
        Dim fieldName1FC2 As String = EditorExtension.MapIndexSettings.MapNumberField
        Dim fieldName2FC2 As String = EditorExtension.MapIndexSettings.MapScaleField  ' TODO: Does the tool need this field?
        Dim colFieldNames2 As New Collection
        colFieldNames2.Add(fieldName1FC2)
        colFieldNames2.Add(fieldName2FC2)

        Dim foundAllFields As Boolean = True 'initial assumption
        Const loadData As Boolean = True

        'TEMPLATE: foundAllFields = foundAllFields AndAlso hasRequiredFields(fcName1, colFieldNames1, loadData)
        'TEMPLATE: foundAllFields = foundAllFields AndAlso hasRequiredFields(fcName2, colFieldNames2, loadData)

        foundAllFields = foundAllFields AndAlso hasRequiredFields(fcName1, colFieldNames1, loadData)
        foundAllFields = foundAllFields AndAlso hasRequiredFields(fcName2, colFieldNames2, loadData)

        Return foundAllFields

    End Function


    Private Sub DoToolOperation(ByVal Button As Integer, ByVal X As Integer, ByVal Y As Integer)

        Try
            ' TODO: [NIS] Define button parameters with enum
            If (Button <> 1) Then
                ' Exit silently.
                Exit Try
            End If

            Dim isTaxlotType As Boolean = (StrComp(Me.TaxlotType, TaxlotAssignment.taxlotNumberTypeTaxlot, CompareMethod.Text) = 0)

            'If taxlot numbering is selected, then make sure value is numeric
            If isTaxlotType Then
                If Not IsNumeric(Me.NumberStartingFrom) Then
                    Throw New InvalidOperationException(String.Format("Expected a number for {0}, got {1}.", "me.NumberStartingFrom", Me.NumberStartingFrom)) ' TODO: [NIS] Find a better exception.
                End If
            End If

            ' Create a search shape out of the point that the user clicked
            Dim thePoint As IPoint = Nothing
            Dim theGeometry As IGeometry = Nothing

            thePoint = EditorExtension.Editor.Display.DisplayTransformation.ToMapPoint(X, Y)
            'TODO: NIS Get rid of this commented line?
            'theGeometry = EditorExtension.Editor.CreateSearchShape(thePoint) 'Returns an IEnvelope
            theGeometry = thePoint 'QI

            ' Initialize the feature class and field data
            initializeData()

            ' Insure the validity of the underlying map index polygon
            Dim theSpatialFilter As ISpatialFilter
            Dim theShapeFieldName As String
            Dim theMIFCursor As IFeatureCursor
            Dim theMIFeature As IFeature

            theSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
            theSpatialFilter.Geometry = theGeometry
            theShapeFieldName = _theMapIndexFClass.ShapeFieldName
            theSpatialFilter.OutputSpatialReference(theShapeFieldName) = _focusMap.SpatialReference
            theSpatialFilter.GeometryField = _theMapIndexFClass.ShapeFieldName
            theSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
            theMIFCursor = _theMapIndexFClass.Search(theSpatialFilter, False)
            theMIFeature = theMIFCursor.NextFeature
            If theMIFeature Is Nothing Then
                MessageBox.Show("Unable to assign taxlot values to polygons" & vbNewLine & _
                                "that are not within a Map Index polygon.", _
                                Me.Name, MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If

            ' Verify the validity and uniqueness of the specified taxlot type taxlot number
            If isTaxlotType Then
                '[Taxlot value is a number...]
                If Not ValidateTaxlotNumber(CStr(Me.NumberStartingFrom), theGeometry) Then  ' TODO: [NIS] Confirm this function with Jim.
                    If MessageBox.Show("The current Taxlot value (" & Me.NumberStartingFrom & ")" & vbNewLine & _
                                       "is not unique within this MapIndex." & vbNewLine & _
                                       "Attribute feature with value anyway?", _
                                       Me.Name, MessageBoxButtons.YesNo, _
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
            theShapeFieldName = _theTaxlotFClass.ShapeFieldName
            theSpatialFilter.OutputSpatialReference(theShapeFieldName) = _focusMap.SpatialReference
            theSpatialFilter.GeometryField = _theTaxlotFClass.ShapeFieldName
            theSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
            theTaxlotFCursor = _theTaxlotFClass.Search(theSpatialFilter, False)
            If theTaxlotFCursor IsNot Nothing Then
                theTaxlotFeature = theTaxlotFCursor.NextFeature
                ' Start the feature update operation
                If theTaxlotFeature IsNot Nothing Then
                    '[At least one taxlot feature is selected...]
                    EditorExtension.Editor.StartOperation()
                Else
                    '[No taxlot features are selected...]
                    MessageBox.Show("No taxlot features have been selected.", _
                                    Me.Name, MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    Exit Try
                End If
            End If

            '------------------------------------------
            ' Get the OrmapMapNumber as a string value
            '------------------------------------------
            Dim theExistOrmapMapNumberVal As String = String.Empty 'initialize
            ' Get the current OrmapMapNumber.
            ' TODO: [NIS] Resolve - UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            theExistOrmapMapNumberVal = CStr(IIf(IsDBNull(theTaxlotFeature.Value(_theTLOrmapMapNumberFldIdx)), "", theTaxlotFeature.Value(_theTLOrmapMapNumberFldIdx)))

            ' Obtain the OrmapMapNumber from a MapIndex polygon if it is not present.
            If Len(theExistOrmapMapNumberVal) = 0 Then
                ' TODO: [NIS] Confirm - "This call will point _theMapIndexFLayer to the MapIndex feature class"
                CalculateTaxlotValues(theTaxlotFeature, _theMapIndexFLayer)  ' TODO: [NIS] Confirm this function with Jim.
                ' Get the current ORMAP Number again after the calculate above.
                ' TODO: [NIS] Resolve - UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                theExistOrmapMapNumberVal = CStr(IIf(IsDBNull(theTaxlotFeature.Value(_theTLOrmapMapNumberFldIdx)), "", theTaxlotFeature.Value(_theTLOrmapMapNumberFldIdx)))

                ' Stop if there is still no current OrmapMapNumber.
                If Len(theExistOrmapMapNumberVal) = 0 Then
                    MessageBox.Show("OrmapMapNumber is empty for this taxlot or MapIndex." & vbNewLine & _
                                    "Use the MapIndex tool to populate the OrmapMapNumber field" & vbNewLine & _
                                    "before using this tool", Me.Name, MessageBoxButtons.OK)
                    Exit Try
                End If
            End If

            ' TODO: [NIS] Implement this (needs code elsewhere as well).
            ''------------------------------------------
            '' Define the MapNumber as a string value.
            ''------------------------------------------
            'Dim theExistMapNumberVal As String = String.Empty 'initialize
            '' TODO: [NIS] Resolve - UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            'theExistMapNumberVal = CStr(IIf(IsDBNull(theTaxlotFeature.Value(_theTLMapNumberFldIdx)), "", theTaxlotFeature.Value(_theTLTaxlotFldIdx)))

            '------------------------------------------
            ' Define the Taxlot number (can be a word 
            ' also) as a string value.
            '------------------------------------------
            Dim theExistTaxlotNumberVal As String = String.Empty 'initialize
            ' TODO: [NIS] Resolve - UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            theExistTaxlotNumberVal = CStr(IIf(IsDBNull(theTaxlotFeature.Value(_theTLTaxlotFldIdx)), "", theTaxlotFeature.Value(_theTLTaxlotFldIdx)))

            ' Optionally, update the taxlot number field
            If Len(theExistTaxlotNumberVal) > 0 And theExistTaxlotNumberVal <> "0" Then
                If MessageBox.Show("Taxlot currently has a Taxlot value (" & theExistTaxlotNumberVal & ")." & vbNewLine & _
                          "Update it?", Me.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Try
                End If
            End If
            ' Taxlot numbers can be less than 5-digits.
            ' The Taxlot value within values in the OrmapTaxlot field must be exactly 5 digits.
            ' Two versions of the Taxlot number will be used for these purposes.
            Dim theNewTLTaxlotNumVal As String = String.Empty 'initialize
            Dim theNewTLTaxlotNumVal_5digit As String = String.Empty 'initialize
            If isTaxlotType Then
                '[Taxlot value is a number...]
                theNewTLTaxlotNumVal = CStr(Me.NumberStartingFrom) 'User entered number
                theNewTLTaxlotNumVal_5digit = theNewTLTaxlotNumVal
                ' Remove leading Zeros for taxlot number if any exist (CInt conversion will remove them)
                theNewTLTaxlotNumVal = CStr(CInt(theNewTLTaxlotNumVal))
                ' Make sure 5-digit number is 5 characters by padding on the left with zeros
                If Len(theNewTLTaxlotNumVal_5digit) < ORMAPNumber.GetOrmap_TaxlotFieldLength Then  ' TODO: [NIS] Why NOT just use 5 here? Other code assumes this length so it should not vary.
                    Do Until Len(theNewTLTaxlotNumVal_5digit) = ORMAPNumber.GetOrmap_TaxlotFieldLength
                        theNewTLTaxlotNumVal_5digit = "0" & theNewTLTaxlotNumVal_5digit
                    Loop
                End If
            Else
                '[Taxlot value is a word...]
                theNewTLTaxlotNumVal = Me.TaxlotType 'Predefined text enum
                theNewTLTaxlotNumVal_5digit = Me.TaxlotType 'Predefined text enum
            End If

            '------------------------------------------
            ' Get the MapSuffixType and MapSuffixNum as 
            ' string values
            '------------------------------------------
            Dim theTLMapSuffixNumVal As String = String.Empty 'initialize
            Dim theTLMapSuffixTypeVal As String = String.Empty 'initialize
            theTLMapSuffixTypeVal = GetMapSuffixType(theTaxlotFeature)
            theTLMapSuffixNumVal = GetMapSuffixNum(theTaxlotFeature)
            ' If value is null, use the default
            ' TODO: [NIS] Resolve - UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            If IsDBNull(theTLMapSuffixTypeVal) Then theTLMapSuffixTypeVal = "0"
            ' TODO: [NIS] Resolve - UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            If IsDBNull(theTLMapSuffixNumVal) Then theTLMapSuffixNumVal = "000"

            '------------------------------------------
            ' Define the Anomaly from MapIndex
            '------------------------------------------
            Dim theORMAPNumberClass As New ORMAPNumber()
            Dim theAnomalyVal As String = String.Empty
            If (theORMAPNumberClass.ParseNumber(theExistOrmapMapNumberVal)) Then
                theAnomalyVal = theORMAPNumberClass.Anomaly
            End If

            '------------------------------------------
            ' Put together the OrmapTaxlot (ORTaxlot, 
            ' NOT to be confused with the Taxlot or 
            ' MapTaxlot fields) string from its parts.
            ' Set the value.
            '------------------------------------------
            Dim theCombinedOrmapTaxlotNumber As String = String.Empty 'initialize
            theCombinedOrmapTaxlotNumber = theExistOrmapMapNumberVal & theNewTLTaxlotNumVal_5digit
            ' TODO: [NIS] Find out if this is VB6 pattern is actually the correct pattern...
            'Dim theShortOrmapMapNumber As String = String.Empty 'initialize
            'theShortOrmapMapNumber = OrmapMapNumberNoCountyCodeSuffix(theExistOrmapMapNumberVal)
            'theCombinedOrmapTaxlotNumber = theShortOrmapMapNumber & theTLMapSuffixTypeVal & theTLMapSuffixNumVal & theNewTLTaxlotNumVal_5digit

            '------------------------------------------
            ' Define the MapTaxlot (NOT to be confused 
            ' with the OrmapTaxlot!) value.
            '------------------------------------------
            'Create  masked value from a combination of ORMapNum and the new taxlot
            'Note: Special code for Lane County (see comment below).
            Dim theMapTaxlotNumber As String = String.Empty 'initialize
            Dim theDefaultCountyCode As Integer
            theDefaultCountyCode = CInt(EditorExtension.DefaultValuesSettings.County)  ' TODO: [NIS] Confirm field choice
            Select Case theDefaultCountyCode
                Case 1 To 19, 21 To 36
                    theMapTaxlotNumber = CreateMapTaxlotValue(theExistOrmapMapNumberVal & theNewTLTaxlotNumVal_5digit, (EditorExtension.TaxLotSettings.MapTaxlotFormatMask))
                Case 20
                    ' 1.  Lane County uses a 2-digit numeric identifier for ranges.
                    '     Special handling is required for east ranges, where 02E is
                    '     stored as 25, 03E as 35, etc.
                    ' 2.  ORMAP standards (OCDES (pg 13); Taxmap Data Model (pg 11)) assert that
                    '     this field should be equal to MAPNUMBER + TAXLOT. In this case, MAPNUMBER
                    '     is already in the right format, thus removing the need for the
                    '     gfn_s_CreateMapTaxlotValue function. Also, in this case, TAXLOT is padded
                    '     on the left with zeros to make it always a 5-digit number (see comment
                    '     above).
                    theMapTaxlotNumber = Trim(Left(theExistOrmapMapNumberVal, 8)) & theNewTLTaxlotNumVal_5digit
                    ' TODO: [NIS] Implement this instead of the above line.
                    'theMapTaxlotNumber = Trim(Left(theExistMapNumberVal, 8)) & theNewTLTaxlotNumVal_5digit
            End Select

            '##################################
            ' Write the Taxlot value
            theTaxlotFeature.Value(_theTLTaxlotFldIdx) = theNewTLTaxlotNumVal
            ' Write the OrmapTaxlot value
            theTaxlotFeature.Value(_theTLOrmapTaxlotNumberFldIdx) = theCombinedOrmapTaxlotNumber
            ' Write the MapTaxlot value
            theTaxlotFeature.Value(_theTLMapTaxlotFldIdx) = theMapTaxlotNumber
            ' Write the Anomaly value
            theTaxlotFeature.Value(_theTLAnomalyFldIdx) = theAnomalyVal
            '##################################

            '------------------------------------------
            ' End the edit operation
            '------------------------------------------
            theTaxlotFeature.Store()
            EditorExtension.Editor.StopOperation("Assign Taxlot Number (AutoIncrement)")

            '------------------------------------------
            ' AutoIncrement if taxlot number type
            '------------------------------------------
            If isTaxlotType AndAlso Me.IncrementNumber > 0 Then
                Me.NumberStartingFrom += Me.IncrementNumber
            End If

            ' Select the feature
            _focusMap.ClearSelection()
            _focusMap.SelectFeature(_theTaxlotFLayer, theTaxlotFeature)

            ' Check for stacked features in the same location
            theTaxlotFeature = theTaxlotFCursor.NextFeature
            If Not theTaxlotFeature Is Nothing Then
                MessageBox.Show("Multiple (""vertical"") features found at this location. This tool can only edit one.", _
                                Me.Name, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                ' TODO: [NIS] Enhance to handle more than one vertical feature.
            End If

            ' Update the view
            Dim theActiveView As IActiveView
            theActiveView = DirectCast(_focusMap, IActiveView)

            ' Partially refresh the display
            theActiveView.PartialRefresh(esriViewDrawPhase.esriViewBackground Or _
                                         esriViewDrawPhase.esriViewGeography Or _
                                         esriViewDrawPhase.esriViewGraphics Or _
                                         esriViewDrawPhase.esriViewGraphicSelection, Nothing, Nothing)

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally

            ' Abort any ongoing edit operations
            EditorExtension.Editor.AbortOperation()

            ' Insure that this tool keeps the focus
            Dim theUID As New UID
            Dim theCmdItem As ICommandItem
            theUID.Value = Me.Name
            theCmdItem = _doc.CommandBars.Find(theUID, True, False)
            _application.CurrentTool = theCmdItem

        End Try
    End Sub

    <ObsoleteAttribute("Use the hasRequiredFields() function instead.", True)> _
    Private Function hasRequiredFeatureLayers(ByVal featureClassNames As Collection, ByVal loadData As Boolean) As Boolean

        Dim foundAllLayers As Boolean = True  'initial assumption

        For Each fcn As String In featureClassNames
            ' Confirm layer is present in current map
            Dim theFLayer As IFeatureLayer
            theFLayer = FindFeatureLayerByDSName(fcn)
            If theFLayer IsNot Nothing Then
                foundAllLayers = True
            Else
                If loadData Then
                    '[Load option accepted...]
                    ' Attempt to load and find the taxlot layer in the map document
                    If LoadFCIntoMap(fcn) Then
                        '[Layer loaded...]
                        foundAllLayers = True
                    Else
                        '[Unable to load the layer...]
                        foundAllLayers = False
                        Exit For
                    End If
                Else
                    '[Data not present and load option refused...]
                    foundAllLayers = False
                    Exit For
                End If
            End If
        Next fcn

        Return foundAllLayers

    End Function

    Private Function hasRequiredFields(ByVal featureClassName As String, ByVal fieldNames As Collection, ByVal loadData As Boolean) As Boolean

        Dim returnValue As Boolean = True 'initial assumption

        ' Confirm data layer is present in current map
        Dim theFLayer As IFeatureLayer

        theFLayer = FindFeatureLayerByDSName(featureClassName)

        If theFLayer Is Nothing Then
            If loadData Then
                '[Load option accepted...]
                ' Attempt to load and find the taxlot layer in the map document
                If LoadFCIntoMap(featureClassName) Then
                    '[Layer loaded...]
                    ' Confirm fields are present
                    Dim foundAllFields As Boolean = True  'initial assumption
                    Dim fieldIndex As Integer
                    For Each fn As String In fieldNames
                        fieldIndex = theFLayer.FeatureClass.FindField(fn)
                        If fieldIndex <> FieldNotFoundIndex Then
                            foundAllFields = True
                        Else
                            foundAllFields = False
                            Exit For
                        End If
                    Next fn
                    returnValue = foundAllFields
                Else
                    '[Unable to load the layer...]
                    returnValue = False
                End If
            Else
                '[Data not present and load option refused...]
                returnValue = False
            End If
        End If

        Return returnValue

    End Function

    ' TODO: [NIS] Replace with local property
    Private Function getTaxlotFeatureLayer() As IFeatureLayer
        ' Find Taxlot feature layer
        Dim theTaxlotFLayer As IFeatureLayer
        With EditorExtension.TableNamesSettings
            ' Find Taxlot feature layer
            theTaxlotFLayer = FindFeatureLayerByDSName(.TaxLotFC)
            If theTaxlotFLayer Is Nothing Then
                ' TODO: [NIS] Raise an exception instead?
                Return Nothing
            End If
        End With
        Return theTaxlotFLayer
    End Function

    ' TODO: [NIS] Replace with local property
    Private Function getMapIndexFeatureLayer() As IFeatureLayer
        ' Find Map Index feature layer
        Dim theMapIndexFLayer As IFeatureLayer
        With EditorExtension.TableNamesSettings
            ' Find MapIndex feature layer
            theMapIndexFLayer = FindFeatureLayerByDSName(.MapIndexFC)
            If theMapIndexFLayer Is Nothing Then
                ' TODO: [NIS] Raise an exception instead?
                Return Nothing
            End If
        End With
        Return theMapIndexFLayer
    End Function

    Private Sub initializeData()

        ' Initialize document and map objects, and their events for tool reference only
        _doc = EditorExtension.Application.Document
        _focusMap = DirectCast(_doc, IMxDocument).FocusMap

        ' Obtain references to feature layer feature classes
        _theTaxlotFLayer = getTaxlotFeatureLayer()
        _theTaxlotFClass = _theTaxlotFLayer.FeatureClass
        _theMapIndexFLayer = getMapIndexFeatureLayer()
        _theMapIndexFClass = _theMapIndexFLayer.FeatureClass

        ' Get field indexes
        With EditorExtension.TaxLotSettings

            ' Find the ORMAP Taxlot field index
            _theTLOrmapTaxlotNumberFldIdx = _theTaxlotFClass.FindField(.OrmapTaxlotField)
            If _theTLOrmapTaxlotNumberFldIdx = FieldNotFoundIndex Then
                ' TODO: [NIS] Raise exception?
            End If

            ' Find the ORMAP Map Number field index
            _theTLOrmapMapNumberFldIdx = _theTaxlotFClass.FindField(.OrmapMapNumberField)
            If _theTLOrmapMapNumberFldIdx = FieldNotFoundIndex Then
                ' TODO: [NIS] Raise exception?
            End If

            ' Find the Map Taxlot field index
            _theTLMapTaxlotFldIdx = _theTaxlotFClass.FindField(.MapTaxlotField)
            If _theTLMapTaxlotFldIdx = FieldNotFoundIndex Then
                ' TODO: [NIS] Raise exception?
            End If

            ' Find the Taxlot field index
            _theTLTaxlotFldIdx = _theTaxlotFClass.FindField(.TaxlotField)
            If _theTLTaxlotFldIdx = FieldNotFoundIndex Then
                ' TODO: [NIS] Raise exception?
            End If

            ' Find the Anomaly field index
            _theTLAnomalyFldIdx = _theTaxlotFClass.FindField(.AnomalyField)
            If _theTLAnomalyFldIdx = FieldNotFoundIndex Then
                ' TODO: [NIS] Raise exception?
            End If

        End With

    End Sub

#End Region

#Region "State Machine"

    ' TODO: [NIS} Embed URL reference to statechart in the XML help for these methods.

    Private Sub TransitionE1()
        StateS1_2(StatePassageType.Exiting)
        CondState1()
    End Sub

    Private Sub TransitionE2()
        StateS1_1(StatePassageType.Exiting)
        StateS1_2(StatePassageType.Entering)
    End Sub

    Private Sub TransitionE3()
        StateS1(StatePassageType.Exiting)
        CondState1()
    End Sub

    Private Sub StateS1(ByVal statePassage As StatePassageType)
        Select Case statePassage
            Case StatePassageType.Entering
                ' Do actions
                ' (none)
                ' Do substate transitions
                StateS1_2(StatePassageType.Entering)
            Case StatePassageType.Exiting
                ' Do actions
                ' (none)
                ' Do substate transitions
                StateS1_1(StatePassageType.Exiting)
                StateS1_2(StatePassageType.Exiting)
        End Select
    End Sub

    Private Sub StateS1_1(ByVal statePassage As StatePassageType)
        '[Tool Enabled...]
        Select Case statePassage
            Case StatePassageType.Entering
                setState(CommandStateType.Enabled)
                ' Do actions
                ' (none)
                ' Do substate transitions
                ' (none)
            Case StatePassageType.Exiting
                ' Do actions
                ' (none)
                ' Do substate transitions
                ' (none)
        End Select
    End Sub

    Private Sub StateS1_2(ByVal statePassage As StatePassageType)
        '[Tool Disabled...]
        Select Case statePassage
            Case StatePassageType.Entering
                setState(CommandStateType.Disabled)
                ' Do actions
                ' (none)
                ' Do substate transitions
                ' (none)
            Case StatePassageType.Exiting
                ' Do actions
                ' (none)
                ' Do substate transitions
                ' (none)
        End Select
    End Sub

    Private Sub CondState1()
        ' Evaluate condition
        If HasRequiredData() Then
            StateS1_1(StatePassageType.Entering)
        Else
            StateS1_2(StatePassageType.Entering)
        End If
    End Sub

#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"

    Private _hasEventHandlers As Boolean = False

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Dim canEnable As Boolean
            canEnable = EditorExtension.CanEnableExtendedEditing
            If canEnable Then
                If Not _hasEventHandlers Then
                    ' Subscribe to edit events.
                    AddHandler EditorExtension.EditEvents.OnStartEditing, AddressOf EditEvents_OnStartEditing
                    AddHandler EditorExtension.EditEvents.OnStopEditing, AddressOf EditEvents_OnStopEditing
                    ' Subscribe to active view events.
                    AddHandler EditorExtension.ActiveViewEvents.FocusMapChanged, AddressOf ActiveViewEvents_FocusMapChanged
                    AddHandler EditorExtension.ActiveViewEvents.ItemAdded, AddressOf ActiveViewEvents_ItemAdded
                    AddHandler EditorExtension.ActiveViewEvents.ItemDeleted, AddressOf ActiveViewEvents_ItemDeleted
                    _hasEventHandlers = True
                End If
            Else
                If _hasEventHandlers Then
                    ' Unsubscribe to edit events.
                    RemoveHandler EditorExtension.EditEvents.OnStartEditing, AddressOf EditEvents_OnStartEditing
                    RemoveHandler EditorExtension.EditEvents.OnStopEditing, AddressOf EditEvents_OnStopEditing
                    ' Unsubscribe to active view events.
                    RemoveHandler EditorExtension.ActiveViewEvents.FocusMapChanged, AddressOf ActiveViewEvents_FocusMapChanged
                    RemoveHandler EditorExtension.ActiveViewEvents.ItemAdded, AddressOf ActiveViewEvents_ItemAdded
                    RemoveHandler EditorExtension.ActiveViewEvents.ItemDeleted, AddressOf ActiveViewEvents_ItemDeleted
                    _hasEventHandlers = False
                End If
            End If
            canEnable = canEnable AndAlso State = CommandStateType.Enabled
            Return canEnable
        End Get
    End Property

#End Region

#Region "Methods"

    Public Overrides Sub OnClick()
        Try
            ' TODO: [NIS] Remove this if not needed anymore.
            '' HACK: [NIS] Don't know why this form is disposed after first use, but 
            '' this check insures it is available again.
            'If PartnerTaxlotAssignmentForm.IsDisposed Then
            '    SetPartnerTaxlotAssignmentForm(New TaxlotAssignmentForm())
            'End If

            ' Show and activate the partner form.
            If PartnerTaxlotAssignmentForm.Visible Then
                PartnerTaxlotAssignmentForm.Activate()
            Else
                PartnerTaxlotAssignmentForm.Show()
            End If
        Catch ex As Exception
            ' TODO: [NIS] Add exception handling here
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

            ' TODO: Add other initialization code?

        Catch ex As Exception
            ' TODO: [NIS] Add exception handling here
        End Try
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        Try
            DoToolOperation(Button, X, Y)
        Catch ex As Exception
            ' TODO: [NIS] Add exception handling here
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



