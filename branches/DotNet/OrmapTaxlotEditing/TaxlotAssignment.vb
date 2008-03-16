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
    Implements IActiveViewEvents
    Implements IEditEvents

#Region "Class-Level Constants And Enumerations"

    ' Taxlot number type constants
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

#Region "Built-In Class Members (Properties, Methods, Events, Event Handlers, Delegates, Etc.)"

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
    Private _map As IMap

    ' Get to feature layers and feature classes
    Private _theTaxlotFLayer As IFeatureLayer
    Private _theTaxlotFClass As IFeatureClass
    Private _theMapIndexFLayer As IFeatureLayer
    Private _theMapIndexFClass As IFeatureClass

    ' Field indexes
    Private _theOrmapTaxlotNumberFld As Integer
    Private _theOrmapMapNumberFld As Integer
    Private _theOrmapMapTaxlotFld As Integer
    Private _theTLTaxlotFld As Integer
    Private _theTLAnomalyFld As Integer

#End Region

#Region "Properties"

    Private _state As CommandStateType

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

    Private WithEvents _partnerTaxlotAssignmentForm As TaxlotAssignmentForm  ' TODO: NIS Is WithEvents needed here?

    Friend ReadOnly Property PartnerTaxlotAssignmentForm() As TaxlotAssignmentForm
        Get
            If _partnerTaxlotAssignmentForm Is Nothing OrElse _partnerTaxlotAssignmentForm.IsDisposed Then
                SetPartnerTaxlotAssignmentForm(New TaxlotAssignmentForm())
            End If
            Return _partnerTaxlotAssignmentForm
        End Get
    End Property

    Private Sub SetPartnerTaxlotAssignmentForm(ByRef value As TaxlotAssignmentForm)
        If value IsNot Nothing Then
            _partnerTaxlotAssignmentForm = value
            ' Wire up partner form events.
            AddHandler _partnerTaxlotAssignmentForm.Load, AddressOf PartnerTaxlotAssignmentForm_Load
            AddHandler _partnerTaxlotAssignmentForm.uxHelp.Click, AddressOf uxHelp_Click
            AddHandler _partnerTaxlotAssignmentForm.uxType.SelectedValueChanged, AddressOf uxType_SelectedValueChanged
        End If
    End Sub

#End Region

#Region "Event Handlers"

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
        ' TODO: NIS Could be replaced with new help mechanism.
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

#Region "Methods"

    Public Function HasRequiredData() As Boolean

        Dim foundAllFields As Boolean = True 'initial assumption

        ' TEMPLATE: Const fcName1 As String = "FeatureClassName1"  'TODO: Insert real fc name
        ' TEMPLATE: Const fieldName1FC1 As String = "FieldName1"  'TODO: Insert real field name
        ' TEMPLATE: Const fieldName2FC1 As String = "FieldName2"  'TODO: Insert real field name
        ' TEMPLATE: Dim colFieldNames1 As New Collection
        ' TEMPLATE: colFieldNames1.Add(fieldName1FC1)
        ' TEMPLATE: colFieldNames1.Add(fieldName2FC1)

        Const fcName1 As String = "FeatureClassName1"  'TODO: NIS Insert real fc name
        Const fieldName1FC1 As String = "FieldName1"  'TODO: NIS Insert real field name
        Const fieldName2FC1 As String = "FieldName2"  'TODO: NIS Insert real field name
        Dim colFieldNames1 As New Collection
        colFieldNames1.Add(fieldName1FC1)
        colFieldNames1.Add(fieldName2FC1)

        Const fcName2 As String = "FeatureClassName2"  'TODO: NIS Insert real fc name
        Const fieldName1FC2 As String = "FieldName1"  'TODO: NIS Insert real field name
        Const fieldName2FC2 As String = "FieldName2"  'TODO: NIS Insert real field name
        Dim colFieldNames2 As New Collection
        colFieldNames1.Add(fieldName1FC2)
        colFieldNames1.Add(fieldName2FC2)

        Const loadData As Boolean = True
        If hasRequiredFields(fcName1, colFieldNames1, loadData) AndAlso _
           hasRequiredFields(fcName2, colFieldNames2, loadData) Then
            foundAllFields = True
        Else
            foundAllFields = False
        End If

        Return foundAllFields

    End Function

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

        Dim returnValue As Boolean = False 'initial assumption

        ' Confirm data layer is present in current map
        Dim theFLayer As IFeatureLayer

        theFLayer = FindFeatureLayerByDSName(featureClassName)

        If theFLayer IsNot Nothing Then
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
        Else
            '[Layer not found...]
            returnValue = False
        End If

        Return returnValue

    End Function

    ' TODO: NIS Replace with local property
    Private Function getTaxlotFeatureLayer() As IFeatureLayer
        ' Find Taxlot feature layer
        Dim theTaxlotFLayer As IFeatureLayer
        With EditorExtension.TableNamesSettings
            ' Find Taxlot feature layer
            theTaxlotFLayer = FindFeatureLayerByDSName(.TaxLotFC)
            If theTaxlotFLayer Is Nothing Then
                ' TODO: NIS Raise an exception instead?
                Return Nothing
            End If
        End With
        Return theTaxlotFLayer
    End Function

    ' TODO: NIS Replace with local property
    Private Function getMapIndexFeatureLayer() As IFeatureLayer
        ' Find Map Index feature layer
        Dim theMapIndexFLayer As IFeatureLayer
        With EditorExtension.TableNamesSettings
            ' Find MapIndex feature layer
            theMapIndexFLayer = FindFeatureLayerByDSName(.MapIndexFC)
            If theMapIndexFLayer Is Nothing Then
                ' TODO: NIS Raise an exception instead?
                Return Nothing
            End If
        End With
        Return theMapIndexFLayer
    End Function

    Private Sub initializeData()

        ' Initialize document and map objects, and their events for tool reference only
        _doc = EditorExtension.Application.Document
        _map = DirectCast(_doc, IMxDocument).FocusMap

        ' Obtain references to feature layer feature classes
        _theTaxlotFLayer = getTaxlotFeatureLayer()
        _theTaxlotFClass = _theTaxlotFLayer.FeatureClass
        _theMapIndexFLayer = getMapIndexFeatureLayer()
        _theMapIndexFClass = _theMapIndexFLayer.FeatureClass

        ' Get field indexes
        With EditorExtension.TaxLotSettings

            ' Find the ORMAP Taxlot field index
            _theOrmapTaxlotNumberFld = _theTaxlotFClass.FindField(.OrmapTaxlotField)
            If _theOrmapTaxlotNumberFld = FieldNotFoundIndex Then
                ' TODO: NIS Raise exception?
            End If

            ' Find the ORMAP Map Number field index
            _theOrmapMapNumberFld = _theTaxlotFClass.FindField(.OrmapMapNumberField)
            If _theOrmapMapNumberFld = FieldNotFoundIndex Then
                ' TODO: NIS Raise exception?
            End If

            ' Find the ORMAP Map Taxlot field index
            _theOrmapMapTaxlotFld = _theTaxlotFClass.FindField(.MapTaxlotField)
            If _theOrmapMapTaxlotFld = FieldNotFoundIndex Then
                ' TODO: NIS Raise exception?
            End If

            ' Find the Taxlot field index
            _theTLTaxlotFld = _theTaxlotFClass.FindField(.TaxlotField)
            If _theTLTaxlotFld = FieldNotFoundIndex Then
                ' TODO: NIS Raise exception?
            End If

            ' Find the Anomaly field index
            _theTLAnomalyFld = _theTaxlotFClass.FindField(.AnomalyField)
            If _theTLAnomalyFld = FieldNotFoundIndex Then
                ' TODO: NIS Raise exception?
            End If

        End With

    End Sub

#End Region

#Region "State Machine"

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
        If hasRequiredData() Then
            StateS1_1(StatePassageType.Entering)
        Else
            StateS1_2(StatePassageType.Entering)
        End If
    End Sub

#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Return MyBase.Enabled AndAlso _
                EditorExtension.Editor IsNot Nothing AndAlso _
                EditorExtension.IsValidWorkspace AndAlso _
                EditorExtension.HasValidLicense AndAlso _
                EditorExtension.AllowedToEditTaxlots AndAlso _
                State = CommandStateType.Enabled
        End Get
    End Property

#End Region

#Region "Methods"

    ''' <summary>
    ''' Occurs when this command is created.
    ''' </summary>
    ''' <param name="hook">A generic <c>Object</c> hook to an instance of the application.</param>
    ''' <remarks>The application hook may not point to an <c>IMxApplication</c> object.</remarks>
    Public Overrides Sub OnCreate(ByVal hook As Object)
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
                SetPartnerTaxlotAssignmentForm(New TaxlotAssignmentForm())
            End If
        End If

        ' TODO: Add other initialization code?

    End Sub

    Public Overrides Sub OnClick()

        '' HACK: NIS Don't know why this form is disposed after first use, but 
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

    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)

        ' TODO: NIS canContinue boolean is unecessary with Try..Catch..Finally -- revise to raise exceptions where canContinue = False

        Try
            initializeData()

            Dim canContinue As Boolean = True

            If canContinue Then
                ' TODO: NIS Confirm this .NET enum works with ESRI COM button parameter
                canContinue = (Button <> MouseButtons.Left)
            End If

            Dim isTaxlotType As Boolean = (StrComp(Me.TaxlotType, TaxlotAssignment.taxlotNumberTypeTaxlot, CompareMethod.Text) = 0)

            If canContinue Then
                'If taxlot numbering is selected, then make sure value is numeric
                If isTaxlotType Then
                    If IsNumeric(Me.NumberStartingFrom) Then
                        canContinue = True
                    Else
                        ' TODO: NIS Handle this another way?
                        canContinue = False
                    End If
                Else
                    canContinue = True
                End If
            End If

            ' Create a search shape out of the point that the user clicked
            Dim thePoint As IPoint = Nothing
            Dim theGeometry As IGeometry = Nothing
            If canContinue Then
                thePoint = EditorExtension.Editor.Display.DisplayTransformation.ToMapPoint(X, Y)
                'TODO: NIS Get rid of this commented line?
                'theGeometry = EditorExtension.Editor.CreateSearchShape(thePoint) 'Returns an IEnvelope  
                theGeometry = thePoint 'QI
            End If

            ' Verify the validity of the specified taxlot number, and uniqueness of
            ' numeric taxlot numbers
            If canContinue Then
                If isTaxlotType Then
                    If Not ValidateTaxlotNumber(CStr(Me.NumberStartingFrom), theGeometry) Then
                        If MessageBox.Show("The current Taxlot value (" & Me.NumberStartingFrom & ")" & vbNewLine & _
                                           "is not unique within this MapIndex." & vbNewLine & _
                                           "Attribute feature with value anyway?", _
                                           Me.Name, MessageBoxButtons.YesNo, _
                                           MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                            canContinue = False
                        Else
                            canContinue = True
                        End If
                    Else
                        canContinue = True
                    End If
                End If
            End If

            ' Insure the validity of the underlying map index polygon
            Dim theSpatialFilter As ISpatialFilter
            Dim theShapeFieldName As String
            Dim theMIFCursor As IFeatureCursor
            Dim theMIFeature As IFeature
            If canContinue Then
                theSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
                theSpatialFilter.Geometry = theGeometry
                theShapeFieldName = _theMapIndexFClass.ShapeFieldName
                theSpatialFilter.OutputSpatialReference(theShapeFieldName) = _map.SpatialReference
                theSpatialFilter.GeometryField = _theMapIndexFClass.ShapeFieldName
                theSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
                theMIFCursor = _theMapIndexFClass.Search(theSpatialFilter, False)
                theMIFeature = theMIFCursor.NextFeature
                If theMIFeature Is Nothing Then
                    MessageBox.Show("Unable to assign taxlot values to polygons that are not within a Map Index polygon.", _
                                    Me.Name, MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    canContinue = False
                End If
            End If

            Dim theTaxlotFCursor As IFeatureCursor = Nothing
            Dim theTaxlotFeature As IFeature = Nothing
            If canContinue Then
                ' Select any feature under the given point in the target layer
                theSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
                theSpatialFilter.Geometry = theGeometry
                theShapeFieldName = _theTaxlotFClass.ShapeFieldName
                theSpatialFilter.OutputSpatialReference(theShapeFieldName) = _map.SpatialReference
                theSpatialFilter.GeometryField = _theTaxlotFClass.ShapeFieldName
                theSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
                theTaxlotFCursor = _theTaxlotFClass.Search(theSpatialFilter, False)
                If theTaxlotFCursor IsNot Nothing Then
                    theTaxlotFeature = theTaxlotFCursor.NextFeature
                    'Update the feature
                    If theTaxlotFeature IsNot Nothing Then
                        canContinue = True
                        EditorExtension.Editor.StartOperation()
                    Else
                        '[No taxlots are selected...]
                        canContinue = False
                    End If
                End If
            End If

            Dim theExistOrmapMapNum As String = String.Empty 'initialize
            If canContinue Then
                ' The current ORMAP Number
                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                theExistOrmapMapNum = CStr(IIf(IsDBNull(theTaxlotFeature.Value(_theOrmapMapNumberFld)), "", theTaxlotFeature.Value(_theOrmapMapNumberFld)))

                ' Obtain the ORMAP Number from a MapIndex polygon if it is not present
                If Len(theExistOrmapMapNum) = 0 Then
                    ' TODO: NIS Confirm - This call will point _theMapIndexFLayer to the MapIndex feature class
                    CalculateTaxlotValues(theTaxlotFeature, _theMapIndexFLayer)

                    ' Refresh the ORMAP Number
                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                    theExistOrmapMapNum = CStr(IIf(IsDBNull(theTaxlotFeature.Value(_theOrmapMapNumberFld)), "", theTaxlotFeature.Value(_theOrmapMapNumberFld)))

                    ' Stop if there is still no ORMAP number
                    If Len(theExistOrmapMapNum) = 0 Then
                        MessageBox.Show("ORMAPMapNumber not present in this taxlot or MapIndex." & vbNewLine & _
                                        "Use the MapIndex tool to populate the ORMAPMapNumber field" & vbNewLine & _
                                        "before using this tool", Me.Name, MessageBoxButtons.OK)
                        canContinue = False
                    Else
                        canContinue = True
                    End If
                Else
                    canContinue = True
                End If
            End If

            ' Assign Taxlot value
            Dim theExistTLNumVal As String = String.Empty 'initialize
            If canContinue Then
                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                theExistTLNumVal = CStr(IIf(IsDBNull(theTaxlotFeature.Value(_theTLTaxlotFld)), "", theTaxlotFeature.Value(_theTLTaxlotFld)))

                ' Optionally, update the taxlot number field
                If Len(theExistTLNumVal) > 0 And theExistTLNumVal <> "0" Then
                    If MessageBox.Show("Taxlot currently has a Taxlot value (" & theExistTLNumVal & ")." & vbNewLine & _
                              "Update it?", Me.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        canContinue = False
                    Else
                        canContinue = True
                    End If
                Else
                    canContinue = True
                End If
            End If

            ' Taxlot numbers can be less than 5-digits.
            ' The Taxlot value in the OrmapMapNum field must be exactly 5 digits.
            ' Two versions of the taxlot number will be used for these purposes.
            Dim theNewTLNum As String = String.Empty 'initialize
            Dim theNewTLNum_5digit As String = String.Empty 'initialize
            If canContinue Then
                If isTaxlotType Then
                    theNewTLNum = CStr(Me.NumberStartingFrom) 'User entered number
                    theNewTLNum_5digit = theNewTLNum
                    ' Make sure number is 5 characters
                    If Len(theNewTLNum_5digit) < ORMAPNumber.GetOrmap_ORTaxlotFieldLength Then 'VB6: was ORMAP_TAXLOT_FIELD_LENGTH
                        Do Until Len(theNewTLNum_5digit) = ORMAPNumber.GetOrmap_ORTaxlotFieldLength 'VB6: was ORMAP_TAXLOT_FIELD_LENGTH
                            theNewTLNum_5digit = "0" & theNewTLNum_5digit
                        Loop
                    End If
                Else
                    ' Remove leading Zeros for taxlot number if any exist
                    theNewTLNum_5digit = Me.TaxlotType 'Predefined selection
                    theNewTLNum = Replace(theNewTLNum_5digit, "0", "")
                End If
                canContinue = True
            End If

            ' Determine if Special Interests field is something other than default
            ' If so, include it in ORMAPtaxlot number
            Dim theTLMapSuffixNumVal As String = String.Empty 'initialize
            Dim theTLMapSuffixTypeVal As String = String.Empty 'initialize
            If canContinue Then
                theTLMapSuffixTypeVal = GetMapSuffixType(theTaxlotFeature)
                theTLMapSuffixNumVal = GetMapSuffixNum(theTaxlotFeature)
                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                If IsDBNull(theTLMapSuffixTypeVal) Then theTLMapSuffixTypeVal = "0"
                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                If IsDBNull(theTLMapSuffixNumVal) Then theTLMapSuffixNumVal = "000"
                theTaxlotFeature.Value(_theTLTaxlotFld) = theNewTLNum
                canContinue = True
            End If

            ' Put together ORMAPTaxlot number from its parts
            Dim theShortOrmapNumber As String = String.Empty 'initialize
            Dim theCombinedOrmapNumber As String = String.Empty 'initialize
            If canContinue Then
                theShortOrmapNumber = OrmapMapNumberNoCountyCode(theExistOrmapMapNum)
                theCombinedOrmapNumber = theShortOrmapNumber & theTLMapSuffixTypeVal & theTLMapSuffixNumVal & theNewTLNum_5digit
                canContinue = True
            End If

            'Create  masked value from a combination of ORMapNum and the new taxlot
            'Note: Special code for Lane County (see comment below).
            Dim theOrmapMapTaxlot As String = String.Empty 'initialize
            Dim theDefaultCountyCode As Integer
            If canContinue Then
                theDefaultCountyCode = CInt(EditorExtension.DefaultValuesSettings.County)  ' TODO: NIS Confirm field choice
                Select Case theDefaultCountyCode
                    Case 1 To 19, 21 To 36
                        theOrmapMapTaxlot = CreateMapTaxlotValue(theExistOrmapMapNum & theNewTLNum_5digit, (EditorExtension.TaxLotSettings.MapTaxlotFormatMask))
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
                        theOrmapMapTaxlot = Trim(Left(theExistOrmapMapNum, 8)) & theNewTLNum_5digit
                End Select
                theTaxlotFeature.Value(_theOrmapMapTaxlotFld) = theOrmapMapTaxlot
                canContinue = True
            End If

            If canContinue Then
                'Assign OrmapTaxlot value
                theTaxlotFeature.Value(_theOrmapTaxlotNumberFld) = theCombinedOrmapNumber
                theTaxlotFeature.Store()

                'AutoIncrement if necessary
                If isTaxlotType AndAlso Me.IncrementNumber > 0 Then
                    Me.NumberStartingFrom += Me.IncrementNumber
                End If
            End If

            'Copy Anomaly from MapIndex
            Dim theORMAPNumberClass As New ORMAPNumber()
            If (theORMAPNumberClass.ParseNumber(theExistOrmapMapNum)) Then
                Dim theAnomalyVal As String
                If canContinue Then
                    theAnomalyVal = theORMAPNumberClass.Anomaly
                    theTaxlotFeature.Value(_theTLAnomalyFld) = theAnomalyVal
                End If
            End If

            ' End the edit operation
            EditorExtension.Editor.StopOperation("AutoIncrement Attribute")

            ' Select the feature
            _map.ClearSelection()
            _map.SelectFeature(_theTaxlotFLayer, theTaxlotFeature)

            ' Check for stacked features in the same location
            theTaxlotFeature = theTaxlotFCursor.NextFeature
            If Not theTaxlotFeature Is Nothing Then
                MessageBox.Show("Multiple (""vertical"") features found at this location. This tool can only edit one.", _
                                Me.Name, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                ' TODO: Enhance to handle more than one vertical feature.
            End If


            ' Update the view
            Dim theActiveView As IActiveView
            theActiveView = DirectCast(_map, IActiveView)

            ' Partially refresh the display
            theActiveView.PartialRefresh(esriViewDrawPhase.esriViewBackground Or _
                                         esriViewDrawPhase.esriViewGeography Or _
                                         esriViewDrawPhase.esriViewGraphics Or _
                                         esriViewDrawPhase.esriViewGraphicSelection, Nothing, Nothing)

        Catch ex As ApplicationException

            ' Handle the specific exception
            ' TODO: NIS

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

#End Region

#End Region

#Region "Implemented Interface Members"

#Region "IActiveViewEvents Implementation"

    Public Sub AfterDraw(ByVal Display As ESRI.ArcGIS.Display.IDisplay, ByVal phase As ESRI.ArcGIS.Carto.esriViewDrawPhase) Implements ESRI.ArcGIS.Carto.IActiveViewEvents.AfterDraw
    End Sub

    Public Sub AfterItemDraw(ByVal Index As Short, ByVal Display As ESRI.ArcGIS.Display.IDisplay, ByVal phase As ESRI.ArcGIS.esriSystem.esriDrawPhase) Implements ESRI.ArcGIS.Carto.IActiveViewEvents.AfterItemDraw
    End Sub

    Public Sub ContentsChanged() Implements ESRI.ArcGIS.Carto.IActiveViewEvents.ContentsChanged
    End Sub

    Public Sub ContentsCleared() Implements ESRI.ArcGIS.Carto.IActiveViewEvents.ContentsCleared
    End Sub

    Public Sub FocusMapChanged() Implements ESRI.ArcGIS.Carto.IActiveViewEvents.FocusMapChanged
        ' State Transistion E3
        TransitionE3()
    End Sub

    Public Sub ItemAdded(ByVal Item As Object) Implements ESRI.ArcGIS.Carto.IActiveViewEvents.ItemAdded
        ' State Transistion E3
        TransitionE3()
    End Sub

    Public Sub ItemDeleted(ByVal Item As Object) Implements ESRI.ArcGIS.Carto.IActiveViewEvents.ItemDeleted
        ' State Transistion E3
        TransitionE3()
    End Sub

    Public Sub ItemReordered(ByVal Item As Object, ByVal toIndex As Integer) Implements ESRI.ArcGIS.Carto.IActiveViewEvents.ItemReordered
    End Sub

    Public Sub SelectionChanged() Implements ESRI.ArcGIS.Carto.IActiveViewEvents.SelectionChanged
    End Sub

    Public Sub SpatialReferenceChanged() Implements ESRI.ArcGIS.Carto.IActiveViewEvents.SpatialReferenceChanged
    End Sub

    Public Sub ViewRefreshed(ByVal view As ESRI.ArcGIS.Carto.IActiveView, ByVal phase As ESRI.ArcGIS.Carto.esriViewDrawPhase, ByVal Data As Object, ByVal envelope As ESRI.ArcGIS.Geometry.IEnvelope) Implements ESRI.ArcGIS.Carto.IActiveViewEvents.ViewRefreshed
    End Sub

#End Region

#Region "IEditEvents Implementation"

    Public Sub AfterDrawSketch(ByVal pDpy As ESRI.ArcGIS.Display.IDisplay) Implements ESRI.ArcGIS.Editor.IEditEvents.AfterDrawSketch
    End Sub

    Public Sub OnChangeFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) Implements ESRI.ArcGIS.Editor.IEditEvents.OnChangeFeature
    End Sub

    Public Sub OnConflictsDetected() Implements ESRI.ArcGIS.Editor.IEditEvents.OnConflictsDetected
    End Sub

    Public Sub OnCreateFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) Implements ESRI.ArcGIS.Editor.IEditEvents.OnCreateFeature
    End Sub

    Public Sub OnCurrentLayerChanged() Implements ESRI.ArcGIS.Editor.IEditEvents.OnCurrentLayerChanged
    End Sub

    Public Sub OnCurrentTaskChanged() Implements ESRI.ArcGIS.Editor.IEditEvents.OnCurrentTaskChanged
    End Sub

    Public Sub OnDeleteFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) Implements ESRI.ArcGIS.Editor.IEditEvents.OnDeleteFeature
    End Sub

    Public Sub OnRedo() Implements ESRI.ArcGIS.Editor.IEditEvents.OnRedo
    End Sub

    Public Sub OnSelectionChanged() Implements ESRI.ArcGIS.Editor.IEditEvents.OnSelectionChanged
    End Sub

    Public Sub OnSketchFinished() Implements ESRI.ArcGIS.Editor.IEditEvents.OnSketchFinished
    End Sub

    Public Sub OnSketchModified() Implements ESRI.ArcGIS.Editor.IEditEvents.OnSketchModified
    End Sub

    Public Sub OnStartEditing() Implements ESRI.ArcGIS.Editor.IEditEvents.OnStartEditing
        ' State Transistion E1
        TransitionE1()
    End Sub

    Public Sub OnStopEditing(ByVal save As Boolean) Implements ESRI.ArcGIS.Editor.IEditEvents.OnStopEditing
        ' State Transistion E2
        TransitionE2()
    End Sub

    Public Sub OnUndo() Implements ESRI.ArcGIS.Editor.IEditEvents.OnUndo
    End Sub

#End Region

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



