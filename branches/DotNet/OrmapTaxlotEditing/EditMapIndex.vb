#Region "Copyright 2008 ORMAP Tech Group"

' File:  EditMapIndex.vb
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
Imports System
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports OrmapTaxlotEditing.DataMonitor
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.StringUtilities
Imports OrmapTaxlotEditing.Utilities
#End Region

<ComVisible(True)> _
<ComClass(EditMapIndex.ClassId, EditMapIndex.InterfaceId, EditMapIndex.EventsId), _
ProgId("ORMAPTaxlotEditing.EditMapIndex")> _
Public NotInheritable Class EditMapIndex
    Inherits BaseCommand
    Implements IDisposable

#Region "Class-Level Constants And Enumerations (none)"
#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' Define protected instance field values for the public properties
        MyBase.m_category = "OrmapToolbar"  'localizable text 
        MyBase.m_caption = "EditMapIndex"   'localizable text 
        MyBase.m_message = "Edit the selected MapIndex polygon and underlying Taxlot polygons."   'localizable text 
        MyBase.m_toolTip = "Edit MapIndex" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_EditMapIndex"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            ' Set the bitmap based on the name of the class.
            _bitmapResourceName = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), _bitmapResourceName)
        Catch ex As ArgumentException
            Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Structures"
    Friend Structure TaxlotFieldMap
        Friend Anomaly As Integer
        Friend County As Integer
        Friend MapAcres As Integer
        Friend MapNumber As Integer
        Friend MapTaxlotNumber As Integer
        Friend MapTaxlot As Integer
        Friend OrmapTaxlotNumber As Integer
        Friend OrmapMapNumber As Integer
        Friend PartialRangeCode As Integer
        Friend PartialTownshipCode As Integer
        Friend Quarter As Integer
        Friend QuarterQuarter As Integer
        Friend Range As Integer
        Friend RangeDirectional As Integer
        Friend Section As Integer
        Friend SuffixNumber As Integer
        Friend SuffixType As Integer
        Friend Taxlot As Integer
        Friend Township As Integer
        Friend TownshipDirectional As Integer
    End Structure

    Friend Structure MapIndexFieldMap
        Friend MapNumber As Integer
        Friend MapScale As Integer
        Friend ORMAPNumber As Integer
        Friend Page As Integer
        Friend Reliability As Integer
    End Structure
#End Region

#Region "Fields"

    Private _application As IApplication
    Private _bitmapResourceName As String

    Private _mapIndexFeatureClass As IFeatureClass
    Private _mapIndexFeature As IFeature
    Private _taxlotFeatureClass As IFeatureClass
    Private WithEvents _ormapNumber As ORMapNum
    Private _mapIndexFields As MapIndexFieldMap
    Private _taxlotFields As TaxlotFieldMap
    Private _editingState As Boolean = False

#End Region

#Region "Properties "

    Public Property EditingState() As Boolean
        Get
            EditingState = _editingState
        End Get
        Set(ByVal value As Boolean)
            _editingState = value
        End Set
    End Property

    Private WithEvents _partnerMapIndexForm As MapIndexForm

    Friend ReadOnly Property PartnerMapIndexForm() As MapIndexForm
        Get
            If _partnerMapIndexForm Is Nothing OrElse _partnerMapIndexForm.IsDisposed Then
                setPartnerMapIndexForm(New MapIndexForm())
            End If
            Return _partnerMapIndexForm
        End Get
    End Property

    Private Sub setPartnerMapIndexForm(ByRef value As MapIndexForm)
        If value IsNot Nothing Then
            _partnerMapIndexForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerMapIndexForm.Load, AddressOf PartnerMapIndexForm_Load
            AddHandler _partnerMapIndexForm.uxEdit.Click, AddressOf uxEdit_Click
            AddHandler _partnerMapIndexForm.uxHelp.Click, AddressOf uxHelp_Click
            AddHandler _partnerMapIndexForm.uxQuit.Click, AddressOf uxQuit_Click
            AddHandler _partnerMapIndexForm.uxCounty.Click, AddressOf uxCounty_Click
            AddHandler _partnerMapIndexForm.uxTownship.Click, AddressOf uxTown_Click
            AddHandler _partnerMapIndexForm.uxTownshipPartial.Click, AddressOf uxTownPart_Click
            AddHandler _partnerMapIndexForm.uxTownshipDirectional.Click, AddressOf uxTownDir_Click
            AddHandler _partnerMapIndexForm.uxRange.Click, AddressOf uxRange_Click
            AddHandler _partnerMapIndexForm.uxRangePartial.Click, AddressOf uxRangePart_Click
            AddHandler _partnerMapIndexForm.uxRangeDirectional.Click, AddressOf uxRangeDir_Click
            AddHandler _partnerMapIndexForm.uxSection.Click, AddressOf uxSection_Click
            AddHandler _partnerMapIndexForm.uxSectionQtr.Click, AddressOf uxQuarter_Click
            AddHandler _partnerMapIndexForm.uxSectionQtrQtr.Click, AddressOf uxQuarterQuarter_Click
            AddHandler _partnerMapIndexForm.uxSuffixType.Click, AddressOf uxSuffixType_Click
            AddHandler _partnerMapIndexForm.uxSuffixNumber.TextChanged, AddressOf uxSuffixNumber_TextChanged
            AddHandler _partnerMapIndexForm.uxAnomaly.TextChanged, AddressOf uxAnomaly_TextChanged
        Else
            ' Unsubscribe to partner form events.
            RemoveHandler _partnerMapIndexForm.Load, AddressOf PartnerMapIndexForm_Load
            RemoveHandler _partnerMapIndexForm.uxEdit.Click, AddressOf uxEdit_Click
            RemoveHandler _partnerMapIndexForm.uxQuit.Click, AddressOf uxQuit_Click
            RemoveHandler _partnerMapIndexForm.uxHelp.Click, AddressOf uxHelp_Click
            RemoveHandler _partnerMapIndexForm.uxCounty.Click, AddressOf uxCounty_Click
            RemoveHandler _partnerMapIndexForm.uxTownship.Click, AddressOf uxTown_Click
            RemoveHandler _partnerMapIndexForm.uxTownshipPartial.Click, AddressOf uxTownPart_Click
            RemoveHandler _partnerMapIndexForm.uxTownshipDirectional.Click, AddressOf uxTownDir_Click
            RemoveHandler _partnerMapIndexForm.uxRange.Click, AddressOf uxRange_Click
            RemoveHandler _partnerMapIndexForm.uxRangePartial.Click, AddressOf uxRangePart_Click
            RemoveHandler _partnerMapIndexForm.uxRangeDirectional.Click, AddressOf uxRangeDir_Click
            RemoveHandler _partnerMapIndexForm.uxSection.Click, AddressOf uxSection_Click
            RemoveHandler _partnerMapIndexForm.uxSectionQtr.Click, AddressOf uxQuarter_Click
            RemoveHandler _partnerMapIndexForm.uxSectionQtrQtr.Click, AddressOf uxQuarterQuarter_Click
            RemoveHandler _partnerMapIndexForm.uxSuffixType.Click, AddressOf uxSuffixType_Click
            RemoveHandler _partnerMapIndexForm.uxSuffixNumber.TextChanged, AddressOf uxSuffixNumber_TextChanged
            RemoveHandler _partnerMapIndexForm.uxAnomaly.TextChanged, AddressOf uxAnomaly_TextChanged
        End If
    End Sub

#End Region

#Region "Event Handlers"
    Private Sub PartnerMapIndexForm_Load(ByVal sender As Object, ByVal e As System.EventArgs)

        If DataMonitor.HasValidMapIndexData And DataMonitor.HasValidTaxlotData Then
            _mapIndexFeatureClass = DataMonitor.MapIndexFeatureLayer.FeatureClass
            _taxlotFeatureClass = DataMonitor.TaxlotFeatureLayer.FeatureClass
        Else
            'what
        End If

        InitializeFieldPositions()
        ToggleControls(False)

    End Sub

    Private Sub uxEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim thisDataSet As IDataset = _mapIndexFeatureClass.FeatureDataset
        Dim thisEditWorkSpace As IWorkspaceEdit = DirectCast(thisDataSet.Workspace, IWorkspaceEdit)

        Try
            Dim validData As Boolean

            If EditingState = True Then
                validData = True
                validData = validData And (Len(PartnerMapIndexForm.uxReliability.Text) <> 0)
                validData = validData And (Len(PartnerMapIndexForm.uxScale.Text) <> 0)
                validData = validData And (Len(PartnerMapIndexForm.uxMapNumber.Text) <> 0)
                validData = validData And (Len(PartnerMapIndexForm.uxPage.Text) <> 0)
                validData = validData And _ormapNumber.IsValidNumber

                If validData = False Then
                    MessageBox.Show("All fields must be filled in before assigning", "Invalid data", MessageBoxButtons.OK)
                    Exit Try
                End If

                'Begin edit process
                thisEditWorkSpace.StartEditOperation()
                'update form caption
                With PartnerMapIndexForm
                    .Text = "Map Index (Map Feature: " & _ormapNumber.GetORMapNum & ")"
                    'mapnumber
                    _mapIndexFeature.Value(_mapIndexFields.MapNumber) = .uxMapNumber.Text
                    'reliability
                    Dim value As String = ConvertCodeValueDomainToCode(_mapIndexFeature.Fields, EditorExtension.MapIndexSettings.ReliabilityCodeField, .uxReliability.Text)
                    Dim valueAsInteger As Integer

                    If Integer.TryParse(value, valueAsInteger) Then
                        _mapIndexFeature.Value(_mapIndexFields.Reliability) = valueAsInteger
                    Else
                        _mapIndexFeature.Value(_mapIndexFields.Reliability) = DBNull.Value
                    End If
                    'scale
                    value = ConvertCodeValueDomainToCode(_mapIndexFeature.Fields, EditorExtension.MapIndexSettings.MapScaleField, .uxScale.Text)
                    If Integer.TryParse(value, valueAsInteger) Then
                        _mapIndexFeature.Value(_mapIndexFields.MapScale) = valueAsInteger
                    Else
                        _mapIndexFeature.Value(_mapIndexFields.MapScale) = DBNull.Value
                    End If
                    'Page
                    value = .uxPage.Text
                    If Integer.TryParse(value, valueAsInteger) Then
                        _mapIndexFeature.Value(_mapIndexFields.Page) = valueAsInteger
                    Else
                        _mapIndexFeature.Value(_mapIndexFields.Page) = DBNull.Value
                    End If
                    'ORMAP Number
                    _mapIndexFeature.Value(_mapIndexFields.ORMAPNumber) = _ormapNumber.GetORMapNum
                    'store the edited feature
                    _mapIndexFeature.Store()
                    'Update all taxlot polygons that underlie this one
                    'UpdateTaxlots
                    'Finalize this edit
                    thisEditWorkSpace.StopEditOperation()
                End With 'PartnerMapindexForm
            End If 'Editing state = true
            'toggle form options after update
            EditingState = Not EditingState
            ToggleControls(EditingState)
            'update form caption
            PartnerMapIndexForm.Text = "Map Index (" & _ormapNumber.GetORMapNum & ")"

        Catch ex As Exception
            Trace.WriteLine(ex.ToString)
            If thisEditWorkSpace.IsBeingEdited Then
                thisEditWorkSpace.AbortEditOperation()
            End If
        End Try
    End Sub

    Private Sub uxQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If EditingState = True Then
            Me.InitForm()
            EditingState = False
            ToggleControls(EditingState)
        Else
            PartnerMapIndexForm.Close()
        End If
    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' TODO Evaluate help systems and implement.
        MessageBox.Show("uxHelp clicked")
    End Sub

    Private Sub uxCounty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'ms-help://MS.VSCC.v80/MS.MSDN.vAug06.en/dv_fxmancli/html/228112e1-1711-42ee-8ffa-ff3555bffe66.htm says the first parameter is a reference to the
        'object that raised the event.
        '_ormapNumber.County = sender.selectedtext

        _ormapNumber.County = PartnerMapIndexForm.uxCounty.SelectedText
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxTown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.Township = PartnerMapIndexForm.uxTownship.SelectedText
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxTownPart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.PartialTownshipCode = PartnerMapIndexForm.uxTownshipPartial.SelectedText
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxTownDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.TownshipDirectional = PartnerMapIndexForm.uxTownshipDirectional.SelectedText
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxRange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.Range = PartnerMapIndexForm.uxRange.SelectedText
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxRangePart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.PartialRangeCode = PartnerMapIndexForm.uxRangePartial.SelectedText
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxRangeDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.RangeDirectional = PartnerMapIndexForm.uxRangeDirectional.SelectedText
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxSection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.Section = PartnerMapIndexForm.uxSection.SelectedText
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxQuarter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.Quarter = PartnerMapIndexForm.uxSectionQtr.SelectedText
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxQuarterQuarter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.QuarterQuarter = PartnerMapIndexForm.uxSectionQtrQtr.SelectedText
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxSuffixType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.SuffixType = ConvertCodeValueDomainToCode(_taxlotFeatureClass.Fields, EditorExtension.TaxLotSettings.MapSuffixTypeField, PartnerMapIndexForm.uxSuffixType.SelectedText)
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxSuffixNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.SuffixNumber = PartnerMapIndexForm.uxSuffixNumber.Text
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxAnomaly_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.Anomaly = PartnerMapIndexForm.uxAnomaly.Text
        If EditingState = True Then
            PartnerMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

#End Region

#Region "Methods"

    Friend Sub DoButtonOperation()

        Try
            ' Check for valid data
            CheckValidDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If

            If Not HasValidTaxlotData Then
                PartnerMapIndexForm.uxEdit.Enabled = False
            Else
                PartnerMapIndexForm.uxEdit.Enabled = True
            End If
            PartnerMapIndexForm.ShowDialog()

        Catch ex As Exception
            Trace.WriteLine(ex.Message)

        End Try

    End Sub

    Private Sub InitializeFieldPositions()

        With _mapIndexFields
            .ORMAPNumber = _mapIndexFeatureClass.FindField(EditorExtension.MapIndexSettings.OrmapMapNumberField)
            .Reliability = _mapIndexFeatureClass.FindField(EditorExtension.MapIndexSettings.ReliabilityCodeField)
            .MapScale = _mapIndexFeatureClass.FindField(EditorExtension.MapIndexSettings.MapScaleField)
            .MapNumber = _mapIndexFeatureClass.FindField(EditorExtension.MapIndexSettings.MapNumberField)
            .Page = _mapIndexFeatureClass.FindField(EditorExtension.MapIndexSettings.PageNumberField)
        End With

        With _taxlotFields
            .Taxlot = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.TaxlotField)
            .Anomaly = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.AnomalyField)
            .County = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.CountyField)
            .OrmapMapNumber = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.OrmapMapNumberField)
            .OrmapTaxlotNumber = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.OrmapTaxlotField)
            .MapTaxlotNumber = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.MapTaxlotField)
            .PartialRangeCode = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.RangePartField)
            .PartialTownshipCode = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.TownshipPartField)
            .Quarter = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.QuarterSectionField)
            .QuarterQuarter = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.QuarterQuarterSectionField)
            .Range = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.RangeField)
            .RangeDirectional = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.RangeDirectionField)
            .Section = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.SectionNumberField)
            .SuffixNumber = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.MapSuffixNumberField)
            .SuffixType = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.MapSuffixTypeField)
            .Township = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.TownshipField)
            .TownshipDirectional = _taxlotFeatureClass.FindField(EditorExtension.TaxLotSettings.TownshipDirectionField)
        End With
    End Sub

    Private Sub ToggleControls(ByVal state As Boolean)
        Try
            Dim ctl As System.Windows.Forms.Control
            For Each ctl In PartnerMapIndexForm.Controls
                If TypeOf ctl Is ComboBox Or TypeOf ctl Is TextBox Then
                    ctl.Enabled = state
                End If
            Next

            With PartnerMapIndexForm
                If EditingState = True Then
                    .uxEdit.Text = "&Save"
                    .uxQuit.Text = "Cancel"
                Else
                    .uxEdit.Text = "&Edit"
                    .uxQuit.Text = "&Quit"
                End If
            End With
        Catch ex As Exception
            Trace.WriteLine(ex.ToString())
        End Try
    End Sub

    Private Function InitForm() As Boolean
        Try
            Dim thisFeatureCursor As IFeatureCursor = GetSelectedFeatures(DataMonitor.MapIndexFeatureLayer)
            If thisFeatureCursor Is Nothing Then
                Return False
                Exit Try
            End If
            _mapIndexFeature = thisFeatureCursor.NextFeature

            Dim thisTable As ITable = _mapIndexFeature.Table
            Dim thisRow As IRow = _mapIndexFeature.Table.GetRow(_mapIndexFeature.OID)

            'TODO jwm validate this method of retrieving values
            _ormapNumber.ParseNumber(ReadValue(thisRow, EditorExtension.MapIndexSettings.OrmapMapNumberField))
            If Not _ormapNumber.IsValidNumber Then
                'initform = initempty
                ToggleControls(True)
                EditingState = True
            Else
                'initform = init withfeature
            End If

            With PartnerMapIndexForm
                .uxORMAPNumberGroupBox.Text = "Map Index (" & _ormapNumber.GetORMapNum & ")"
                .uxORMAPNumberLabel.Text = _ormapNumber.GetORMapNum
                .Refresh()
            End With

        Catch ex As Exception
            Trace.WriteLine(ex.ToString)
        End Try
    End Function

    Private Function InitEmpty(ByVal mapIndexFields As IFields, ByVal taxlotFields As IFields) As Boolean
        Try
            _ormapNumber = New ORMapNum
            ResetControls()
            With _ormapNumber
                .County = EditorExtension.DefaultValuesSettings.County
                .Township = ""
                .TownshipDirectional = EditorExtension.DefaultValuesSettings.TownshipDirection
                .PartialTownshipCode = EditorExtension.DefaultValuesSettings.TownshipPart
                .Range = ""
                .RangeDirectional = EditorExtension.DefaultValuesSettings.RangeDirection
                .PartialRangeCode = EditorExtension.DefaultValuesSettings.RangePart
                .Section = ""
                .Quarter = EditorExtension.DefaultValuesSettings.QuarterSection
                .QuarterQuarter = EditorExtension.DefaultValuesSettings.QuarterQuarterSection
                .SuffixNumber = EditorExtension.DefaultValuesSettings.MapSuffixNumber
                .SuffixType = EditorExtension.DefaultValuesSettings.MapSuffixType
                .Anomaly = EditorExtension.DefaultValuesSettings.Anomaly
            End With
            With PartnerMapIndexForm
                .Text = "Map Index (Map Feature: <Not Attributed>)"
                'reliability
                AddCodesToCombo(EditorExtension.MapIndexSettings.ReliabilityCodeField, mapIndexFields, .uxReliability, "", True)
                'scale
                AddCodesToCombo(EditorExtension.MapIndexSettings.MapScaleField, mapIndexFields, .uxScale, "", True)
                'county
                AddCodesToCombo(EditorExtension.TaxLotSettings.CountyField, taxlotFields, .uxCounty, ConvertCodeValueDomainToDescription(taxlotFields, EditorExtension.TaxLotSettings.CountyField, _ormapNumber.County), True)
                'township
                AddCodesToCombo(EditorExtension.TaxLotSettings.TownshipDirectionField, taxlotFields, .uxTownship, _ormapNumber.Township, True)
                'partial township code
                AddCodesToCombo(EditorExtension.TaxLotSettings.TownshipPartField, taxlotFields, .uxTownshipPartial, ConvertCodeValueDomainToDescription(taxlotFields, EditorExtension.TaxLotSettings.TownshipPartField, _ormapNumber.PartialTownshipCode), True)
                'township directional
                AddCodesToCombo(EditorExtension.TaxLotSettings.TownshipDirectionField, taxlotFields, .uxTownshipDirectional, _ormapNumber.TownshipDirectional, True)
                'Ranges
                AddCodesToCombo(EditorExtension.TaxLotSettings.RangeField, taxlotFields, .uxRange, _ormapNumber.Range, True)
                'Partial range code
                AddCodesToCombo(EditorExtension.TaxLotSettings.RangePartField, taxlotFields, .uxRangePartial, ConvertCodeValueDomainToDescription(taxlotFields, EditorExtension.TaxLotSettings.RangePartField, _ormapNumber.PartialRangeCode), True)
                'Range directionals
                AddCodesToCombo(EditorExtension.TaxLotSettings.RangeDirectionField, taxlotFields, .uxRangeDirectional, _ormapNumber.RangeDirectional, True)
                'sections
                AddCodesToCombo(EditorExtension.TaxLotSettings.SectionNumberField, taxlotFields, .uxSection, _ormapNumber.Section, True)
                'Quarter
                AddCodesToCombo(EditorExtension.TaxLotSettings.QuarterSectionField, taxlotFields, .uxSectionQtr, _ormapNumber.Quarter, True)
                'Quarter Quarter
                AddCodesToCombo(EditorExtension.TaxLotSettings.QuarterQuarterSectionField, taxlotFields, .uxSectionQtrQtr, _ormapNumber.QuarterQuarter, True)
                'suffix type
                AddCodesToCombo(EditorExtension.TaxLotSettings.MapSuffixTypeField, taxlotFields, .uxSuffixType, _ormapNumber.SuffixType, True)
                'anomaly, page and suffix number
                .uxAnomaly.Text = _ormapNumber.Anomaly
                .uxPage.Text = "0"
                .uxSuffixNumber.Text = _ormapNumber.SuffixNumber
            End With

            Return True
        Catch ex As Exception
            Trace.WriteLine(ex.ToString)
            Return False
        End Try
    End Function

    Private Function InitWithFeature(ByVal feature As IFeature, ByVal mapIndexFields As IFields, ByVal taxlotFields As IFields) As Boolean
        Try
            ResetControls()

            Dim thisRow As IRow = feature.Table.GetRow(feature.OID)

            With PartnerMapIndexForm
                .Text = "Map Index (Map Feature: " & _ormapNumber.GetORMapNum
                .uxMapNumber.Text = ReadValue(thisRow, EditorExtension.MapIndexSettings.MapNumberField)
                'reliability
                AddCodesToCombo(EditorExtension.MapIndexSettings.ReliabilityCodeField, mapIndexFields, .uxReliability, ReadValue(thisRow, EditorExtension.MapIndexSettings.ReliabilityCodeField), True)
                'scale
                AddCodesToCombo(EditorExtension.MapIndexSettings.MapScaleField, mapIndexFields, .uxScale, ReadValue(thisRow, EditorExtension.MapIndexSettings.MapScaleField), True)
                'Page
                .uxPage.Text = ReadValue(thisRow, EditorExtension.MapIndexSettings.PageNumberField)
                'county
                AddCodesToCombo(EditorExtension.MapIndexSettings.CountyField, _taxlotFeatureClass.Fields, .uxCounty, ConvertCodeValueDomainToDescription(taxlotFields, EditorExtension.TaxLotSettings.CountyField, _ormapNumber.County), True)
                'township
                AddCodesToCombo(EditorExtension.TaxLotSettings.TownshipField, taxlotFields, .uxTownship, _ormapNumber.Township, True)
                'Partial township code
                AddCodesToCombo(EditorExtension.TaxLotSettings.TownshipPartField, taxlotFields, .uxTownshipPartial, "0" & _ormapNumber.PartialTownshipCode, True)
                'township directional
                AddCodesToCombo(EditorExtension.TaxLotSettings.TownshipDirectionField, taxlotFields, .uxTownshipDirectional, _ormapNumber.TownshipDirectional, True)
                'Ranges
                AddCodesToCombo(EditorExtension.TaxLotSettings.RangeField, taxlotFields, .uxRange, _ormapNumber.Range, True)
                'Partial range code
                AddCodesToCombo(EditorExtension.TaxLotSettings.RangePartField, taxlotFields, .uxRangePartial, "0" & _ormapNumber.PartialRangeCode, True)
                'Range directionals
                AddCodesToCombo(EditorExtension.TaxLotSettings.RangeDirectionField, taxlotFields, .uxRangeDirectional, _ormapNumber.RangeDirectional, True)
                'sections
                AddCodesToCombo(EditorExtension.TaxLotSettings.SectionNumberField, taxlotFields, .uxSection, _ormapNumber.Section, True)
                'Quarter
                AddCodesToCombo(EditorExtension.TaxLotSettings.QuarterSectionField, taxlotFields, .uxSectionQtr, _ormapNumber.Quarter, True)
                'Quarter Quarter
                AddCodesToCombo(EditorExtension.TaxLotSettings.QuarterQuarterSectionField, taxlotFields, .uxSectionQtrQtr, _ormapNumber.QuarterQuarter, True)
                'suffix type
                AddCodesToCombo(EditorExtension.TaxLotSettings.MapSuffixTypeField, taxlotFields, .uxSuffixType, ConvertCodeValueDomainToDescription(taxlotFields, EditorExtension.TaxLotSettings.MapSuffixTypeField, _ormapNumber.SuffixType), True)
                'anomaly and suffix number
                .uxAnomaly.Text = _ormapNumber.Anomaly
                .uxSuffixNumber.Text = _ormapNumber.SuffixNumber
            End With
            Return True
        Catch ex As Exception
            Trace.WriteLine(ex.ToString)
            Return False
        End Try

    End Function

    Private Sub ResetControls()

        Dim ctl As System.Windows.Forms.Control
        Dim cmb As System.Windows.Forms.ComboBox

        For Each ctl In PartnerMapIndexForm.Controls
            If TypeOf ctl Is TextBox Then
                ctl.Text = ""
            ElseIf TypeOf ctl Is ComboBox Then
                cmb = CType(ctl, ComboBox)
                'PartnerMapIndexForm.uxCounty.Items.Clear()
                cmb.Items.Clear()
            End If
        Next
    End Sub

    Private Function UpdateTaxlots(ByVal theFeature As IFeature, ByVal theTaxlotFeatureClass As IFeatureClass) As Boolean
        Try
            _application.StatusBar.Message(esriStatusBarPanes.esriStatusMain) = "Updating underlyling taxlot features..."
            ' Finds any taxlots that are underneath the map index polygon
            Dim thisSpatialQuery As ISpatialFilter = New SpatialFilter
            thisSpatialQuery.Geometry = theFeature.ShapeCopy
            thisSpatialQuery.SpatialRel = esriSpatialRelEnum.esriSpatialRelContains
            Dim thisFeatureSelection As IFeatureCursor = _taxlotFeatureClass.Update(thisSpatialQuery, False)
            'loop through the selected features
            Dim thisTaxlotFeature As IFeature = thisFeatureSelection.NextFeature
            Dim taxlot As String

            Do While Not thisTaxlotFeature Is Nothing
                'gets the formatted taxlot value
                If Not IsDBNull(thisTaxlotFeature.Value(_taxlotFields.Taxlot)) Then
                    taxlot = AddLeadingZeros(CStr(thisTaxlotFeature.Value(_taxlotFields.Taxlot)), ORMapNum.GetOrmap_TaxlotFieldLength)
                Else
                    taxlot = "00000"
                End If
                'special interest has been removed 
                'see Tracker 1922332 on http://sourceforge.net/tracker/index.php?func=detail&aid=1922332&group_id=151824&atid=782248

                'get mapnumber value
                Dim mapNumber As String
                If Not IsDBNull(thisTaxlotFeature.Value(_mapIndexFields.MapNumber)) Then
                    mapNumber = CStr(thisTaxlotFeature.Value(_mapIndexFields.MapNumber))
                Else
                    mapNumber = String.Empty
                End If
                'copy new attributes to  the taxlot table
                Dim mapTaxlotID As String = _ormapNumber.GetORMapNum & taxlot
                Dim countyCode As Short = CShort(EditorExtension.DefaultValuesSettings.County)
                Dim mapTaxlotValue As String = String.Empty
                Select Case countyCode
                    Case 1 To 19, 21 To 36
                        mapTaxlotValue = GenerateMapTaxlotValue(mapTaxlotID, EditorExtension.TaxLotSettings.MapTaxlotFormatMask)
                    Case 20
                        mapTaxlotValue = mapNumber.TrimEnd(CChar(mapNumber.Substring(0, 8))) & taxlot
                End Select
                With thisTaxlotFeature
                    .Value(_taxlotFields.County) = _ormapNumber.County
                    .Value(_taxlotFields.Township) = _ormapNumber.Township
                    .Value(_taxlotFields.PartialTownshipCode) = _ormapNumber.PartialTownshipCode
                    .Value(_taxlotFields.TownshipDirectional) = _ormapNumber.TownshipDirectional
                    .Value(_taxlotFields.Range) = _ormapNumber.Range
                    .Value(_taxlotFields.PartialRangeCode) = _ormapNumber.PartialRangeCode
                    .Value(_taxlotFields.RangeDirectional) = _ormapNumber.RangeDirectional
                    .Value(_taxlotFields.Section) = _ormapNumber.Section
                    .Value(_taxlotFields.Quarter) = _ormapNumber.Quarter
                    .Value(_taxlotFields.QuarterQuarter) = _ormapNumber.QuarterQuarter
                    .Value(_taxlotFields.SuffixType) = _ormapNumber.SuffixType
                    .Value(_taxlotFields.SuffixNumber) = _ormapNumber.SuffixNumber
                    .Value(_taxlotFields.Anomaly) = _ormapNumber.Anomaly
                    .Value(_taxlotFields.MapNumber) = theFeature.Value(_mapIndexFields.MapNumber)
                    .Value(_taxlotFields.OrmapMapNumber) = _ormapNumber.GetOrmapMapNumber
                    .Value(_taxlotFields.Taxlot) = CInt(taxlot)
                    'special interest used to go here
                    .Value(_taxlotFields.MapTaxlotNumber) = mapTaxlotValue
                    .Value(_taxlotFields.OrmapTaxlotNumber) = _ormapNumber.GetORMapNum & taxlot
                    .Store()
                End With
                thisTaxlotFeature = thisFeatureSelection.NextFeature
            Loop
            _application.StatusBar.Message(esriStatusBarPanes.esriStatusMain) = String.Empty

            thisTaxlotFeature = Nothing
            thisFeatureSelection = Nothing
            thisSpatialQuery = Nothing

        Catch ex As Exception
            Trace.WriteLine(ex.ToString)
        End Try

    End Function
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
            Return canEnable
        End Get
    End Property

#End Region

#Region "Methods"

    ''' <summary>
    ''' Called by ArcMap when this command is created.
    ''' </summary>
    ''' <param name="hook">A generic <c>Object</c> hook to an instance of the application.</param>
    ''' <remarks>The application hook may not point to an <c>IMxApplication</c> object.</remarks>
    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                _application = DirectCast(hook, IApplication)
                setPartnerMapIndexForm(New MapIndexForm())
                MyBase.m_enabled = True
                setPartnerEditMapIndexForm(New EditMapIndexForm)
            Else
                MyBase.m_enabled = False
            End If
        End If

        ' NOTE: Add other initialization code here...

    End Sub

    Public Overrides Sub OnClick()
        DoButtonOperation()
        'System.Windows.Forms.MessageBox.Show("Add EditMapIndex.OnClick implementation")
    End Sub

    Sub ORMAPNumber_OnChange(ByVal sender As Object, ByVal e As EventArgs) Handles _ormapNumber.OnChange
        PartnerMapIndexForm.uxORMAPNumberLabel.Text = _ormapNumber.GetORMapNum
    End Sub

#End Region

#End Region

#Region "Implemented Interface Members"

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
    Public Const ClassId As String = "2c5ecd6a-2175-4544-9a25-6281febb6d67"
    Public Const InterfaceId As String = "88034039-6ce9-46ed-973e-ffe70c3a3238"
    Public Const EventsId As String = "6432ad18-ea02-44c9-9589-0ef8cfb6898a"
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



