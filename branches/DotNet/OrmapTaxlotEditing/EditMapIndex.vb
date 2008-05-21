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
Imports System.Environment
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Windows.Forms.Control
Imports Microsoft.Practices.EnterpriseLibrary.ExceptionHandling
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
            Dim rethrow As Boolean = ExceptionPolicy.HandleException(ex, "Log Only Policy")
            If (rethrow) Then
                Throw
            End If
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
    Private _mapIndexFeature As IFeature
    Private WithEvents _ormapNumber As ORMapNum
    Private _mapIndexFields As MapIndexFieldMap
    Private _taxlotFields As TaxlotFieldMap
    Private _editingState As Boolean

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

    Private WithEvents _partnerEditMapIndexForm As EditMapIndexForm

    Friend ReadOnly Property PartnerEditMapIndexForm() As EditMapIndexForm
        Get
            If _partnerEditMapIndexForm Is Nothing OrElse _partnerEditMapIndexForm.IsDisposed Then
                setPartnerEditMapIndexForm(New EditMapIndexForm())
            End If
            Return _partnerEditMapIndexForm
        End Get
    End Property

    Private Sub setPartnerEditMapIndexForm(ByVal value As EditMapIndexForm)
        If value IsNot Nothing Then
            _partnerEditMapIndexForm = value
            ' Subscribe to partner form events.
            With _partnerEditMapIndexForm
                AddHandler .uxEdit.Click, AddressOf Me.uxEdit_Click
                AddHandler .uxHelp.Click, AddressOf Me.uxHelp_Click
                AddHandler .uxQuit.Click, AddressOf Me.uxQuit_Click
                AddHandler .uxCounty.Click, AddressOf Me.uxCounty_Click
                AddHandler .uxTownship.Click, AddressOf Me.uxTownship_Click
                AddHandler .uxTownshipPartial.Click, AddressOf Me.uxTownshipPartial_Click
                AddHandler .uxTownshipDirectional.Click, AddressOf Me.uxTownshipDirectional_Click
                AddHandler .uxRange.Click, AddressOf Me.uxRange_Click
                AddHandler .uxRangePartial.Click, AddressOf Me.uxRangePartial_Click
                AddHandler .uxRangeDirectional.Click, AddressOf Me.uxRangeDirectional_Click
                AddHandler .uxSection.Click, AddressOf Me.uxSection_Click
                AddHandler .uxSectionQtr.Click, AddressOf Me.uxSectionQtr_Click
                AddHandler .uxSectionQtrQtr.Click, AddressOf Me.uxSectionQrtrQrtr_Click
                AddHandler .uxSuffixType.Click, AddressOf Me.uxSuffixType_Click
                AddHandler .uxSuffixNumber.TextChanged, AddressOf Me.uxSuffixNumber_TextChanged
                AddHandler .uxAnomaly.TextChanged, AddressOf Me.uxAnomaly_TextChanged
            End With
        Else
            ' Unsubscribe to partner form events.
            With _partnerEditMapIndexForm
                RemoveHandler .uxEdit.Click, AddressOf Me.uxEdit_Click
                RemoveHandler .uxQuit.Click, AddressOf Me.uxQuit_Click
                RemoveHandler .uxHelp.Click, AddressOf Me.uxHelp_Click
                RemoveHandler .uxCounty.Click, AddressOf Me.uxCounty_Click
                RemoveHandler .uxTownship.Click, AddressOf Me.uxTownship_Click
                RemoveHandler .uxTownshipPartial.Click, AddressOf Me.uxTownshipPartial_Click
                RemoveHandler .uxTownshipDirectional.Click, AddressOf Me.uxTownshipDirectional_Click
                RemoveHandler .uxRange.Click, AddressOf Me.uxRange_Click
                RemoveHandler .uxRangePartial.Click, AddressOf Me.uxRangePartial_Click
                RemoveHandler .uxRangeDirectional.Click, AddressOf Me.uxRangeDirectional_Click
                RemoveHandler .uxSection.Click, AddressOf Me.uxSection_Click
                RemoveHandler .uxSectionQtr.Click, AddressOf Me.uxSectionQtr_Click
                RemoveHandler .uxSectionQtrQtr.Click, AddressOf Me.uxSectionQrtrQrtr_Click
                RemoveHandler .uxSuffixType.Click, AddressOf Me.uxSuffixType_Click
                RemoveHandler .uxSuffixNumber.TextChanged, AddressOf Me.uxSuffixNumber_TextChanged
                RemoveHandler .uxAnomaly.TextChanged, AddressOf Me.uxAnomaly_TextChanged
            End With
        End If

    End Sub

#End Region

#Region "Event Handlers"

    Private Sub PartnerMapIndexForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles _partnerEditMapIndexForm.Load
        initializeFieldPositions()
        toggleControls(False)
    End Sub

    Sub ORMAPNumber_OnChange(ByVal sender As Object, ByVal e As EventArgs) Handles _ormapNumber.OnChange
        PartnerEditMapIndexForm.uxORMAPNumberLabel.Text = _ormapNumber.GetORMapNum
    End Sub

    Private Sub uxEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim theMapIndexFC As IFeatureClass = DataMonitor.MapIndexFeatureLayer.FeatureClass
        Dim theMapIndexDataset As IDataset = DirectCast(theMapIndexFC, IDataset)
        Dim theEditWorkSpace As IWorkspaceEdit = TryCast(theMapIndexDataset.Workspace, IWorkspaceEdit)

        Try
            If EditingState Then
                Dim validData As Boolean = True
                validData = validData And (Len(PartnerEditMapIndexForm.uxReliability.Text) <> 0)
                validData = validData And (Len(PartnerEditMapIndexForm.uxScale.Text) <> 0)
                validData = validData And (Len(PartnerEditMapIndexForm.uxMapNumber.Text) <> 0)
                validData = validData And (Len(PartnerEditMapIndexForm.uxPage.Text) <> 0)
                validData = validData And _ormapNumber.IsValidNumber

                If Not validData Then
                    MessageBox.Show("Invalid data. All fields must be filled in before assigning.", "Edit Map Index", MessageBoxButtons.OK)
                    Exit Try
                End If
                If theEditWorkSpace IsNot Nothing Then
                    ' Begin edit process
                    theEditWorkSpace.StartEditOperation()

                    With PartnerEditMapIndexForm
                        ' Update form caption
                        .Text = "Map Index (Map Feature: " & _ormapNumber.GetORMapNum & ")"

                        ' Update all taxlot polygons that underlie this one

                        ' Set MapNumber
                        _mapIndexFeature.Value(_mapIndexFields.MapNumber) = .uxMapNumber.Text

                        ' Set Reliability
                        Dim value As String = ConvertCodeValueDomainToCode(_mapIndexFeature.Fields, EditorExtension.MapIndexSettings.ReliabilityCodeField, .uxReliability.Text)
                        Dim valueAsInteger As Integer
                        If Integer.TryParse(value, valueAsInteger) Then
                            _mapIndexFeature.Value(_mapIndexFields.Reliability) = valueAsInteger
                        Else
                            _mapIndexFeature.Value(_mapIndexFields.Reliability) = DBNull.Value
                        End If

                        ' Set MapScale
                        value = ConvertCodeValueDomainToCode(_mapIndexFeature.Fields, EditorExtension.MapIndexSettings.MapScaleField, .uxScale.Text)
                        If Integer.TryParse(value, valueAsInteger) Then
                            _mapIndexFeature.Value(_mapIndexFields.MapScale) = valueAsInteger
                        Else
                            _mapIndexFeature.Value(_mapIndexFields.MapScale) = DBNull.Value
                        End If

                        ' Set Page
                        value = .uxPage.Text
                        If Integer.TryParse(value, valueAsInteger) Then
                            _mapIndexFeature.Value(_mapIndexFields.Page) = valueAsInteger
                        Else
                            _mapIndexFeature.Value(_mapIndexFields.Page) = DBNull.Value
                        End If

                        ' Set ORMAPNumber
                        _mapIndexFeature.Value(_mapIndexFields.ORMAPNumber) = _ormapNumber.GetORMapNum

                        ' Store the edited feature
                        _mapIndexFeature.Store()

                    End With 'PartnerMapindexForm

                    ' Finalize this edit
                    theEditWorkSpace.StopEditOperation()

                End If
            End If 'EditingState = True

            ' Toggle form options after update
            EditingState = False
            toggleControls(EditingState)

            ' Update form caption
            PartnerEditMapIndexForm.Text = "Map Index (" & _ormapNumber.GetORMapNum & ")"

        Catch ex As Exception
            If theEditWorkSpace.IsBeingEdited Then
                theEditWorkSpace.AbortEditOperation()
            End If
            Trace.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub uxQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If EditingState = True Then
            Me.initForm()
            EditingState = False
            toggleControls(EditingState)
        Else
            PartnerEditMapIndexForm.Close()
        End If
    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' TODO: [ALL] Evaluate help systems and implement.
        MessageBox.Show("Sorry. Help not implemented at this time.")
    End Sub

    Private Sub uxCounty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'ms-help://MS.VSCC.v80/MS.MSDN.vAug06.en/dv_fxmancli/html/228112e1-1711-42ee-8ffa-ff3555bffe66.htm says the first parameter is a reference to the
        'object that raised the event.
        '_ormapNumber.County = sender.selectedtext

        _ormapNumber.County = PartnerEditMapIndexForm.uxCounty.SelectedText
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxTownship_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.Township = PartnerEditMapIndexForm.uxTownship.SelectedText
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxTownshipPartial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.PartialTownshipCode = PartnerEditMapIndexForm.uxTownshipPartial.SelectedText
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxTownshipDirectional_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.TownshipDirectional = PartnerEditMapIndexForm.uxTownshipDirectional.SelectedText
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxRange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.Range = PartnerEditMapIndexForm.uxRange.SelectedText
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxRangePartial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.PartialRangeCode = PartnerEditMapIndexForm.uxRangePartial.SelectedText
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxRangeDirectional_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.RangeDirectional = PartnerEditMapIndexForm.uxRangeDirectional.SelectedText
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxSection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.Section = PartnerEditMapIndexForm.uxSection.SelectedText
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxSectionQtr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.Quarter = PartnerEditMapIndexForm.uxSectionQtr.SelectedText
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxSectionQrtrQrtr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.QuarterQuarter = PartnerEditMapIndexForm.uxSectionQtrQtr.SelectedText
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxSuffixType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.SuffixType = ConvertCodeValueDomainToCode(DataMonitor.TaxlotFeatureLayer.FeatureClass.Fields, EditorExtension.TaxLotSettings.MapSuffixTypeField, PartnerEditMapIndexForm.uxSuffixType.SelectedText)
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxSuffixNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.SuffixNumber = PartnerEditMapIndexForm.uxSuffixNumber.Text
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

    Private Sub uxAnomaly_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        _ormapNumber.Anomaly = PartnerEditMapIndexForm.uxAnomaly.Text
        If EditingState = True Then
            PartnerEditMapIndexForm.uxEdit.Enabled = _ormapNumber.IsValidNumber
        Else
            PartnerEditMapIndexForm.uxEdit.Enabled = True
        End If
    End Sub

#End Region

#Region "Methods"

    Friend Sub DoButtonOperation()

        Try

            ' Check for valid data
            CheckValidMapIndexDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & NewLine & _
                                "Please load this dataset into your map.", _
                                "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            Else
                ' TODO: [ALL] Why does this return false when one or more MapIndex features are selected?
                'If Not HasSelectedFeatureCount(MapIndexFeatureLayer, 1) Then
                ' HACK: [NIS] This will get past the problem for now...
                Dim theEnumFeature As IEnumFeature
                Dim theFeature As IFeature
                Dim n As Integer = 0 'The MapIndex feature count
                theEnumFeature = EditorExtension.Editor.EditSelection
                theFeature = theEnumFeature.Next
                Do While (Not theFeature Is Nothing)
                    If theFeature.Class Is DataMonitor.MapIndexFeatureLayer.FeatureClass Then
                        n += 1
                        If n = 1 Then
                            _mapIndexFeature = theFeature
                        End If
                    End If
                    theFeature = theEnumFeature.Next
                Loop
                Select Case n 'The MapIndex feature count
                    Case 1
                        ' Do nothing, this is what we want.
                    Case 0
                        MessageBox.Show("No features selected in the MapIndex layer." & NewLine & _
                                        "Please select one MapIndex feature.", _
                                        "Edit Map Index", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Try
                    Case Else
                        MessageBox.Show(CStr(n) & " features selected in the MapIndex layer." & NewLine & _
                                        "Please select just one MapIndex feature.", _
                                        "Edit Map Index", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Try
                End Select
            End If

            CheckValidTaxlotDataProperties()
            If Not HasValidTaxlotData Then
                ' Enable control to allow the user to edit the ORMapIndex.
                PartnerEditMapIndexForm.uxEdit.Enabled = False
            Else
                PartnerEditMapIndexForm.uxEdit.Enabled = True
            End If
            If Me.initForm Then
                PartnerEditMapIndexForm.ShowDialog()
            End If

        Catch ex As Exception
            Trace.WriteLine(ex.Message)
        End Try

    End Sub

    Private Sub initializeFieldPositions()
        Dim mapindexFC As IFeatureClass = DataMonitor.MapIndexFeatureLayer.FeatureClass
        Dim taxlotFC As IFeatureClass = DataMonitor.TaxlotFeatureLayer.FeatureClass

        With _mapIndexFields
            .ORMAPNumber = mapindexFC.FindField(EditorExtension.MapIndexSettings.OrmapMapNumberField)
            .Reliability = mapindexFC.FindField(EditorExtension.MapIndexSettings.ReliabilityCodeField)
            .MapScale = mapindexFC.FindField(EditorExtension.MapIndexSettings.MapScaleField)
            .MapNumber = mapindexFC.FindField(EditorExtension.MapIndexSettings.MapNumberField)
            .Page = mapindexFC.FindField(EditorExtension.MapIndexSettings.PageNumberField)
        End With

        With _taxlotFields
            .Taxlot = taxlotFC.FindField(EditorExtension.TaxLotSettings.TaxlotField)
            .Anomaly = taxlotFC.FindField(EditorExtension.TaxLotSettings.AnomalyField)
            .County = taxlotFC.FindField(EditorExtension.TaxLotSettings.CountyField)
            .OrmapMapNumber = taxlotFC.FindField(EditorExtension.TaxLotSettings.OrmapMapNumberField)
            .OrmapTaxlotNumber = taxlotFC.FindField(EditorExtension.TaxLotSettings.OrmapTaxlotField)
            .MapTaxlotNumber = taxlotFC.FindField(EditorExtension.TaxLotSettings.MapTaxlotField)
            .PartialRangeCode = taxlotFC.FindField(EditorExtension.TaxLotSettings.RangePartField)
            .PartialTownshipCode = taxlotFC.FindField(EditorExtension.TaxLotSettings.TownshipPartField)
            .Quarter = taxlotFC.FindField(EditorExtension.TaxLotSettings.QuarterSectionField)
            .QuarterQuarter = taxlotFC.FindField(EditorExtension.TaxLotSettings.QuarterQuarterSectionField)
            .Range = taxlotFC.FindField(EditorExtension.TaxLotSettings.RangeField)
            .RangeDirectional = taxlotFC.FindField(EditorExtension.TaxLotSettings.RangeDirectionField)
            .Section = taxlotFC.FindField(EditorExtension.TaxLotSettings.SectionNumberField)
            .SuffixNumber = taxlotFC.FindField(EditorExtension.TaxLotSettings.MapSuffixNumberField)
            .SuffixType = taxlotFC.FindField(EditorExtension.TaxLotSettings.MapSuffixTypeField)
            .Township = taxlotFC.FindField(EditorExtension.TaxLotSettings.TownshipField)
            .TownshipDirectional = taxlotFC.FindField(EditorExtension.TaxLotSettings.TownshipDirectionField)
        End With

    End Sub

    Private Sub toggleControls(ByVal state As Boolean)
        Try
            Dim ctl As System.Windows.Forms.Control
            For Each ctl In PartnerEditMapIndexForm.Controls
                If TypeOf ctl Is ComboBox Or TypeOf ctl Is TextBox Then
                    ctl.Enabled = state
                End If
                If ctl.Controls.Count > 0 Then
                    For Each subControl As Control In ctl.Controls
                        If TypeOf subControl Is TextBox Or TypeOf subControl Is ComboBox Then
                            subControl.Enabled = state
                        End If
                    Next
                End If
            Next

            With PartnerEditMapIndexForm
                If EditingState = True Then
                    .uxEdit.Text = "&Save"
                    .uxQuit.Text = "&Cancel"
                Else
                    .uxEdit.Text = "&Edit"
                    .uxQuit.Text = "&Quit"
                End If
            End With

        Catch ex As Exception
            Trace.WriteLine(ex.ToString())
        End Try
    End Sub

    Private Sub LoopThruSelection()
        Dim pEnumFeat As IEnumFeature
        Dim pFeat As IFeature

        pEnumFeat = EditorExtension.Editor.EditSelection
        pFeat = pEnumFeat.Next
        Do While (Not pFeat Is Nothing)
            Debug.Print(CStr(pFeat.Value(1)))
            pFeat = pEnumFeat.Next
        Loop
    End Sub

    Private Function initForm() As Boolean
        ' HACK: [NIS] GetSelectedFeatures does not seem to work for the MapIndex feature layer.
        'Dim thisFeatureCursor As IFeatureCursor = GetSelectedFeatures(DataMonitor.MapIndexFeatureLayer)
        '
        'If thisFeatureCursor Is Nothing Then
        '    Return False
        'End If
        '_mapIndexFeature = thisFeatureCursor.NextFeature

        Dim theRow As IRow = _mapIndexFeature.Table.GetRow(_mapIndexFeature.OID)
        If theRow Is Nothing Then
            Return False
        End If

        _ormapNumber = New ORMapNum
        _ormapNumber.ParseNumber(ReadValue(theRow, EditorExtension.MapIndexSettings.OrmapMapNumberField))

        If _ormapNumber.IsValidNumber Then
            initForm = initWithFeature(_mapIndexFeature, DataMonitor.MapIndexFeatureLayer.FeatureClass.Fields, DataMonitor.TaxlotFeatureLayer.FeatureClass.Fields)
        Else
            initForm = initEmpty(DataMonitor.MapIndexFeatureLayer.FeatureClass.Fields, DataMonitor.TaxlotFeatureLayer.FeatureClass.Fields)
            toggleControls(True)
            EditingState = True
        End If

        With PartnerEditMapIndexForm
            .uxORMAPNumberGroupBox.Text = "Map Index (" & _ormapNumber.GetORMapNum & ")"
            .uxORMAPNumberLabel.Text = _ormapNumber.GetORMapNum
            .Refresh()
        End With

    End Function

    ''' <summary>
    ''' Initialize the form for a new selection
    ''' </summary>
    ''' <param name="mapIndexFields"></param>
    ''' <param name="taxlotFields"></param>
    ''' <returns></returns>
    ''' <remarks>Clear the form and set default value for fields in accordance with a selection that has no values set</remarks>
    Private Function initEmpty(ByVal mapIndexFields As IFields, ByVal taxlotFields As IFields) As Boolean
        Try
            resetControls(PartnerEditMapIndexForm.Controls, True)

            _ormapNumber = New ORMapNum
            With _ormapNumber
                .County = Integer.Parse(EditorExtension.DefaultValuesSettings.County).ToString
                .Township = String.Empty
                .TownshipDirectional = EditorExtension.DefaultValuesSettings.TownshipDirection
                .PartialTownshipCode = EditorExtension.DefaultValuesSettings.TownshipPart
                .Range = String.Empty
                .RangeDirectional = EditorExtension.DefaultValuesSettings.RangeDirection
                .PartialRangeCode = EditorExtension.DefaultValuesSettings.RangePart
                .Section = String.Empty
                .Quarter = EditorExtension.DefaultValuesSettings.QuarterSection
                .QuarterQuarter = EditorExtension.DefaultValuesSettings.QuarterQuarterSection
                .SuffixNumber = EditorExtension.DefaultValuesSettings.MapSuffixNumber
                .SuffixType = EditorExtension.DefaultValuesSettings.MapSuffixType
                .Anomaly = EditorExtension.DefaultValuesSettings.Anomaly
            End With

            With PartnerEditMapIndexForm
                .Text = "Map Index (Map Feature: <Not Attributed>)"
                'reliability
                AddCodesToCombo(EditorExtension.MapIndexSettings.ReliabilityCodeField, mapIndexFields, .uxReliability, "", True)
                'scale
                AddCodesToCombo(EditorExtension.MapIndexSettings.MapScaleField, mapIndexFields, .uxScale, "", True)
                'county
                AddCodesToCombo(EditorExtension.TaxLotSettings.CountyField, taxlotFields, .uxCounty, ConvertCodeValueDomainToDescription(taxlotFields, EditorExtension.TaxLotSettings.CountyField, Integer.Parse(_ormapNumber.County).ToString), True)
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

    Private Function initWithFeature(ByVal feature As IFeature, ByVal mapIndexFields As IFields, ByVal taxlotFields As IFields) As Boolean
        Try
            Dim thisRow As IRow = feature.Table.GetRow(feature.OID)

            With PartnerEditMapIndexForm
                resetControls(.Controls, True)

                .Text = "Map Index (Map Feature: " & _ormapNumber.GetORMapNum
                .uxMapNumber.Text = ReadValue(thisRow, EditorExtension.MapIndexSettings.MapNumberField)
                'reliability
                AddCodesToCombo(EditorExtension.MapIndexSettings.ReliabilityCodeField, mapIndexFields, .uxReliability, ReadValue(thisRow, EditorExtension.MapIndexSettings.ReliabilityCodeField), True)
                'scale
                AddCodesToCombo(EditorExtension.MapIndexSettings.MapScaleField, mapIndexFields, .uxScale, ReadValue(thisRow, EditorExtension.MapIndexSettings.MapScaleField), True)
                'Page
                .uxPage.Text = ReadValue(thisRow, EditorExtension.MapIndexSettings.PageNumberField)
                'county
                AddCodesToCombo(EditorExtension.MapIndexSettings.CountyField, taxlotFields, .uxCounty, ConvertCodeValueDomainToDescription(taxlotFields, EditorExtension.TaxLotSettings.CountyField, Integer.Parse(_ormapNumber.County).ToString), True)
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

    Private Overloads Sub resetControls(ByVal inControls As ControlCollection)
        resetControls(inControls, True)
    End Sub

    Private Overloads Sub resetControls(ByVal inControls As ControlCollection, ByVal inRecursive As Boolean)

        Dim ctl As Control
        For Each ctl In inControls
            If TypeOf ctl Is IContainerControl Then
                resetControls(ctl.Controls, inRecursive)
            ElseIf TypeOf ctl Is TextBox Then
                ctl.Text = ""
            ElseIf TypeOf ctl Is ComboBox Then
                Dim cmb As ComboBox = DirectCast(ctl, ComboBox)
                cmb.Items.Clear()
            End If
        Next ctl

    End Sub

    Private Function updateTaxlots(ByVal theMapIndexFeature As IFeature) As Boolean
        Try
            _application.StatusBar.Message(esriStatusBarPanes.esriStatusMain) = "Updating underlyling taxlot features..."
            ' Finds any taxlots that are underneath the map index polygon
            Dim thisSpatialQuery As ISpatialFilter = New SpatialFilter
            thisSpatialQuery.Geometry = theMapIndexFeature.ShapeCopy
            thisSpatialQuery.SpatialRel = esriSpatialRelEnum.esriSpatialRelContains
            Dim thisFeatureSelection As IFeatureCursor = DataMonitor.TaxlotFeatureLayer.FeatureClass.Update(thisSpatialQuery, False)
            'loop through the selected features
            Dim thisTaxlotFeature As IFeature = thisFeatureSelection.NextFeature
            Dim taxlot As String

            Do While Not thisTaxlotFeature Is Nothing
                'gets the formatted taxlot value
                If Not IsDBNull(thisTaxlotFeature.Value(_taxlotFields.Taxlot)) Then
                    taxlot = AddLeadingZeros(CStr(thisTaxlotFeature.Value(_taxlotFields.Taxlot)), ORMapNum.GetTaxlotFieldLength)
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
                    .Value(_taxlotFields.MapNumber) = theMapIndexFeature.Value(_mapIndexFields.MapNumber)
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
                setPartnerEditMapIndexForm(New EditMapIndexForm)
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If

        ' NOTE: Add other initialization code here...

    End Sub

    Public Overrides Sub OnClick()
        DoButtonOperation()
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




