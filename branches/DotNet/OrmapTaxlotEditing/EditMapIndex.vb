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
            MyBase.m_bitmap = My.Resources.ORMAPToolBarResource.EditMapIndex
        Catch ex As ArgumentException
            Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

    Public Structure TaxlotFieldMap
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

    Public Structure MapIndexFieldMap
        Friend MapNumber As Integer
        Friend MapScale As Integer
        Friend ORMAPNumber As Integer
        Friend Page As Integer
        Friend Reliability As Integer
    End Structure

#Region "Fields"

    Private _application As IApplication
    Private _mapIndexFeatureClass As IFeatureClass
    Private _mapIndexFeature As IFeature
    Private _taxlotFeatureClass As IFeatureClass
    Private _ormapNumber As ORMapNum
    Private _mapIndexFields As MapIndexFieldMap
    Private _taxlotFields As TaxlotFieldMap
    Private _editingState As Boolean
#End Region

#Region "Properties "
    Private WithEvents _partnerMapIndexForm As MapIndexForm

    Friend ReadOnly Property PartnerMapIndexForm() As MapIndexForm
        Get
            Return _partnerMapIndexForm
        End Get
    End Property

    Public Property EditingState() As Boolean
        Get
            EditingState = _editingState
        End Get
        Set(ByVal value As Boolean)
            _editingState = value
        End Set
    End Property
    Private Sub setPartnerLocateFeatureForm(ByRef value As MapIndexForm)
        If value IsNot Nothing Then
            _partnerMapIndexForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerMapIndexForm.Load, AddressOf PartnerMapIndexForm_Load
            AddHandler _partnerMapIndexForm.uxEdit.Click, AddressOf uxEdit_Click
            AddHandler _partnerMapIndexForm.uxQuit.Click, AddressOf uxQuit_Click
        Else
            ' Unsubscribe to partner form events.
            RemoveHandler _partnerMapIndexForm.Load, AddressOf PartnerMapIndexForm_Load
            RemoveHandler _partnerMapIndexForm.uxEdit.Click, AddressOf uxEdit_Click
            RemoveHandler _partnerMapIndexForm.uxQuit.Click, AddressOf uxQuit_Click
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

    End Sub

    Private Sub uxQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' TODO Evaluate help systems and implement.
        MessageBox.Show("uxHelp clicked")
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
            MessageBox.Show(ex.Message)

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
        Dim ctl As System.Windows.Forms.Control
        For Each ctl In PartnerMapIndexForm.Controls
            If TypeOf ctl Is ComboBox Or TypeOf ctl Is TextBox Then
                ctl.Enabled = state
            End If
        Next

        With PartnerMapIndexForm
            If state Then
                .uxEdit.Text = "&Save"
                .uxQuit.Text = "Cancel"
            Else
                .uxEdit.Text = "&Edit"
                .uxQuit.Text = "&Quit"
            End If
        End With

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
                AddCodesToCombo(EditorExtension.MapIndexSettings.ReliabilityCodeField, mapIndexFields, .uxReliability, "", True)
                AddCodesToCombo(EditorExtension.MapIndexSettings.MapScaleField, mapIndexFields, .uxScale, "", True)
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
            _application = DirectCast(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
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

#End Region

#End Region

#Region "Implemented Interface Members (none)"
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



