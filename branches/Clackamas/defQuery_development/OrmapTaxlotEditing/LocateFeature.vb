#Region "Copyright 2008 ORMAP Tech Group"

' File:  LocateFeature.vb
'
' Original Author:  OPET.NET Migration Team (Shad Campbell, James Moore, 
'                   Nick Seigal)
'
' Date Created:  January 8, 2008
'
' Copyright Holder:  ORMAP Tech Group  
' Contact Info:  ORMAP Tech Group may be reached at 
' ORMAP_ESRI_Programmers@listsmart.osl.state.or.us
'
' This file is part of the ORMAP Taxlot Editing Toolbar.
'
' ORMAP Taxlot Editing Toolbar is free software; you can redistribute it and/or
' modify it under the terms of the Lesser GNU General Public License as 
' published by the Free Software Foundation; either version 3 of the License, 
' or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT 
' ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or 
' FITNESS FOR A PARTICULAR PURPOSE.  See the Lesser GNU General Public License 
' located in the COPYING.LESSER.txt file for more details.
'
' You should have received a copy of the Lesser GNU General Public License 
' along with the ORMAP Taxlot Editing Toolbar; if not, write to the Free 
' Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 
' 02110-1301 USA.

#End Region

#Region "Subversion Keyword Expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision: 406 $
'Date of Last Change: $Date: 2009-11-30 22:49:20 -0800 (Mon, 30 Nov 2009) $
#End Region

#Region "Imported Namespaces"
Imports System.Drawing
Imports System.Environment
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports OrmapTaxlotEditing.DataMonitor
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.StringUtilities
Imports OrmapTaxlotEditing.Utilities
Imports OrmapTaxlotEditing.OrmapSettings

#End Region

''' <summary>
''' Provides an ArcMap Command with functionality to 
''' allow users to find and zoom to MapIndex or Taxlot 
''' features and/or specify a MapIndex (Map) to override the
''' auto attributing of features via the on create event.
''' </summary>
''' <remarks><seealso cref="LocateFeatureDockWin"/><seealso cref="LocateFeatureUserControl"/></remarks>
<ComVisible(True)> _
<ComClass(LocateFeature.ClassId, LocateFeature.InterfaceId, LocateFeature.EventsId), _
ProgId("ORMAPTaxlotEditing.LocateFeature")> _
Public NotInheritable Class LocateFeature
    Inherits BaseCommand
    Implements IDisposable

#Region "Class-Level Constants and Enumerations (none)"
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
        MyBase.m_caption = "LocateFeature"   'localizable text 
        MyBase.m_message = "Locate a Taxlot or Mapindex"   'localizable text 
        MyBase.m_toolTip = "Locate Taxlot or Mapindex" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_LocateFeature"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            ' Set the bitmap based on the name of the class.
            _bitmapResourceName = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), _bitmapResourceName)
        Catch ex As ArgumentException
            EditorExtension.ProcessUnhandledException(ex)
        End Try

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication
    Private _bitmapResourceName As String
    Private _locateFeatureDockWinMgr As IDockableWindowManager
    Private _locateFeatureDockWin As IDockableWindow

#End Region

#Region "Properties"

    Private WithEvents _partnerLocateFeatureUserControl As LocateFeatureUserControl

    Friend ReadOnly Property PartnerLocateFeatureUserControl() As LocateFeatureUserControl
        Get
            If _partnerLocateFeatureUserControl Is Nothing OrElse _partnerLocateFeatureUserControl.IsDisposed Then
                setPartnerLocateFeatureUserControl(New LocateFeatureUserControl())
            End If
            Return _partnerLocateFeatureUserControl
        End Get
    End Property

    Private Sub setPartnerLocateFeatureUserControl(ByVal value As LocateFeatureUserControl)
        If value IsNot Nothing Then
            _partnerLocateFeatureUserControl = value
            ' Subscribe to partner form events.
            AddHandler _partnerLocateFeatureUserControl.uxMapNumber.TextChanged, AddressOf uxMapNumber_TextChanged
            AddHandler _partnerLocateFeatureUserControl.uxTaxlot.Enter, AddressOf uxTaxlot_Enter
            AddHandler _partnerLocateFeatureUserControl.uxFind.Click, AddressOf uxFind_Click
            AddHandler _partnerLocateFeatureUserControl.uxHelp.Click, AddressOf uxHelp_Click
            AddHandler _partnerLocateFeatureUserControl.uxSelectFeatures.CheckedChanged, AddressOf uxSelectFeatures_CheckedChanged
            AddHandler _partnerLocateFeatureUserControl.uxSetAttributeMode.Click, AddressOf uxSetAttributeMode_Click
            AddHandler _partnerLocateFeatureUserControl.uxTimer.Tick, AddressOf uxTimer_Tick
            AddHandler _partnerLocateFeatureUserControl.uxORMAPProperties.Click, AddressOf uxORMAPProperties_Click
            AddHandler _partnerLocateFeatureUserControl.uxSetDefinitionQuery.Click, AddressOf uxSetDefinitionQuery_Click
            AddHandler _partnerLocateFeatureUserControl.uxClearDefinitionQuery.Click, AddressOf uxClearDefinitionQuery_Click

        Else
            ' Unsubscribe to partner form events.
            RemoveHandler _partnerLocateFeatureUserControl.uxMapNumber.TextChanged, AddressOf uxMapNumber_TextChanged
            RemoveHandler _partnerLocateFeatureUserControl.uxTaxlot.Enter, AddressOf uxTaxlot_Enter
            RemoveHandler _partnerLocateFeatureUserControl.uxFind.Click, AddressOf uxFind_Click
            RemoveHandler _partnerLocateFeatureUserControl.uxHelp.Click, AddressOf uxHelp_Click
            RemoveHandler _partnerLocateFeatureUserControl.uxSelectFeatures.CheckedChanged, AddressOf uxSelectFeatures_CheckedChanged
            RemoveHandler _partnerLocateFeatureUserControl.uxSetAttributeMode.Click, AddressOf uxSetAttributeMode_Click
            RemoveHandler _partnerLocateFeatureUserControl.uxTimer.Tick, AddressOf uxTimer_Tick
            RemoveHandler _partnerLocateFeatureUserControl.uxORMAPProperties.Click, AddressOf uxORMAPProperties_Click
            RemoveHandler _partnerLocateFeatureUserControl.uxSetDefinitionQuery.Click, AddressOf uxSetDefinitionQuery_Click
            RemoveHandler _partnerLocateFeatureUserControl.uxClearDefinitionQuery.Click, AddressOf uxClearDefinitionQuery_Click
        End If
    End Sub

    Private _uxSelectFeaturesChecked As Boolean = False
    Friend Property uxSelectFeaturesChecked() As Boolean
        Get
            Return _uxSelectFeaturesChecked
        End Get
        Set(ByVal value As Boolean)
            _uxSelectFeaturesChecked = value
        End Set
    End Property

    Private _mapIndexHasBeenChanged As Boolean = False

#End Region

#Region "Event Handlers"

    Private Sub PartnerLocateFeatureForm_Load()

        With PartnerLocateFeatureUserControl
            Try
                .UseWaitCursor = True

                '' NOTE: [SC] Calculating a AutoCompleteCustomSource using the List of strings is considerably faster than
                '' adding with the .add method.
                Dim mapNumberStringList As List(Of String) = New List(Of String)
                Dim theMapIndexFClass As IFeatureClass = MapIndexFeatureLayer.FeatureClass
                Dim theQueryFilter As IQueryFilter = New QueryFilter
                theQueryFilter.SubFields = EditorExtension.MapIndexSettings.MapNumberField
                Dim theFeatCursor As IFeatureCursor = theMapIndexFClass.Search(theQueryFilter, True)
                Dim theFeature As IFeature = theFeatCursor.NextFeature
                If theFeature Is Nothing Then Exit Sub
                Dim theFieldIdx As Integer = theFeature.Fields.FindField(EditorExtension.MapIndexSettings.MapNumberField)
                Do Until theFeature Is Nothing
                    Dim theMapNumberVal As String = theFeature.Value(theFieldIdx).ToString
                    If Not mapNumberStringList.Contains(theMapNumberVal) Then
                        mapNumberStringList.Add(theMapNumberVal)
                    End If
                    theFeature = theFeatCursor.NextFeature
                Loop

                ' Populate the combobox from the List object
                .uxMapNumber.AutoCompleteCustomSource.AddRange(mapNumberStringList.ToArray)

                ' Set control defaults
                .uxMapNumber.Text = String.Empty
                .uxTaxlot.Text = String.Empty
                If uxSelectFeaturesChecked Then .uxSelectFeatures.Checked = True

            Finally
                .UseWaitCursor = False
            End Try
        End With

    End Sub


    Private Sub uxTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxTimer.Tick

        PartnerLocateFeatureUserControl.uxEditingGroupBox.Enabled = EditorExtension.AllowedToAutoUpdateAllFields

    End Sub

    Private Sub uxSetAttributeMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxSetAttributeMode.Click

        With PartnerLocateFeatureUserControl

            If EditorExtension.OverrideAutoAttribute Then
                .uxSetAttributeMode.Text = "Set &Manual"
                .uxAttributeMode.Text = "Auto"
                EditorExtension.OverrideAutoAttribute = False
                .uxCurrentlyAttLbl.Visible = False
                .uxCurrentlyAttNum.Visible = False
                EditorExtension.OverrideMapScale = String.Empty
                EditorExtension.OverrideORMapNumber = String.Empty
                EditorExtension.OverrideMapNumber = String.Empty

            Else

                Try
                    .UseWaitCursor = True

                    Dim uxMapnumber As TextBox = .uxMapNumber
                    Dim theMapNumber As String = uxMapnumber.Text.Trim

                    If theMapNumber = String.Empty OrElse Not uxMapnumber.AutoCompleteCustomSource.Contains(theMapNumber) Then
                        .UseWaitCursor = False
                        MessageBox.Show("Invalid MapNumber. Please try again.", "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Exit Sub
                    End If

                    ' Get the mapscale from the MapIndex
                    Dim theQueryFilter As IQueryFilter = New QueryFilter
                    Dim theWhereClause As String = "[" & EditorExtension.MapIndexSettings.MapNumberField & "] = '" & theMapNumber & "'"
                    Dim theMapIndexFeatureClass As IFeatureClass = MapIndexFeatureLayer.FeatureClass
                    theQueryFilter.WhereClause = formatWhereClause(theWhereClause, theMapIndexFeatureClass)
                    Dim theFeatCursor As IFeatureCursor = theMapIndexFeatureClass.Search(theQueryFilter, True)
                    Dim thisFeature As IFeature = theFeatCursor.NextFeature
                    Dim theMapScaleFieldIdx As Integer = thisFeature.Fields.FindField(EditorExtension.MapIndexSettings.MapScaleField)
                    Dim theORMapNumberFieldIdx As Integer = thisFeature.Fields.FindField(EditorExtension.MapIndexSettings.OrmapMapNumberField)
                    EditorExtension.OverrideMapScale = thisFeature.Value(theMapScaleFieldIdx).ToString()
                    EditorExtension.OverrideORMapNumber = thisFeature.Value(theORMapNumberFieldIdx).ToString()
                    EditorExtension.OverrideMapNumber = theMapNumber

                    .uxSetAttributeMode.Text = "Set &Auto"
                    .uxAttributeMode.Text = "Manual Override"
                    EditorExtension.OverrideAutoAttribute = True
                    .uxCurrentlyAttLbl.Visible = True
                    .uxCurrentlyAttNum.Visible = True
                    .uxCurrentlyAttNum.Text = theMapNumber

                Finally
                    .UseWaitCursor = False
                End Try

            End If

        End With

    End Sub


    Private Sub uxTaxlot_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxTaxlot.Enter

        If Not _mapIndexHasBeenChanged Then Exit Sub

        With PartnerLocateFeatureUserControl
            Try
                .UseWaitCursor = True

                Dim uxMapnumber As TextBox = .uxMapNumber
                Dim uxTaxlot As TextBox = .uxTaxlot

                If uxMapnumber.AutoCompleteCustomSource.Contains(uxMapnumber.Text.Trim) Then

                    Dim theMapNumberVal As String = uxMapnumber.Text.Trim
                    If theMapNumberVal = String.Empty Then Exit Sub

                    Dim theTaxlotFClass As IFeatureClass = TaxlotFeatureLayer.FeatureClass

                    Dim theQueryFilter As IQueryFilter = New QueryFilter
                    theQueryFilter.SubFields = EditorExtension.TaxLotSettings.MapNumberField & ", " & EditorExtension.TaxLotSettings.TaxlotField
                    '' NOTE: [SC] The following whereclause was slower than querying the entire recordset and then filtering 
                    '' results (see feature loop below).
                    'Dim theWhereClause As String
                    'theWhereClause = "[" & EditorExtension.TaxLotSettings.MapNumberField & "] = '" & theMapNumberVal & "'"
                    'theQueryFilter.WhereClause = formatWhereClause(theWhereClause, theTaxlotFClass)

                    '' NOTE: [SC] Calculating a AutoCompleteCustomSource using the List of strings is considerably faster than
                    '' adding with the .add method.
                    Dim taxlotStringList As List(Of String) = New List(Of String)
                    Dim theFeatCursor As IFeatureCursor = theTaxlotFClass.Search(theQueryFilter, True)
                    Dim theFeature As IFeature = theFeatCursor.NextFeature
                    If theFeature Is Nothing Then Exit Sub
                    Dim theTaxlotFieldIdx As Integer = theFeature.Fields.FindField(EditorExtension.TaxLotSettings.TaxlotField)
                    Dim theMapNumberFieldIdx As Integer = theFeature.Fields.FindField(EditorExtension.MapIndexSettings.MapNumberField)

                    Do Until theFeature Is Nothing
                        Dim theTaxlotVal As String = theFeature.Value(theTaxlotFieldIdx).ToString
                        If Not uxTaxlot.AutoCompleteCustomSource.Contains(theTaxlotVal) AndAlso theMapNumberVal = theFeature.Value(theMapNumberFieldIdx).ToString Then
                            taxlotStringList.Add(theTaxlotVal)
                        End If
                        theFeature = theFeatCursor.NextFeature
                    Loop

                    uxTaxlot.AutoCompleteCustomSource.AddRange(taxlotStringList.ToArray)
                    _mapIndexHasBeenChanged = False '-- reset 

                End If

            Finally
                .UseWaitCursor = False
            End Try

        End With
    End Sub


    Private Sub uxMapNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.uxMapNumber.TextChanged

        Dim uxMapnumber As TextBox = PartnerLocateFeatureUserControl.uxMapNumber
        Dim uxTaxlot As TextBox = PartnerLocateFeatureUserControl.uxTaxlot

        If Not uxMapnumber.AutoCompleteCustomSource.Contains(uxMapnumber.Text.Trim) Then
            If uxMapnumber.AutoCompleteCustomSource.Contains(uxMapnumber.Text.ToLower.Trim) Then
                uxMapnumber.Text = uxMapnumber.Text.ToLower.Trim
                uxMapnumber.SelectionStart = uxMapnumber.Text.Length
            End If
            If uxMapnumber.AutoCompleteCustomSource.Contains(uxMapnumber.Text.ToUpper.Trim) Then
                uxMapnumber.Text = uxMapnumber.Text.ToUpper.Trim
                uxMapnumber.SelectionStart = uxMapnumber.Text.Length
            End If
        End If

        If uxMapnumber.AutoCompleteCustomSource.Contains(uxMapnumber.Text.Trim) Then
            _mapIndexHasBeenChanged = True
            uxTaxlot.AutoCompleteCustomSource.Clear()
            uxTaxlot.Text = String.Empty
        End If

    End Sub
    Private Sub uxFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.uxFind.Click

        With PartnerLocateFeatureUserControl
            Try
                .UseWaitCursor = True

                Dim uxMapnumber As TextBox = PartnerLocateFeatureUserControl.uxMapNumber
                Dim theMapNumberVal As String = uxMapnumber.Text.Trim

                If theMapNumberVal = String.Empty OrElse Not uxMapnumber.AutoCompleteCustomSource.Contains(theMapNumberVal) Then
                    .UseWaitCursor = False
                    MessageBox.Show("Invalid MapNumber. Please try again.", "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If

                Dim uxTaxlot As TextBox = .uxTaxlot
                Dim theTaxlotVal As String = uxTaxlot.Text.Trim

                If theTaxlotVal <> String.Empty AndAlso Not uxTaxlot.AutoCompleteCustomSource.Contains(theTaxlotVal) Then
                    .UseWaitCursor = False
                    MessageBox.Show("Invalid Taxlot. Please try again.", "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If

                '-- Make sure these are set... turning the light bulb on/off will reset these.
                If MapIndexFeatureLayer Is Nothing Then CheckValidMapIndexDataProperties()
                If TaxlotFeatureLayer Is Nothing Then CheckValidTaxlotDataProperties()

                Dim theQueryFilter As IQueryFilter = New QueryFilter
                Dim theXFlayer As IFeatureLayer = Nothing '-- Set as either the MapIndex or Taxlot Feature Layer.
                Dim theWhereClause As String

                If theTaxlotVal = String.Empty Then
                    '[Looking for just a MapIndex...]
                    theXFlayer = MapIndexFeatureLayer
                    theWhereClause = "[" & EditorExtension.MapIndexSettings.MapNumberField & "] = '" & theMapNumberVal & "'"
                Else
                    '[Looking for a MapIndex and Taxlot...]
                    theXFlayer = TaxlotFeatureLayer
                    theWhereClause = "[" & EditorExtension.TaxLotSettings.MapNumberField & "] = '" & theMapNumberVal & "' AND [" & EditorExtension.TaxLotSettings.TaxlotField & "] = '" & theTaxlotVal & "'"
                End If

                theQueryFilter.WhereClause = formatWhereClause(theWhereClause, theXFlayer.FeatureClass)

                Dim theXFClass As IFeatureClass = theXFlayer.FeatureClass
                Dim theFeatCursor As IFeatureCursor = theXFClass.Search(theQueryFilter, True)
                Dim thisFeature As IFeature = theFeatCursor.NextFeature

                Dim theFeatureSelection As IFeatureSelection = DirectCast(theXFlayer, IFeatureSelection)
                If .uxSelectFeatures.Checked Then theFeatureSelection.Clear()

                If thisFeature Is Nothing Then '-- Must be due to invalid mapindex or taxlot entered into the text boxes.
                    MessageBox.Show("Feature does not exist.", "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                Else
                    If .uxSelectFeatures.Checked Then
                        ' Flag the original selection
                        RefreshDisplaySelection()
                        ' Clear all selections
                        Dim theArcMapDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
                        theArcMapDoc.FocusMap.ClearSelection()
                    End If

                    Dim theEnvelope As IEnvelope = thisFeature.Shape.Envelope
                    Do Until thisFeature Is Nothing
                        theEnvelope.Union(thisFeature.Shape.Envelope)
                        If .uxSelectFeatures.Checked Then
                            SetSelectedFeature(theXFlayer, thisFeature, True, False)
                            'theFeatureSelection.Add(thisFeature)
                            'theFeatureSelection.SelectionChanged()
                        End If
                        thisFeature = theFeatCursor.NextFeature
                    Loop

                    If .uxSelectFeatures.Checked Then
                        ' Flag the new selection
                        RefreshDisplaySelection()
                    End If

                    ZoomToEnvelope(theEnvelope)
                End If
            Finally
                .UseWaitCursor = False
            End Try
        End With

    End Sub

    Private Sub uxSelectFeatures_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.uxHelp.Click

        uxSelectFeaturesChecked = PartnerLocateFeatureUserControl.uxSelectFeatures.Checked

    End Sub
    
#Region "OldCode"

    'Private Sub uxMapNumberButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Calls the DisplayList Form to identify layers for mapnumber definition query
    '    Dim theInputList As New ArrayList
    '    If _theLayerList.Count <> 0 Then
    '        theInputList = _theLayerList
    '    Else
    '        Try
    '            Dim theMxDocument As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
    '            Dim theMap As IMap = theMxDocument.FocusMap
    '            Dim theLayerCount As Integer = theMap.LayerCount()
    '            Dim theAddSting As String
    '            For count As Integer = 0 To theLayerCount - 1
    '                Try
    '                    Dim thisLayer As ILayer = theMap.Layer(count)
    '                    Dim thisFeatureLayer As IFeatureLayer = DirectCast(thisLayer, IFeatureLayer)
    '                    Dim thisFeatureClass As IFeatureClass = thisFeatureLayer.FeatureClass()
    '                    If thisFeatureClass.FeatureType = esriFeatureType.esriFTSimple Or thisFeatureClass.FeatureType = esriFeatureType.esriFTAnnotation Then
    '                        Dim theFields As IFields = thisFeatureClass.Fields
    '                        Dim theFieldName As String = EditorExtension.MapIndexSettings.MapNumberField
    '                        Dim i As Integer = theFields.FindField(theFieldName)
    '                        If i > 0 Then
    '                            Dim theDefQuery As IFeatureLayerDefinition2 = DirectCast(thisLayer, IFeatureLayerDefinition2)
    '                            Dim theCheckDef As String = UCase(Mid(theDefQuery.DefinitionExpression, 1, 9))
    '                            If UCase(theCheckDef) = UCase(EditorExtension.MapIndexSettings.MapNumberField) Then
    '                                theAddSting = "1," & thisFeatureLayer.Name
    '                            Else
    '                                theAddSting = "0," & thisFeatureLayer.Name
    '                            End If
    '                            theInputList.Add(theAddSting)
    '                            MsgBox(theAddSting)
    '                        End If
    '                    End If
    '                Catch ex As Exception
    '                    'MsgBox("Find out how to deal with topo layers")
    '                End Try
    '            Next count
    '        Catch ex As Exception
    '            EditorExtension.ProcessUnhandledException(ex)
    '        End Try
    '    End If

    '    _theDisplayList.InputList(theInputList)
    '    _theDisplayList.DisplayListForm_Load(_application)
    '    _theLayerList = _theDisplayList.GetLayerList()

    '    For i As Integer = 0 To _theLayerList.Count - 1
    '        MsgBox("Returned : " & DirectCast(_theLayerList.Item(i), String))
    '    Next i

    '    If _theDisplayList.GetLayerList.Count <> 0 And PartnerLocateFeatureUserControl.uxMapNumber.Text <> "" Then
    '        PartnerLocateFeatureUserControl.uxDisplay.Enabled = True
    '    Else
    '        PartnerLocateFeatureUserControl.uxDisplay.Enabled = False
    '    End If

    'End Sub
    'Private Sub uxDisplayChecked_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles Definition Query
    '    Try
    '        If PartnerLocateFeatureUserControl.uxDisplay.CheckState = CheckState.Checked Then
    '            _uxDisplayChecked = True
    '        Else
    '            _uxDisplayChecked = False
    '        End If

    '        ApplyDefinitionQuery()

    '    Catch ex As Exception
    '        EditorExtension.ProcessUnhandledException(ex)
    '    End Try

    'End Sub
    'Private Sub ClearDefinitionQuery() 'Clears the definition Queries.  Currently not used.
    '    Dim theMxDocument As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
    '    Dim theMap As IMap = theMxDocument.FocusMap
    '    Dim theMapLayerCount As Integer = theMap.LayerCount()
    '    Dim theLayerListCount As Integer = _theLayerList.Count
    '    For theListCount As Integer = 0 To theLayerListCount - 1
    '        Dim theLayerNameLen As Integer = Len(_theLayerList.Item(theListCount))
    '        Dim theInputString As String = DirectCast(_theLayerList.Item(theListCount), String)
    '        Dim theLayerName As String = Mid(theInputString, 3, theLayerNameLen)
    '        For theMapCount As Integer = 0 To theMapLayerCount - 1
    '            If theMap.Layer(theMapCount).Name = theLayerName Then
    '                Dim thisLayer As ILayer = theMap.Layer(theMapCount)
    '                Dim thisFeatureLayer As IFeatureLayer = DirectCast(thisLayer, IFeatureLayer)
    '                Dim thisFeatureClass As IFeatureClass = thisFeatureLayer.FeatureClass
    '                Dim theFields As IFields = thisFeatureClass.Fields
    '                Dim theFieldName As String = EditorExtension.MapIndexSettings.MapNumberField
    '                Dim theLayerDef As IFeatureLayerDefinition2 = DirectCast(thisLayer, IFeatureLayerDefinition2)
    '                Dim i As Integer = theFields.FindField(theFieldName)
    '                If i > 0 Then
    '                    theLayerDef.DefinitionExpression = ""
    '                End If
    '            End If
    '        Next theMapCount
    '    Next theListCount
    'End Sub
#End Region

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.uxHelp.Click
        ' TODO: [NIS] Could be replaced with new help mechanism.

        Dim theRTFStream As System.IO.Stream = _
           Me.GetType().Assembly.GetManifestResourceStream("OrmapTaxlotEditing.LocateFeature_help.rtf")
        OpenHelp("Locate Feature Help", theRTFStream)


    End Sub
    ''' <summary>
    ''' Calls the ORMAPSettingsForm.
    ''' </summary>
    ''' <remarks>It needs to open on the tab with the list of layers that may participate in the definition query</remarks>
    Private Sub uxORMAPProperties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.uxORMAPProperties.click

        Dim theORMAPSettingsForm As New OrmapSettingsForm
        Dim i As Integer = 5 'Index number of tab with mapnumber settings
        theORMAPSettingsForm.uxSettingsTabs.SelectTab(i)
        theORMAPSettingsForm.ShowDialog()

    End Sub
    ''' <summary>
    ''' Calls the Map Definition Form.  Passes the mapnumber value that is currently in the uxMapnumber.
    ''' It builds the definition query from the input from the form and passes it to the ApplyTheDefinitionQuery.
    ''' </summary>
    ''' <remarks>None</remarks>
    Private Sub uxSetDefinitionQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.uxSetDefinitionQuery.click

        Dim theMapDefinition As New MapDefinition
        theMapDefinition.CallForm(_application, _partnerLocateFeatureUserControl.uxMapNumber.Text)
        If theMapDefinition.ApplyQuery Then
            Dim theDefinitionQuery As String = "[" & EditorExtension.MapIndexSettings.MapNumberField & "] = '" & theMapDefinition.MapNumber & "'"
            ApplyTheDefinitionQuery(theDefinitionQuery)
        End If

    End Sub
    ''' <summary>
    ''' Sets the definition to an empty string.
    ''' </summary>
    ''' <remarks>None</remarks>
    Private Sub uxClearDefinitionQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.uxClearDefinition.click

        Dim theDefinitionQuery As String = ""
        ApplyTheDefinitionQuery(theDefinitionQuery)

    End Sub

#End Region

#Region "Methods"
    ''' <summary>
    ''' This is for development purposes only.
    ''' </summary>
    ''' <remarks>WARNING: Do not put computation-intensive code here.</remarks>
    Friend Function CreateString() As List(Of String) 'Temporary Function.  Needs to be deleted before deployment.
        Dim StringList As New List(Of String)
        StringList.Add("Taxlot")
        StringList.Add("Reference Lines")
        StringList.Add("CartographicLines")
        StringList.Add("Anno 0010 Scale")
        StringList.Add("Anno 0020 Scale")
        StringList.Add("Anno 0030 Scale")
        StringList.Add("Anno 0040 Scale")
        StringList.Add("Anno 0050 Scale")
        StringList.Add("Anno 0100 Scale")
        StringList.Add("Anno 0200 Scale")
        StringList.Add("Anno 2000 Scale")
        StringList.Add("Taxlot Number Anno")
        StringList.Add("Address Grid Anno")
        StringList.Add("Taxlot Acres Anno")
        StringList.Add("Situs Address Anno")
        StringList.Add("Map Index")
        Return StringList
    End Function
    Friend Function CreateScaleString() As List(Of String)
        Dim StringList As New List(Of String)
        StringList.Add("PLSS Corner")
        StringList.Add("PLSS Lines")
        Return StringList
    End Function
    ''' <summary>
    ''' This searches the table of contents for Layers that are identified as participating in the definition query.
    ''' A query sting has to be passed to this routine.
    ''' <param name="QueryString"> The Definition Query String that is applied to specified features layers.</param>
    ''' </summary>
    ''' <remarks>None</remarks>
    Friend Sub ApplyTheDefinitionQuery(ByRef QueryString As String)
        Dim theLayerList As List(Of String) = CreateString()
        If theLayerList.Count > 0 Then
            Dim theMxDocument As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
            Dim theMap As IMap = theMxDocument.FocusMap
            Dim theEnumLayerList As IEnumLayer = theMap.Layers
            theEnumLayerList.Reset()
            Dim theLayer As ILayer = theEnumLayerList.Next
            Do While Not theLayer Is Nothing
                'MsgBox("the current layer is: " & theLayer.Name)
                If theLayerList.Contains(theLayer.Name) Then
                    DefinitionQuery(theLayer, QueryString)
                End If
                theLayer = theEnumLayerList.Next
            Loop
        End If

        Dim theScaleLayerList As List(Of String) = CreateScaleString()
        If theScaleLayerList.Count > 0 Then
            Dim theMxDocument As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
            Dim theMap As IMap = theMxDocument.FocusMap
            Dim theEnumLayerList As IEnumLayer = theMap.Layers
            theEnumLayerList.Reset()
            Dim theLayer As ILayer = theEnumLayerList.Next
            Do While Not theLayer Is Nothing
                'MsgBox("the current layer is: " & theLayer.Name)
                If theScaleLayerList.Contains(theLayer.Name) Then
                    DefinitionQuery(theLayer, QueryString)
                End If
                theLayer = theEnumLayerList.Next
            Loop
        End If

    End Sub
    ''' <summary>
    ''' Called to apply a definition query to a layer.  An Ilayer object and the query needs to be passed to this routine.
    ''' <param name="theLayer" >The Layer that the Definition Query is applies </param>
    ''' <param name="QueryString">The query that is applied to the layer</param>
    ''' </summary>
    ''' <remarks>None</remarks>
    Friend Sub DefinitionQuery(ByRef theLayer As ILayer, ByRef QueryString As String) 'Mover to Editor Extension?
        Try
            If TypeOf theLayer Is IFeatureLayer And Not TypeOf theLayer Is IAnnotationSublayer Then
                'MsgBox(theLayer.Name & " is a feature layer")
                Dim theFeatureLayer As IFeatureLayer = DirectCast(theLayer, IFeatureLayer)
                Dim theFeatureLayerDefinition As IFeatureLayerDefinition = DirectCast(theFeatureLayer, IFeatureLayerDefinition)
                theFeatureLayerDefinition.DefinitionExpression = QueryString
            Else
                If TypeOf theLayer Is IAnnotationSublayer Then
                    'MsgBox(theLayer.Name & " is a annotation sublayer")
                Else
                    'MsgBox(theLayer.Name & " is not a feature layer")
                End If
            End If
        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try


    End Sub
    Friend Sub DoButtonOperation()

        Try
            PartnerLocateFeatureUserControl.uxMapNumber.Enabled = False
            PartnerLocateFeatureUserControl.uxTaxlot.Enabled = False
            PartnerLocateFeatureUserControl.uxFind.Enabled = False

            ' Check for valid data.
            CheckValidMapIndexDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & NewLine & _
                                "Please load this dataset into your map.", _
                                "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            Else
                PartnerLocateFeatureUserControl.uxMapNumber.Enabled = True
                PartnerLocateFeatureUserControl.uxFind.Enabled = True
                PartnerLocateFeatureUserControl.uxClearDefinitionQuery.Enabled = True 'Jon Added
            End If

            CheckValidTaxlotDataProperties()
            If HasValidTaxlotData Then
                PartnerLocateFeatureUserControl.uxTaxlot.Enabled = True
            Else
                PartnerLocateFeatureUserControl.uxTaxlot.Enabled = False
            End If

            _locateFeatureDockWin.Show(Not _locateFeatureDockWin.IsVisible)
            If _locateFeatureDockWin.IsVisible AndAlso PartnerLocateFeatureUserControl.uxMapNumber.AutoCompleteCustomSource.Count = 0 Then PartnerLocateFeatureForm_Load()

            If _locateFeatureDockWin.IsVisible() Then
                PartnerLocateFeatureUserControl.uxTimer.Enabled = True
                PartnerLocateFeatureUserControl.uxEditingGroupBox.Enabled = EditorExtension.AllowedToAutoUpdateAllFields
                PartnerLocateFeatureUserControl.uxMapNumber.Focus()
            Else
                PartnerLocateFeatureUserControl.uxTimer.Enabled = False
            End If

        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)

        End Try

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

        Try
            If Not hook Is Nothing Then

                'Disable tool if parent application is not ArcMap
                If TypeOf hook Is IMxApplication Then
                    _application = DirectCast(hook, IApplication)

                    '-- For testing
                    'setPartnerLocateFeatureForm(New LocateFeatureForm())

                    ' Get a reference to the dockable window
                    _locateFeatureDockWinMgr = DirectCast(hook, IDockableWindowManager)
                    Dim locateFeatureDockWinUID As New UID
                    locateFeatureDockWinUID.Value = "{7c5bb546-215f-477a-8df4-16cc1c993309}"
                    _locateFeatureDockWin = _locateFeatureDockWinMgr.GetDockableWindow(locateFeatureDockWinUID)
                    setPartnerLocateFeatureUserControl(DirectCast(_locateFeatureDockWin.UserData, LocateFeatureUserControl))
                    ' Close this when the application starts...
                    _locateFeatureDockWin.Show(False)

                    MyBase.m_enabled = True
                Else
                    MyBase.m_enabled = False
                End If
            End If

            ' NOTE: Add other initialization code here...

        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try
    End Sub


    Public Overrides ReadOnly Property Checked() As Boolean
        Get
            Checked = _locateFeatureDockWin.IsVisible()
        End Get
    End Property


    Public Overrides Sub OnClick()
        Try
            DoButtonOperation()
        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try
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
    Public Const ClassId As String = "ecace054-c346-4be5-a980-8d81ede68fcb"
    Public Const InterfaceId As String = "effc83d3-a320-460c-9b1e-7f48a4b282d8"
    Public Const EventsId As String = "591a6993-fcb1-4c0c-879f-5488e0c88312"
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
    '''' <summary>
    '''' Required method for ArcGIS Component Category registration -
    '''' Do not modify the contents of this method with the code editor.
    '''' </summary>
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



