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
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports OrmapTaxlotEditing.SpatialUtilities

<ComVisible(True)> _
<ComClass(TaxlotAssignment.ClassId, TaxlotAssignment.InterfaceId, TaxlotAssignment.EventsId), _
ProgId("ORMAPTaxlotEditing.TaxlotAssignment")> _
Public NotInheritable Class TaxlotAssignment
    Inherits BaseTool

#Region "Class-Level Constants And Enumerations"

    ' Taxlot number type constants
    Private Const taxlotNumberTypeTaxlot As String = "TAXLOT" 'normal taxlot number
    Private Const taxlotNumberTypeRoads As String = "ROADS"
    Private Const taxlotNumberTypeWater As String = "WATER"
    Private Const taxlotNumberTypeRails As String = "RAILS"
    Private Const taxlotNumberTypeNontaxlot As String = "NONTL"

    Private Const defaultCommand As String = "esriArcMapUI.SelectTool"

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
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As ArgumentException
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

        Try
            ' Set the (enabled) cursor based on the name of the class.
            Dim cursorResourceName As String = Me.GetType().Name + ".cur"
            MyBase.m_cursor = New Cursor(Me.GetType(), cursorResourceName)
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

#End Region

#Region "Properties"

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

    Public ReadOnly Property NumberStartingFrom() As Integer
        Get
            _numberStartingFrom = CInt(PartnerTaxlotAssignmentForm.uxStartingFrom.Text)
            Return _numberStartingFrom
        End Get
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

    Private Sub SetPartnerTaxlotAssignmentForm(ByVal value As TaxlotAssignmentForm)
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
            If .uxType.SelectedIndex <> -1 Then
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

#Region "Methods (none)"
#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Return MyBase.Enabled AndAlso _
                EditorExtension.Editor IsNot Nothing AndAlso _
                EditorExtension.Editor.EditState = esriEditState.esriStateEditing AndAlso _
                EditorExtension.IsValidWorkspace AndAlso _
                EditorExtension.HasValidLicense AndAlso _
                EditorExtension.AllowedToEditTaxlots
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

        ' TODO: Add other initialization code

    End Sub

    Public Overrides Sub OnClick()

        ' HACK: NIS Don't know why this form is disposed after first use, but 
        ' this check insures it is available again.
        If PartnerTaxlotAssignmentForm.IsDisposed Then
            SetPartnerTaxlotAssignmentForm(New TaxlotAssignmentForm())
        End If

        ' Show and activate the partner form.
        If PartnerTaxlotAssignmentForm.Visible Then
            PartnerTaxlotAssignmentForm.Activate()
        Else
            PartnerTaxlotAssignmentForm.Show()
        End If

    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)

        'TODO: NIS Move the following block to its own method which sets properties of some static object

        ' Initialize document and map objects, and their events for tool reference only
        Dim theDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
        Dim theMap As ESRI.ArcGIS.Carto.IMap
        theDoc = DirectCast(EditorExtension.Application.Document, IMxDocument)
        theMap = theDoc.FocusMap

        ' Find Taxlot and Map Index feature layers
        Dim theTaxlotFLayer As ESRI.ArcGIS.Carto.IFeatureLayer
        'Dim theMapIndexFLayer As ESRI.ArcGIS.Carto.IFeatureLayer

        With EditorExtension.TableNamesSettings

            ' Find Taxlot feature layer
            theTaxlotFLayer = FindFeatureLayerByDSName(.TaxLotFC)
            If theTaxlotFLayer Is Nothing Then
                ' Attempt to load and find the taxlot layer in the map document
                If LoadFCIntoMap(.MapIndexFC) Then
                    theTaxlotFLayer = FindFeatureLayerByDSName(.MapIndexFC)
                Else
                    ' TODO: NIS
                End If

                ' Exit if unable to load the Map Index layer
                If theTaxlotFLayer Is Nothing Then
                    ' TODO: NIS
                End If
            End If

            '' Find MapIndex feature layer
            'theMapIndexFLayer = FindFeatureLayerByDSName(.MapIndexFC)
            'If theMapIndexFLayer Is Nothing Then
            '    ' Attempt to load and find the Map Index layer in the map document
            '    If LoadFCIntoMap(.MapIndexFC) Then
            '        theMapIndexFLayer = FindFeatureLayerByDSName(.MapIndexFC)
            '    Else
            '        ' TODO: NIS
            '    End If

            '    ' Exit if unable to load the Map Index layer
            '    If theMapIndexFLayer Is Nothing Then
            '        ' TODO: NIS
            '    End If
            'End If

        End With

        ' Obtain references to feature layer feature classes
        Dim theTaxlotFClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
        theTaxlotFClass = theTaxlotFLayer.FeatureClass
        'Dim theMapIndexFClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
        'theMapIndexFClass = theMapIndexFLayer.FeatureClass

        ' Get field indexes
        Dim theOMTLNumFld As Integer
        Dim theOMNumFld As Integer
        Dim theOMMapTaxlotFld As Integer
        Dim theTLTaxlotFld As Integer
        Dim theTLAnomalyFld As Integer

        With EditorExtension.TaxLotSettings

            Const fieldNotFoundIndex As Integer = -1

            ' Find the ORMAP Taxlot field index
            theOMTLNumFld = theTaxlotFClass.FindField(.OrmapTaxlotField)
            If theOMTLNumFld = fieldNotFoundIndex Then
                ' TODO: NIS
            End If

            ' Find the ORMAP Map Number field index
            theOMNumFld = theTaxlotFClass.FindField(.OrmapMapNumberField)
            If theOMNumFld = fieldNotFoundIndex Then
                ' TODO: NIS
            End If

            ' Find the ORMAP Map Taxlot field index
            theOMMapTaxlotFld = theTaxlotFClass.FindField(.MapTaxlotField)
            If theOMMapTaxlotFld = fieldNotFoundIndex Then
                ' TODO: NIS
            End If

            ' Find the Taxlot field index
            theTLTaxlotFld = theTaxlotFClass.FindField(.TaxlotField)
            If theTLTaxlotFld = fieldNotFoundIndex Then
                ' TODO: NIS
            End If

            ' Find the Anomaly field index
            theTLAnomalyFld = theTaxlotFClass.FindField(.AnomalyField)
            If theTLAnomalyFld = fieldNotFoundIndex Then
                ' TODO: NIS
            End If

        End With

        'Dim theEditStatus As Boolean
        'Dim theToolStatus As Boolean


        'TODO: NIS Port TaxlotAssignment.OnMouseDown implementation

        'On Error GoTo ErrorHandler

        'Dim theActiveView As ESRI.ArcGIS.Carto.IActiveView
        'Dim theCmdItem As ESRI.ArcGIS.Framework.ICommandItem
        'Dim theFeature As ESRI.ArcGIS.Geodatabase.IFeature
        'Dim theFCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
        'Dim theSpatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter
        'Dim theEnv As ESRI.ArcGIS.Geometry.IEnvelope
        'Dim theGeometry As ESRI.ArcGIS.Geometry.IGeometry
        'Dim thePoint As ESRI.ArcGIS.Geometry.IPoint
        'Dim theUID As New ESRI.ArcGIS.esriSystem.UID

        'Dim iResponse As Short
        'Dim lIncrement As Integer
        'Dim lValue As Integer
        'Dim sAnom As String
        'Dim sExistOMNum As String
        'Dim sExistTLNumVal As String
        'Dim sNewTLNum As String
        'Dim sNewTLNum_5digit As String
        'Dim sOMMapTaxlot As String
        'Dim sOMTLNval As String
        'Dim sShapeFieldName As String
        'Dim sShortOMNum As String
        'Dim sTLMapSufNumVal As String
        'Dim sTLMapSufTypeVal As String

        'If Button <> MouseButtons.Left Then  ' TODO: confirm this works
        '    Exit Sub
        'End If

        '' Insure that it is the tool has the proper data to continue
        'If Not EditorExtension.AllowedToEditTaxlots Then
        '    Exit Sub
        'End If

        ''If "NUMBER" selected, then make sure value is numeric
        'If StrComp(m_pFrmTaxlots.PolygonType, "NUMBER", CompareMethod.Text) = 0 Then
        '    If IsNumeric(m_pFrmTaxlots.CurrentValue) Then
        '        lValue = m_pFrmTaxlots.CurrentValue
        '        lIncrement = m_pFrmTaxlots.Increment
        '    Else
        '        Exit Sub
        '    End If
        'End If

        '' Create a search shape out of the point that the user clicked
        ''UPGRADE_WARNING: Couldn't resolve default property of object m_pEditor.Map. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pMap = m_pEditor.Map
        ''UPGRADE_WARNING: Couldn't resolve default property of object m_pEditor.Display. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pPoint = m_pEditor.Display.DisplayTransformation.ToMapPoint(X, Y)
        ''Set pGeometry = m_pEditor.CreateSearchShape(pPoint) 'Returns an IEnvelope
        'pGeometry = pPoint 'QI

        '' Verify the validity of the specified taxlot number, and uniqueness of
        '' numeric taxlot numbers
        'If StrComp(m_pFrmTaxlots.PolygonType, "NUMBER", CompareMethod.Text) = 0 Then
        '    If Not ValidateTaxlotNum(CStr(m_pFrmTaxlots.CurrentValue), pGeometry) Then
        '        If MsgBox("The current Taxlot value (" & m_pFrmTaxlots.CurrentValue & ") is not unique within this MapIndex.  " & vbCrLf & "Attribute feature with value anyways?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.No Then
        '            GoTo Process_Exit
        '        End If
        '    End If
        'End If

        '' Insure the validity of the underlying map index polygon
        ''Set pEnv = pGeometry 'QI
        'pSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
        'pSpatialFilter.Geometry = pGeometry
        'sShapeFieldName = m_pMIFclass.ShapeFieldName
        ''UPGRADE_WARNING: Couldn't resolve default property of object pSpatialFilter.OutputSpatialReference. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pSpatialFilter.OutputSpatialReference(sShapeFieldName) = pMap.SpatialReference
        'pSpatialFilter.GeometryField = m_pMIFclass.ShapeFieldName
        'pSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
        ''UPGRADE_WARNING: Couldn't resolve default property of object pSpatialFilter. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pFCursor = m_pMIFclass.Search(pSpatialFilter, False)
        'pFeature = pFCursor.NextFeature
        'If pFeature Is Nothing Then
        '    MsgBox("Unable to assign taxlot values to polygons that are not within a Map Index polygon", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
        '    GoTo Process_Exit
        'End If

        '' Select any feature under the given point in the target layer
        ''Set pEnv = pGeometry 'QI
        'pSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
        'pSpatialFilter.Geometry = pGeometry
        'sShapeFieldName = m_pTaxlotFClass.ShapeFieldName
        ''UPGRADE_WARNING: Couldn't resolve default property of object pSpatialFilter.OutputSpatialReference. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pSpatialFilter.OutputSpatialReference(sShapeFieldName) = pMap.SpatialReference
        'pSpatialFilter.GeometryField = m_pTaxlotFClass.ShapeFieldName
        'pSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
        ''UPGRADE_WARNING: Couldn't resolve default property of object pSpatialFilter. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pFCursor = m_pTaxlotFClass.Search(pSpatialFilter, False)

        '' Exit if no taxlots are selected
        'If pFCursor Is Nothing Then
        '    GoTo Process_Exit
        'End If

        'pFeature = pFCursor.NextFeature
        'If Not pFeature Is Nothing Then
        '    'Update the feature
        '    'UPGRADE_WARNING: Couldn't resolve default property of object m_pEditor.StartOperation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    m_pEditor.StartOperation()

        '    ' The current ORMAP Number
        '    'UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        '    sExistOMNum = IIf(IsDBNull(pFeature.Value(m_lOMNumFld)), "", pFeature.Value(m_lOMNumFld))

        '    ' Obtain the ORMAP Number from a MapIndex polygon if it is not present
        '    If Len(sExistOMNum) = 0 Then
        '        ' This call will point m_pMIFlayer to the MapIndex feature class
        '        'UPGRADE_WARNING: Couldn't resolve default property of object m_pMIFlayer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        CalcTaxlotValues(pFeature, m_pMIFlayer)

        '        ' Refresh the ORMAP Number
        '        'UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        '        sExistOMNum = IIf(IsDBNull(pFeature.Value(m_lOMNumFld)), "", pFeature.Value(m_lOMNumFld))

        '        ' Stop if there is still no ORMAP number
        '        If Len(sExistOMNum) = 0 Then
        '            '++ START JWalton 2/1/2007 Changed user message
        '            MsgBox("ORMAPMapNumber not present in this taxlot or MapIndex" & vbCrLf & "Use the MapIndex tool to populate the ORMAPMapNumber field before using this tool", MsgBoxStyle.OkOnly, "Create Tax Number")
        '            GoTo Process_Exit
        '        End If
        '    End If
        'End If

        ''Assign Taxlot value
        ''UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        'sExistTLNumVal = IIf(IsDBNull(pFeature.Value(m_lTLTaxlotFld)), "", pFeature.Value(m_lTLTaxlotFld))

        '' Optionally, update the taxlot number field
        'If Len(sExistTLNumVal) > 0 And sExistTLNumVal <> "0" Then
        '    iResponse = MsgBox("Taxlot currently has a Taxlot value (" & sExistTLNumVal & ").  Update it?", MsgBoxStyle.YesNo)
        '    If iResponse = MsgBoxResult.No Then GoTo Process_Exit
        'End If

        ''Taxlot can be less than 5-digits
        ''The Taxlot value in OrMapMapNum must be 5 digits.
        ''Two versions of the taxlot number will be used for these purposes.
        'If StrComp(m_pFrmTaxlots.PolygonType, "NUMBER", CompareMethod.Text) = 0 Then
        '    sNewTLNum = CStr(m_pFrmTaxlots.CurrentValue) 'User entered number
        '    sNewTLNum_5digit = sNewTLNum
        '    ' Make sure number is 5 characters
        '    If Len(sNewTLNum_5digit) < ORMAP_TAXLOT_FIELD_LENGTH Then
        '        Do Until Len(sNewTLNum_5digit) = ORMAP_TAXLOT_FIELD_LENGTH
        '            sNewTLNum_5digit = "0" & sNewTLNum_5digit
        '        Loop
        '    End If
        'Else
        '    'Remove leading Zeros for taxlot number if any exist
        '    sNewTLNum_5digit = m_pFrmTaxlots.PolygonType 'Predefined selection
        '    sNewTLNum = Replace(sNewTLNum_5digit, "0", "")
        'End If

        ''Determine if Special Interests field is something other than default
        ''If so, include it in ORMAPtaxlot
        'sTLMapSufTypeVal = GetMapSufType(pFeature)
        'sTLMapSufNumVal = GetMapSufNum(pFeature)
        ''UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        'If IsDBNull(sTLMapSufTypeVal) Then sTLMapSufTypeVal = "0"
        ''UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        'If IsDBNull(sTLMapSufNumVal) Then sTLMapSufNumVal = "000"
        ''UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pFeature.Value(m_lTLTaxlotFld) = sNewTLNum

        'sShortOMNum = ShortenOMMapNum(sExistOMNum)
        'sOMTLNval = sShortOMNum & sTLMapSufTypeVal & sTLMapSufNumVal & sNewTLNum_5digit
        ''Create  masked value from a combination of ORMapNum and the new taxlot
        ''Note: Special code for Lane County (see comment below).
        'Dim iCountyCode As Short
        'iCountyCode = CShort(g_pFldnames.DefCounty)
        'Select Case iCountyCode
        '    Case 1 To 19, 21 To 36
        '        sOMMapTaxlot = gfn_s_CreateMapTaxlotValue(sExistOMNum & sNewTLNum_5digit, (g_pFldnames.MapTaxlotFormatString))
        '    Case 20
        '        ' 1.  Lane County uses a 2-digit numeric identifier for ranges.
        '        '     Special handling is required for east ranges, where 02E is
        '        '     stored as 25, 03E as 35, etc.
        '        ' 2.  ORMAP standards (OCDES (pg 13); Taxmap Data Model (pg 11)) assert that
        '        '     this field should be equal to MAPNUMBER + TAXLOT. In this case, MAPNUMBER
        '        '     is already in the right format, thus removing the need for the
        '        '     gfn_s_CreateMapTaxlotValue function. Also, in this case, TAXLOT is padded
        '        '     on the left with zeros to make it always a 5-digit number (see comment
        '        '     above).
        '        sOMMapTaxlot = Trim(Left(sExistOMNum, 8)) & sNewTLNum_5digit
        'End Select
        ''UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pFeature.Value(m_lOMMapTaxlotFld) = sOMMapTaxlot

        ''Copy Anomaly from MapIndex
        'sAnom = ParseOMMapNum(sExistOMNum, "anomaly")
        ''UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pFeature.Value(m_lTLAnomalyFld) = sAnom

        ''Assign OrmapTaxlot value
        ''UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pFeature.Value(m_lOMTLNumFld) = sOMTLNval
        ''UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Store. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pFeature.Store()

        ''AutoIncrement if necessary
        'If m_pFrmTaxlots.Increment > 0 And StrComp(m_pFrmTaxlots.PolygonType, "NUMBER", CompareMethod.Text) = 0 Then
        '    m_pFrmTaxlots.CurrentValue = lValue + lIncrement
        'End If

        '' End the edit operation
        ''UPGRADE_WARNING: Couldn't resolve default property of object m_pEditor.StopOperation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'm_pEditor.StopOperation("AutoIncrement Attribute")

        '' Select the feature
        'pMap.ClearSelection()
        ''UPGRADE_WARNING: Couldn't resolve default property of object m_pTaxlotFLayer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'pMap.SelectFeature(m_pTaxlotFLayer, pFeature)

        '' Insure that this tool keeps the focus
        'm_bToolStatus = True
        'pUID.Value = "TaxlotEditing.cmdTaxlotAssignment"
        'pDoc = m_pDoc
        'pCmdItem = pDoc.CommandBars.Find(pUID, True, False)
        'g_pApp.CurrentTool = pCmdItem
        'm_bToolStatus = False

        '' Update the view
        'pActiveView = pMap

        ''Partially refresh the display
        'pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewBackground Or ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeography Or ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGraphics Or ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGraphicSelection, Nothing, Nothing)


        'Process_Exit:
        '        Exit Sub

        'ErrorHandler:
        '        ' Abort any ongoing edit operations
        '        'UPGRADE_WARNING: Couldn't resolve default property of object m_pEditor.AbortOperation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        m_pEditor.AbortOperation()

        '        ' Handle the error
        '        ' TODO: NIS

    End Sub

#End Region

#End Region

#Region "Implemented Interface Members"
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



