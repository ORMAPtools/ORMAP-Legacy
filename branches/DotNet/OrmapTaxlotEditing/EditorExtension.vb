#Region "Copyright 2008 ORMAP Tech Group"

' File:  EditorExtension.vb
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
Imports System.Collections.Generic
Imports system.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports OrmapTaxlotEditing.DataMonitor
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.StringUtilities
Imports OrmapTaxlotEditing.Utilities
#End Region

<ComVisible(True)> _
<ComClass(EditorExtension.ClassId, EditorExtension.InterfaceId, EditorExtension.EventsId), _
ProgId("ORMAPTaxlotEditing.EditorExtension")> _
Public NotInheritable Class EditorExtension
    Implements IExtension
    Implements IExtensionAccelerators
    Implements IPersistVariant

#Region "Class-Level Constants And Enumerations (none)"
#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Properties"

    Private Shared _application As IApplication

    Friend Shared ReadOnly Property Application() As IApplication
        Get
            Return _application
        End Get
    End Property

    Private Sub setApplication(ByVal value As IApplication)
        _application = value
    End Sub

    Private Shared _editor As IEditor2

    Friend Shared ReadOnly Property Editor() As IEditor2
        Get
            Return _editor
        End Get
    End Property

    Private Sub setEditor(ByVal value As IEditor2)
        _editor = value
    End Sub

    Private Shared _editEvents As IEditEvents_Event

    Friend Shared ReadOnly Property EditEvents() As IEditEvents_Event
        Get
            Return _editEvents
        End Get
    End Property

    Private Sub setEditEvents(ByVal value As IEditEvents_Event)
        _editEvents = value
    End Sub

    Private Shared _activeViewEvents As IActiveViewEvents_Event

    Friend Shared ReadOnly Property ActiveViewEvents() As IActiveViewEvents_Event
        Get
            Return _activeViewEvents
        End Get
    End Property

    Private Sub setActiveViewEvents(ByVal value As IActiveViewEvents_Event)
        _activeViewEvents = value
    End Sub

    Friend Shared ReadOnly Property TableNamesSettings() As TableNamesSettings
        Get
            Return New TableNamesSettings
        End Get
    End Property
    Friend Shared ReadOnly Property AnnoTableNamesSettings() As AnnoTableNamesSettings
        Get
            Return New AnnoTableNamesSettings
        End Get
    End Property
    Friend Shared ReadOnly Property AllTablesSettings() As AllTablesSettings
        Get
            Return New AllTablesSettings
        End Get
    End Property
    Friend Shared ReadOnly Property MapIndexSettings() As MapIndexSettings
        Get
            Return New MapIndexSettings
        End Get
    End Property
    Friend Shared ReadOnly Property TaxLotSettings() As TaxLotSettings
        Get
            Return New TaxLotSettings
        End Get
    End Property
    Friend Shared ReadOnly Property TaxLotLinesSettings() As TaxLotLinesSettings
        Get
            Return New TaxLotLinesSettings
        End Get
    End Property
    Friend Shared ReadOnly Property CartographicLinesSettings() As CartographicLinesSettings
        Get
            Return New CartographicLinesSettings
        End Get
    End Property
    Friend Shared ReadOnly Property TaxlotAcreageAnnoSettings() As TaxlotAcreageAnnoSettings
        Get
            Return New TaxlotAcreageAnnoSettings
        End Get
    End Property
    Friend Shared ReadOnly Property TaxlotNumberAnnoSettings() As TaxlotNumberAnnoSettings
        Get
            Return New TaxlotNumberAnnoSettings
        End Get
    End Property
    Friend Shared ReadOnly Property DefaultValuesSettings() As DefaultValuesSettings
        Get
            Return New DefaultValuesSettings
        End Get
    End Property

    Friend Shared ReadOnly Property CanEnableExtendedEditing() As Boolean
        Get
            Dim canEnable As Boolean = True
            canEnable = canEnable AndAlso EditorExtension.Editor IsNot Nothing
            canEnable = canEnable AndAlso EditorExtension.AllowedToEditTaxlots
            canEnable = canEnable AndAlso EditorExtension.HasValidLicense
            canEnable = canEnable AndAlso EditorExtension.IsValidWorkspace
            Return canEnable
        End Get
    End Property

    Private Shared _hasValidLicense As Boolean '= False

    Friend Shared ReadOnly Property HasValidLicense() As Boolean
        Get
            Return _hasValidLicense
        End Get
    End Property

    Private Sub setHasValidLicense(ByVal value As Boolean)
        _hasValidLicense = value
    End Sub

    Private Shared _isValidWorkspace As Boolean '= False

    Friend Shared ReadOnly Property IsValidWorkspace() As Boolean
        Get
            Return _isValidWorkspace
        End Get
    End Property

    Private Sub setIsValidWorkspace(ByVal value As Boolean)
        _isValidWorkspace = value
    End Sub

    Private Shared _allowedToEditTaxlots As Boolean = True

    Friend Shared Property AllowedToEditTaxlots() As Boolean
        Get
            Return _allowedToEditTaxlots
        End Get
        Set(ByVal value As Boolean)
            _allowedToEditTaxlots = value
        End Set
    End Property

    Private Shared _allowedToAutoUpdate As Boolean = True

    Friend Shared Property AllowedToAutoUpdate() As Boolean
        Get
            Return _allowedToAutoUpdate
        End Get
        Set(ByVal value As Boolean)
            _allowedToAutoUpdate = value
        End Set
    End Property

    Private Shared _allowedToAutoUpdateAllFields As Boolean = True

    Friend Shared Property AllowedToAutoUpdateAllFields() As Boolean
        Get
            Return _allowedToAutoUpdateAllFields
        End Get
        Set(ByVal value As Boolean)
            _allowedToAutoUpdateAllFields = value
        End Set
    End Property

#End Region

#Region "Fields"

    Private _duringAutoUpdate As Boolean = False

#End Region

#Region "Event Handlers"

#Region "Editor Event Handlers"

    ''' <summary>
    ''' Updates fields based on the feature that was just changed.
    ''' </summary>
    ''' <param name="obj">The feature that was just changed.</param>
    ''' <remarks>Handles EditEvents.OnCreateFeature events.</remarks>
    Private Sub EditEvents_OnChangeFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) 'Handles EditEvents.OnChangeFeature

        Try
            If Not EditorExtension.CanEnableExtendedEditing Then Exit Try
            If Not EditorExtension.AllowedToAutoUpdate Then Exit Try
            If Not IsORMAPFeature(obj) Then Exit Try

            ' Update the minimum auto-calculated fields
            UpdateMinimumAutoFields(DirectCast(obj, IFeature))

            If Not EditorExtension.AllowedToAutoUpdateAllFields Then Exit Try

            ' Avoid rentrancy
            If _duringAutoUpdate = False Then
                _duringAutoUpdate = True
            Else
                Throw New InvalidOperationException("Already in AutoUpdate mode. Cannot initiate AutoUpdate.")
                Exit Try
            End If

            ' Note: Must check here for if required data is available
            ' (in case subroutines called don't check).

            ' Check for valid data (will try to load data if not found).
            CheckValidDataProperties()
            If Not HasValidTaxlotData Then
                MessageBox.Show("Unable to update Taxlot field values." & vbNewLine & _
                                "Missing data: Valid ORMAP Taxlot layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnChangeFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If
            If Not HasValidMapIndexData Then
                MessageBox.Show("Unable to update taxlot field values." & vbNewLine & _
                                "Missing data: Valid ORMAP MapIndex layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnChangeFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)

                Exit Try
            End If
            
            ' Variable declarations
            Dim theFeature As ESRI.ArcGIS.Geodatabase.IFeature
            Dim theAnnotationFeature As ESRI.ArcGIS.Carto.IAnnotationFeature

            Dim theParentID As Integer
            If IsTaxlot(obj) Then
                ' Obtain OrmapMapNumber via overlay and calculate other field values.
                CalculateTaxlotValues(DirectCast(obj, IFeature), FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.MapIndexFC))
            ElseIf IsAnno(obj) Then
                theAnnotationFeature = DirectCast(obj, IAnnotationFeature)

                'Get the parent feature so mapnumber can be obtained
                theParentID = theAnnotationFeature.LinkedFeatureID
                If theParentID > -1 Then 'Feature linked
                    theFeature = GetRelatedObjects(obj)
                    If theFeature Is Nothing Then Exit Try
                Else
                    'Not feature linked anno, so we can use the feature as is
                    theFeature = DirectCast(obj, IFeature)
                End If

                'Set anno size
                SetAnnoSize(obj, theFeature)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        Finally
            _duringAutoUpdate = False

        End Try

    End Sub

    ''' <summary>
    ''' Updates fields based on the feature that was just created.
    ''' </summary>
    ''' <param name="obj">The feature that was just created.</param>
    ''' <remarks>Handles EditEvents.OnCreateFeature events.</remarks>
    Private Sub EditEvents_OnCreateFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) 'Handles EditEvents.OnCreateFeature

        ' TODO: [NIS] Add code to check for if required data is available (see OnDeleteFeature).

        Try
            If Not EditorExtension.CanEnableExtendedEditing Then Exit Try
            If Not EditorExtension.AllowedToAutoUpdate Then Exit Try
            If Not IsORMAPFeature(obj) Then Exit Try

            ' Update the minimum auto-calculated fields
            UpdateMinimumAutoFields(DirectCast(obj, IFeature))

            If Not EditorExtension.AllowedToAutoUpdateAllFields Then Exit Try

            ' Note: Must check here for if required data is available
            ' (in case subroutines called don't check).

            ' Check for valid data (will try to load data if not found).
            CheckValidDataProperties()
            If Not HasValidTaxlotData Then
                MessageBox.Show("Unable to populate Taxlot field values." & vbNewLine & _
                                "Missing data: Valid ORMAP Taxlot layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnCreateFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If
            If Not HasValidMapIndexData Then
                MessageBox.Show("Unable to populate taxlot field values." & vbNewLine & _
                                "Missing data: Valid ORMAP MapIndex layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnCreateFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If

            ' TODO: [NIS] Move this code to within the below logic

            Dim theMapNumberFldIdx As Integer
            Dim theMapScaleFldIdx As Integer
            ' TODO: [NIS] Test this DirectCast.
            theMapNumberFldIdx = (DirectCast(obj, IFeature)).Fields.FindField(EditorExtension.MapIndexSettings.MapNumberField)
            ' TODO: [NIS] Test this DirectCast.
            theMapScaleFldIdx = (DirectCast(obj, IFeature)).Fields.FindField(EditorExtension.MapIndexSettings.MapScaleField)

            Dim theFeature As ESRI.ArcGIS.Geodatabase.IFeature
            Dim theGeometry As ESRI.ArcGIS.Geometry.IGeometry
            Dim theEnvelope As ESRI.ArcGIS.Geometry.IEnvelope
            Dim theCenterPoint As ESRI.ArcGIS.Geometry.IPoint
            Dim theMapScaleVal As String
            Dim theMapNumberVal As String

            If IsTaxlot(obj) Then
                '[Edited object is a ORMAP taxlot feature...]

                ' Obtain OrmapMapNumber via overlay and calculate other field values.
                theFeature = DirectCast(obj, IFeature)
                CalculateTaxlotValues(theFeature, MapIndexFeatureLayer)

            ElseIf IsAnno(obj) Then
                '[Edited object is an ORMAP annotation feature...]

                Dim theAnnotationFeature As ESRI.ArcGIS.Carto.IAnnotationFeature
                theAnnotationFeature = DirectCast(obj, IAnnotationFeature)

                'Get the parent feature so mapnumber can be obtained
                Dim theParentID As Integer
                theParentID = theAnnotationFeature.LinkedFeatureID
                If theParentID > FieldNotFoundIndex Then 'Feature linked
                    theFeature = GetRelatedObjects(obj)
                    If theFeature Is Nothing Then Exit Try
                Else
                    'Not feature linked anno, so we can use the feature as is
                    theFeature = DirectCast(obj, IFeature)
                End If

                ' Retrieve the map number and scale from the overlaying map index polygon
                theGeometry = theFeature.Shape
                If theGeometry.IsEmpty Then Exit Try
                theEnvelope = theGeometry.Envelope
                theCenterPoint = DirectCast(theEnvelope, ESRI.ArcGIS.Geometry.IPoint)
                theGeometry = theCenterPoint 'QI

                ' Capture MapNumber for each anno feature created
                ' TODO: [NIS] Test use of pFeat instead of obj here.
                Dim theAnnoMapNumFldIdx As Integer
                theAnnoMapNumFldIdx = theFeature.Fields.FindField(EditorExtension.MapIndexSettings.MapNumberField)
                If theAnnoMapNumFldIdx = FieldNotFoundIndex Then Exit Try

                theMapNumberVal = GetValueViaOverlay(theGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapNumberField, EditorExtension.MapIndexSettings.MapNumberField)
                ' TODO: [NIS] Test use of pFeat instead of obj here.
                theFeature.Value(theAnnoMapNumFldIdx) = theMapNumberVal
                If theMapScaleFldIdx > FieldNotFoundIndex Then
                    theMapScaleVal = GetValueViaOverlay(theGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapScaleField, EditorExtension.MapIndexSettings.MapNumberField)
                    If Len(theMapScaleVal) > 0 Then
                        ' TODO: [NIS] Test use of pFeat instead of obj here.
                        theFeature.Value(theMapScaleFldIdx) = theMapScaleVal
                    Else
                        ' TODO: [NIS] Test use of pFeat instead of obj here.
                        theFeature.Value(theMapScaleFldIdx) = System.DBNull.Value
                    End If
                End If
                ' Set size based on mapscale
                SetAnnoSize(obj, theFeature)
            Else
                '[Edited object is another kind of ORMAP feature (not taxlot or annotation)...]

                ' Update MapScale and mapnumber for all features with a MapScale field (except MapIndex)
                If theMapScaleFldIdx > FieldNotFoundIndex And Not IsMapIndex(obj) Then
                    theFeature = CType(obj, IFeature)
                    theGeometry = theFeature.Shape
                    If theGeometry.IsEmpty Then Exit Try
                    theEnvelope = theGeometry.Envelope
                    If theGeometry.GeometryType <> ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryBezier3Curve And _
                            theGeometry.GeometryType <> ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryCircularArc And _
                            theGeometry.GeometryType <> ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryEllipticArc And _
                            theGeometry.GeometryType <> ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryLine And _
                            theGeometry.GeometryType <> ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPath And _
                            theGeometry.GeometryType <> ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon And _
                            theGeometry.GeometryType <> ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then
                        ' Convert the geometry to a point
                        theCenterPoint = GetCenterOfEnvelope(theEnvelope)
                        theGeometry = theCenterPoint 'QI
                    Else
                        ' Use the geometry as-is
                    End If

                    theMapScaleVal = GetValueViaOverlay(theGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapScaleField, EditorExtension.MapIndexSettings.MapNumberField)
                    If Len(theMapScaleVal) > 0 Then
                        ' TODO: [NIS] Test use of pFeat instead of obj here.
                        theFeature.Value(theMapScaleFldIdx) = theMapScaleVal
                    Else
                        ' TODO: [NIS] Test use of pFeat instead of obj here.
                        theFeature.Value(theMapScaleFldIdx) = System.DBNull.Value
                    End If
                    ' If a dataset with MapNumber, populate it
                    If theMapNumberFldIdx > FieldNotFoundIndex Then
                        theMapNumberVal = GetValueViaOverlay(theGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapNumberField, EditorExtension.MapIndexSettings.MapNumberField)
                        If Len(theMapNumberVal) > 0 Then
                            ' TODO: [NIS] Test use of pFeat instead of obj here.
                            theFeature.Value(theMapNumberFldIdx) = theMapNumberVal
                        Else
                            ' TODO: [NIS] Test use of pFeat instead of obj here.
                            theFeature.Value(theMapNumberFldIdx) = System.DBNull.Value
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Sub

    ''' <summary>
    ''' Records in the Cancelled Numbers object class the map number
    ''' and taxlot number from the feature that was just deleted.
    ''' </summary>
    ''' <param name="obj">The feature that was just deleted.</param>
    ''' <remarks>Handles EditEvents.OnDeleteFeature events.</remarks>
    Private Sub theEditorEvents_OnDeleteFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) 'Handles EditEvents.OnDeleteFeature

        Try
            If Not EditorExtension.CanEnableExtendedEditing Then Exit Try
            If Not EditorExtension.AllowedToAutoUpdate Then Exit Try
            If Not IsORMAPFeature(obj) Then Exit Try ' TODO: [NIS] Is this even needed here?
            If Not EditorExtension.AllowedToAutoUpdateAllFields Then Exit Try

            ' Note: Must check here for if required data is available
            ' (in case subroutines called don't check).

            ' Check for valid data (will try to load data if not found).
            CheckValidDataProperties()
            If Not HasValidTaxlotData Then
                MessageBox.Show("Missing data: Valid ORMAP Taxlot layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnDeleteFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If
            If Not HasValidMapIndexData Then
                MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnDeleteFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If
            If Not HasValidCancelledNumbersTableData Then
                MessageBox.Show("Missing data: Valid ORMAP CancelledNumbersTable not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnDeleteFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If

            If IsTaxlot(obj) Then
                '[Deleting taxlots...]

                ' Capture the mapnumber and taxlot and record them in CancelledNumbers.

                Dim theFeature As ESRI.ArcGIS.Geodatabase.IFeature
                Dim theTaxlotFClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
                Dim theDataSet As ESRI.ArcGIS.Geodatabase.IDataset
                Dim theWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
                Dim theFeatureWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace

                ' Get reference to the Cancelled Numbers object table.
                Dim theRow As ESRI.ArcGIS.Geodatabase.IRow
                theFeature = DirectCast(obj, IFeature)
                theTaxlotFClass = DirectCast(theFeature.Class, IFeatureClass)
                theDataSet = DirectCast(theTaxlotFClass, IDataset)
                theWorkspace = theDataSet.Workspace
                theFeatureWorkspace = DirectCast(theWorkspace, IFeatureWorkspace)

                ' Attempt to get a reference to the Cancelled Number table.
                
                ' Retrieve field positions.
                Dim theTLTaxlotFldIdx As Integer = theTaxlotFClass.FindField(EditorExtension.TaxLotSettings.TaxlotField)
                Dim theTLMapNumberFldIdx As Integer = theTaxlotFClass.FindField(EditorExtension.TaxLotSettings.MapNumberField)
                Dim theCNTaxlotFldIdx As Integer = CancelledNumbersTable.Table.FindField(EditorExtension.TaxLotSettings.TaxlotField)
                Dim theCNMapNumberFldIdx As Integer = CancelledNumbersTable.Table.FindField(EditorExtension.TaxLotSettings.MapNumberField)

                ' If no null values, copy them to Cancelled numbers
                If Not IsDBNull(theFeature.Value(theTLTaxlotFldIdx)) And Not IsDBNull(theFeature.Value(theTLMapNumberFldIdx)) Then
                    theRow = CancelledNumbersTable.Table.CreateRow
                    theRow.Value(theCNTaxlotFldIdx) = theFeature.Value(theTLTaxlotFldIdx)
                    theRow.Value(theCNMapNumberFldIdx) = theFeature.Value(theTLMapNumberFldIdx)
                    theRow.Store()
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Sub

    Private Sub EditEvents_OnStartEditing()

        Try
            ' Test for a valid ArcGIS license.
            setHasValidLicense((validateLicense(esriLicenseProductCode.esriLicenseProductCodeArcEditor) OrElse _
                            validateLicense(esriLicenseProductCode.esriLicenseProductCodeArcInfo)))

            ' Test for a valid workspace.
            If EditorExtension.Editor.EditWorkspace.Type = esriWorkspaceType.esriFileSystemWorkspace Then
                setIsValidWorkspace(False)
            Else
                setIsValidWorkspace(True)
            End If

            If HasValidLicense AndAlso IsValidWorkspace Then

                ' Set the Application property
                setApplication(DirectCast(Editor.Parent, IApplication))

                ' Set active view events object
                Dim theMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
                theMxDoc = DirectCast(EditorExtension.Application.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)
                setActiveViewEvents(DirectCast(theMxDoc.FocusMap, IActiveViewEvents_Event))  ' TODO: [NIS] Reset this when the focus map changes?

                ' Set up document keyboard accelerators for extension commands.
                CreateAccelerators()

                ' Subscribe to edit events.
                AddHandler EditEvents.OnChangeFeature, AddressOf EditEvents_OnChangeFeature
                AddHandler EditEvents.OnCreateFeature, AddressOf EditEvents_OnCreateFeature

                ' Subscribe to active view events.
                AddHandler ActiveViewEvents.FocusMapChanged, AddressOf ActiveViewEvents_FocusMapChanged
                AddHandler ActiveViewEvents.ItemAdded, AddressOf ActiveViewEvents_ItemAdded
                AddHandler ActiveViewEvents.ItemDeleted, AddressOf ActiveViewEvents_ItemDeleted

                ' Set the valid data properties.
                ClearValidDataProperties()

            End If

        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Trace.WriteLine(ex.ToString)

        End Try

    End Sub

    Private Sub EditEvents_OnStopEditing(ByVal save As Boolean)

        Try
            ' Unsubscribe to edit events.
            RemoveHandler EditEvents.OnChangeFeature, AddressOf EditEvents_OnChangeFeature
            RemoveHandler EditEvents.OnCreateFeature, AddressOf EditEvents_OnCreateFeature

            ' Unsubscribe to active view events.
            RemoveHandler EditorExtension.ActiveViewEvents.FocusMapChanged, AddressOf ActiveViewEvents_FocusMapChanged
            RemoveHandler EditorExtension.ActiveViewEvents.ItemAdded, AddressOf ActiveViewEvents_ItemAdded
            RemoveHandler EditorExtension.ActiveViewEvents.ItemDeleted, AddressOf ActiveViewEvents_ItemDeleted

        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            Trace.WriteLine(ex.ToString)

        Finally
            setApplication(Nothing)
            setActiveViewEvents(Nothing)

            SetHasValidTaxlotData(False)
            SetHasValidMapIndexData(False)

        End Try

    End Sub

#End Region

#Region "ActiveViewEvents Event Handlers"

    Public Sub ActiveViewEvents_FocusMapChanged() 'Handles ESRI.ArcGIS.Carto.IActiveViewEvents.FocusMapChanged
        ' TODO: [NIS] Determine why this event never fires...
        ClearValidDataProperties()
    End Sub

    Public Sub ActiveViewEvents_ItemAdded(ByVal Item As Object) 'Handles ESRI.ArcGIS.Carto.IActiveViewEvents.ItemAdded
        ClearValidDataProperties()
    End Sub

    Public Sub ActiveViewEvents_ItemDeleted(ByVal Item As Object) 'Handles ESRI.ArcGIS.Carto.IActiveViewEvents.ItemDeleted
        ClearValidDataProperties()
    End Sub

#End Region

#End Region

#Region "Methods"

    ' TODO: [NIS] Test (not sure this how this will work with editor extension)
    Private Shared Sub setAccelerator(ByRef acceleratorTable As IAcceleratorTable, _
            ByVal classID As UID, ByVal key As Integer, _
            ByVal usesCtrl As Boolean, ByVal usesAlt As Boolean, _
            ByVal usesShift As Boolean)
        ' Create accelerator only if nothing else is using it

        Dim accelerator As IAccelerator

        accelerator = AcceleratorTable.FindByKey(key, usesCtrl, usesAlt, usesShift)
        If accelerator Is Nothing Then
            'The clsid of one of the commands in the ext
            AcceleratorTable.Add(ClassId, key, usesCtrl, usesAlt, usesShift)
        End If

    End Sub

    Private Shared Function validateLicense(ByVal requiredProductCode As esriLicenseProductCode) As Boolean
        ' Validate the license (e.g. ArcEditor or ArcInfo).

        Dim theAoInitializeClass As New AoInitializeClass()
        Dim productCode As esriLicenseProductCode = theAoInitializeClass.InitializedProduct()

        Return (productCode = requiredProductCode)
    End Function

#End Region

#End Region

#Region "Inherited Class Members (none)"
#End Region

#Region "Implemented Interface Members"

#Region "IExtension Interface Implementations"

    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
        Get
            Return "EditorExtension"
        End Get
    End Property

    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
        setEditor(Nothing)
        setEditEvents(Nothing)

        ' Unsubscribe to edit events.
        RemoveHandler EditEvents.OnStartEditing, AddressOf EditEvents_OnStartEditing
        RemoveHandler EditEvents.OnStopEditing, AddressOf EditEvents_OnStopEditing

    End Sub

    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
        If Not initializationData Is Nothing AndAlso TypeOf initializationData Is IEditor2 Then
            ' Set the Editor and EditEvents properties.
            setEditor(DirectCast(initializationData, IEditor2))
            setEditEvents(DirectCast(EditorExtension.Editor, IEditEvents_Event))

            My.User.InitializeWithWindowsUser()

            ' Subscribe to edit events.
            AddHandler EditEvents.OnStartEditing, AddressOf EditEvents_OnStartEditing
            AddHandler EditEvents.OnStopEditing, AddressOf EditEvents_OnStopEditing

        End If
    End Sub

#End Region

#Region "IExtensionAccelerators Interface Implementations"

    Public Sub CreateAccelerators() Implements ESRI.ArcGIS.esriSystem.IExtensionAccelerators.CreateAccelerators
        ' Create the keyboard accelerators for this extension.
        ' TODO: [NIS] Test this (not sure this will work with an editor extension)
        Dim key As Integer
        Dim usesCtrl As Boolean
        Dim usesAlt As Boolean
        Dim usesShift As Boolean
        Dim uid As New UID
        Dim doc As IDocument = EditorExtension.Application.Document
        Dim acceleratorTable As IAcceleratorTable = doc.Accelerators

        ' Set LocateFeature accelerator keys to Ctrl + Alt + L
        key = Convert.ToInt32(Keys.L)
        usesCtrl = True
        usesAlt = True
        usesShift = False
        uid.Value = "{" & OrmapTaxlotEditing.LocateFeature.ClassId & "}"
        setAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set TaxlotAssignment accelerator keys to Ctrl + Alt + T
        key = Convert.ToInt32(Keys.T)
        usesCtrl = True
        usesAlt = True
        usesShift = False
        uid.Value = "{" & OrmapTaxlotEditing.TaxlotAssignment.ClassId & "}"
        setAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set EditMapIndex accelerator keys to Ctrl + Alt + E
        key = Convert.ToInt32(Keys.E)
        usesCtrl = True
        usesAlt = True
        usesShift = False
        uid.Value = "{" & OrmapTaxlotEditing.EditMapIndex.ClassId & "}"
        setAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set CombineTaxlots accelerator keys to Ctrl + Alt + C
        key = Convert.ToInt32(Keys.C)
        usesCtrl = True
        usesAlt = True
        usesShift = False
        uid.Value = "{" & OrmapTaxlotEditing.CombineTaxlots.ClassId & "}"
        setAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

        ' Set AddArrows accelerator keys to Ctrl + Alt + A
        key = Convert.ToInt32(Keys.A)
        usesCtrl = True
        usesAlt = True
        usesShift = False
        uid.Value = "{" & OrmapTaxlotEditing.AddArrows.ClassId & "}"
        setAccelerator(acceleratorTable, uid, key, usesCtrl, usesAlt, usesShift)

    End Sub

#End Region

#Region "IPersistVariant Interface Implementations"

    Public ReadOnly Property ID() As ESRI.ArcGIS.esriSystem.UID Implements ESRI.ArcGIS.esriSystem.IPersistVariant.ID
        Get
            Dim uid As New UIDClass()
            uid.Value = "{" & OrmapTaxlotEditing.EditorExtension.ClassId & "}"
            Return uid
        End Get
    End Property

    Public Sub Load(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Load

        If Stream Is Nothing Then
            Throw New ArgumentNullException("Stream")
        End If

        AllowedToEditTaxlots = CBool(Stream.Read())
        AllowedToAutoUpdate = CBool(Stream.Read())
        AllowedToAutoUpdateAllFields = CBool(Stream.Read())

    End Sub

    Public Sub Save(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Save

        If Stream Is Nothing Then
            Throw New ArgumentNullException("Stream")
        End If

        Stream.Write(AllowedToEditTaxlots)
        Stream.Write(AllowedToAutoUpdate)
        Stream.Write(AllowedToAutoUpdateAllFields)

    End Sub

#End Region

#End Region

#Region "Other Members"

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "3ffddc1a-bf54-45b4-a9dc-88740d97bcc2"
    Public Const InterfaceId As String = "cf8fd284-b76e-4012-a738-bce6e0cbbff4"
    Public Const EventsId As String = "e5719155-369f-4b3e-9e5e-99856449f05b"
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
        EditorExtensions.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        EditorExtensions.Unregister(regKey)

    End Sub

#End Region
#End Region

#End Region

End Class

