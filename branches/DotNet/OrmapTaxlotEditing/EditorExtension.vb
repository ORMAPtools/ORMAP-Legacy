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
Imports ESRI.ArcGIS.Geometry
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

    Private Shared Sub setApplication(ByVal value As IApplication)
        _application = value
    End Sub

    Private Shared _editor As IEditor2

    Friend Shared ReadOnly Property Editor() As IEditor2
        Get
            Return _editor
        End Get
    End Property

    Private Shared Sub setEditor(ByVal value As IEditor2)
        _editor = value
    End Sub

    Private Shared _editEvents As IEditEvents_Event

    Friend Shared ReadOnly Property EditEvents() As IEditEvents_Event
        Get
            Return _editEvents
        End Get
    End Property

    Private Shared Sub setEditEvents(ByVal value As IEditEvents_Event)
        _editEvents = value
    End Sub

    Private Shared _activeViewEvents As IActiveViewEvents_Event

    Friend Shared ReadOnly Property ActiveViewEvents() As IActiveViewEvents_Event
        Get
            Return _activeViewEvents
        End Get
    End Property

    Private Shared Sub setActiveViewEvents(ByVal value As IActiveViewEvents_Event)
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

    Private Shared Sub setHasValidLicense(ByVal value As Boolean)
        _hasValidLicense = value
    End Sub

    Private Shared _isValidWorkspace As Boolean '= False

    Friend Shared ReadOnly Property IsValidWorkspace() As Boolean
        Get
            Return _isValidWorkspace
        End Get
    End Property

    Private Shared Sub setIsValidWorkspace(ByVal value As Boolean)
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

    Private _isDuringAutoUpdate As Boolean

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
            If Not IsOrmapFeature(obj) Then Exit Try

            ' Update the minimum auto-calculated fields
            UpdateMinimumAutoFields(DirectCast(obj, IFeature))

            If Not EditorExtension.AllowedToAutoUpdateAllFields Then Exit Try

            ' Avoid rentrancy
            If _isDuringAutoUpdate = False Then
                _isDuringAutoUpdate = True
            Else
                Throw New InvalidOperationException("Already in AutoUpdate mode. Cannot initiate AutoUpdate.")
                Exit Try
            End If

            ' Note: Must check here for if required data is available
            ' (in case subroutines called don't check).

            ' Check for valid data (will try to load data if not found).
            CheckValidTaxlotDataProperties()
            If Not HasValidTaxlotData Then
                MessageBox.Show("Unable to update Taxlot field values." & vbNewLine & _
                                "Missing data: Valid ORMAP Taxlot layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnChangeFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If
            CheckValidMapIndexDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Unable to update taxlot field values." & vbNewLine & _
                                "Missing data: Valid ORMAP MapIndex layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnChangeFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)

                Exit Try
            End If

            If IsTaxlot(obj) Then
                '[Edited object is a ORMAP taxlot feature...]

                ' Obtain OrmapMapNumber via overlay and calculate other field values.
                CalculateTaxlotValues(DirectCast(obj, IFeature), FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.MapIndexFC))

            ElseIf IsAnno(obj) Then
                '[Edited object is an ORMAP annotation feature...]

                ' Set anno size based on the map scale.
                SetAnnoSize(obj)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        Finally
            _isDuringAutoUpdate = False

        End Try

    End Sub

    ''' <summary>
    ''' Updates fields based on the feature that was just created.
    ''' </summary>
    ''' <param name="obj">The feature that was just created.</param>
    ''' <remarks>Handles EditEvents.OnCreateFeature events.</remarks>
    Private Sub EditEvents_OnCreateFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) 'Handles EditEvents.OnCreateFeature

        Try
            If Not EditorExtension.CanEnableExtendedEditing Then Exit Try
            If Not EditorExtension.AllowedToAutoUpdate Then Exit Try
            If Not IsOrmapFeature(obj) Then Exit Try

            ' Update the minimum auto-calculated fields
            UpdateMinimumAutoFields(DirectCast(obj, IFeature))

            If Not EditorExtension.AllowedToAutoUpdateAllFields Then Exit Try

            ' Note: Must check here for if required data is available
            ' (in case subroutines called don't check).

            ' Check for valid data (will try to load data if not found).
            CheckValidTaxlotDataProperties()
            If Not HasValidTaxlotData Then
                MessageBox.Show("Unable to populate Taxlot field values." & vbNewLine & _
                                "Missing data: Valid ORMAP Taxlot layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnCreateFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If
            CheckValidMapIndexDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Unable to populate taxlot field values." & vbNewLine & _
                                "Missing data: Valid ORMAP MapIndex layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnCreateFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If

            ' Get the feature
            Dim theFeature As ESRI.ArcGIS.Geodatabase.IFeature
            theFeature = DirectCast(obj, IFeature)

            ' Get the feature geometry
            Dim theGeometry As ESRI.ArcGIS.Geometry.IGeometry
            theGeometry = theFeature.Shape
            If theGeometry.IsEmpty Then
                Exit Try
            End If

            If IsTaxlot(obj) Then
                '[Edited object is a ORMAP taxlot feature...]

                ' Obtain OrmapMapNumber via overlay and calculate other field values.
                CalculateTaxlotValues(theFeature, MapIndexFeatureLayer)

            ElseIf IsAnno(obj) Then
                '[Edited object is an ORMAP annotation feature...]
                ' Update MapScale, MapNumber and Anno Size:

                ' Get the annotation feature.
                Dim theAnnotationFeature As ESRI.ArcGIS.Carto.IAnnotationFeature
                theAnnotationFeature = DirectCast(theFeature, IAnnotationFeature)

                'Get the parent feature
                Dim theParentID As Integer
                theParentID = theAnnotationFeature.LinkedFeatureID
                If theParentID > FieldNotFoundIndex Then 'Feature linked
                    theFeature = GetRelatedObjects(obj)
                    If theFeature Is Nothing Then Exit Try
                Else
                    'Not feature linked anno, so we can use the feature as is

                End If

                setMapIndexAndScale(obj, theFeature, theGeometry)

                ' Set anno size based on the map scale.
                SetAnnoSize(obj)
            Else
                '[Edited object is another kind of ORMAP feature (not taxlot or annotation)...]
                ' Update MapScale and MapNumber (except for on the MapIndex feature class):

                setMapIndexAndScale(obj, theFeature, theGeometry)

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
    Private Sub EditEvents_OnDeleteFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) 'Handles EditEvents.OnDeleteFeature

        Try
            If Not EditorExtension.CanEnableExtendedEditing Then Exit Try
            If Not EditorExtension.AllowedToAutoUpdate Then Exit Try
            If Not IsOrmapFeature(obj) Then Exit Try ' TODO: [NIS] Is this even needed here?
            If Not EditorExtension.AllowedToAutoUpdateAllFields Then Exit Try

            ' Note: Must check here for if required data is available
            ' (in case subroutines called don't check).

            ' Check for valid data (will try to load data if not found).
            CheckValidTaxlotDataProperties()
            If Not HasValidTaxlotData Then
                MessageBox.Show("Missing data: Valid ORMAP Taxlot layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnDeleteFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If
            CheckValidMapIndexDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnDeleteFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If
            CheckValidCancelledNumbersTableDataProperties()
            If Not HasValidCancelledNumbersTableData Then
                MessageBox.Show("Missing data: Valid ORMAP CancelledNumbersTable not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "ORMAP Taxlot Editing (OnDeleteFeature)", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If

            If IsTaxlot(obj) Then
                '[Deleting taxlots...]

                ' Capture the mapnumber and taxlot and record them in CancelledNumbers.

                ' Retrieve field positions.
                Dim theTLTaxlotFieldIndex As Integer = TaxlotFeatureLayer.FeatureClass.FindField(EditorExtension.TaxLotSettings.TaxlotField)
                Dim theTLMapNumberFieldIndex As Integer = TaxlotFeatureLayer.FeatureClass.FindField(EditorExtension.TaxLotSettings.MapNumberField)
                Dim theCNTaxlotFieldIndex As Integer = CancelledNumbersTable.Table.FindField(EditorExtension.TaxLotSettings.TaxlotField)
                Dim theCNMapNumberFieldIndex As Integer = CancelledNumbersTable.Table.FindField(EditorExtension.TaxLotSettings.MapNumberField)

                Dim theFeature As IFeature = DirectCast(obj, IFeature)
                
                ' If no null values, copy them to Cancelled numbers
                If Not IsDBNull(theFeature.Value(theTLTaxlotFieldIndex)) And Not IsDBNull(theFeature.Value(theTLMapNumberFieldIndex)) Then

                    ' Taxlots will send their numbers to the CancelledNumbers table 
                    ' ONLY if they are unique in the map at the time of deletion.
                    Dim theTaxlotNumber As String = CStr(theFeature.Value(theTLTaxlotFieldIndex))
                    Dim theArea As IArea = DirectCast(theFeature.Shape, IArea)
                    If IsTaxlotNumberLocallyUnique(theTaxlotNumber, theArea.Centroid) Then
                        Dim theRow As ESRI.ArcGIS.Geodatabase.IRow
                        theRow = CancelledNumbersTable.Table.CreateRow
                        theRow.Value(theCNTaxlotFieldIndex) = theFeature.Value(theTLTaxlotFieldIndex)
                        theRow.Value(theCNMapNumberFieldIndex) = theFeature.Value(theTLMapNumberFieldIndex)
                        theRow.Store()
                    End If

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
                AddHandler EditEvents.OnDeleteFeature, AddressOf EditEvents_OnDeleteFeature

                ' Subscribe to active view events.
                AddHandler ActiveViewEvents.FocusMapChanged, AddressOf ActiveViewEvents_FocusMapChanged
                AddHandler ActiveViewEvents.ItemAdded, AddressOf ActiveViewEvents_ItemAdded
                AddHandler ActiveViewEvents.ItemDeleted, AddressOf ActiveViewEvents_ItemDeleted

                ' Set the valid data properties.
                ClearAllValidDataProperties()

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
            RemoveHandler EditEvents.OnDeleteFeature, AddressOf EditEvents_OnDeleteFeature

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
        ClearAllValidDataProperties()
    End Sub

    Public Sub ActiveViewEvents_ItemAdded(ByVal Item As Object) 'Handles ESRI.ArcGIS.Carto.IActiveViewEvents.ItemAdded
        ClearAllValidDataProperties()
    End Sub

    Public Sub ActiveViewEvents_ItemDeleted(ByVal Item As Object) 'Handles ESRI.ArcGIS.Carto.IActiveViewEvents.ItemDeleted
        ClearAllValidDataProperties()
    End Sub

#End Region

#End Region

#Region "Methods"

    ' TODO: [NIS] Test (not sure this how this will work with editor extension)
    Private Shared Sub setAccelerator(ByVal acceleratorTable As IAcceleratorTable, _
            ByVal classID As UID, ByVal key As Integer, _
            ByVal usesCtrl As Boolean, ByVal usesAlt As Boolean, _
            ByVal usesShift As Boolean)
        ' Create accelerator only if nothing else is using it

        Dim accelerator As IAccelerator

        accelerator = acceleratorTable.FindByKey(key, usesCtrl, usesAlt, usesShift)
        If accelerator Is Nothing Then
            'The clsid of one of the commands in the ext
            acceleratorTable.Add(classID, key, usesCtrl, usesAlt, usesShift)
        End If

    End Sub

    Private Shared Function validateLicense(ByVal requiredProductCode As esriLicenseProductCode) As Boolean
        ' Validate the license (e.g. ArcEditor or ArcInfo).

        Dim theAoInitializeClass As New AoInitializeClass()
        Dim productCode As esriLicenseProductCode = theAoInitializeClass.InitializedProduct()

        Return (productCode = requiredProductCode)
    End Function

    Private Shared Sub setMapIndexAndScale(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject, ByVal theFeature As ESRI.ArcGIS.Geodatabase.IFeature, ByVal theGeometry As ESRI.ArcGIS.Geometry.IGeometry)
        ' Set the map index (if the field exists) to the Map Index map index for the feature location:

        Dim theMapScale As String
        Dim theMapNumber As String

        ' Get the Map Index map number field index.
        Dim theMapNumFieldIndex As Integer
        theMapNumFieldIndex = theFeature.Fields.FindField(EditorExtension.MapIndexSettings.MapNumberField)
        If theMapNumFieldIndex > FieldNotFoundIndex And Not IsMapIndex(obj) Then
            ' Get the Map Index map number for the location of the new feature.
            theMapNumber = GetValueViaOverlay(theGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapNumberField, EditorExtension.MapIndexSettings.MapNumberField)
            ' Set the feature map number.
            If Len(theMapNumber) > 0 Then
                theFeature.Value(theMapNumFieldIndex) = theMapNumber
            Else
                theFeature.Value(theMapNumFieldIndex) = System.DBNull.Value
            End If
        End If

        ' Set the map scale (if the field exists) to the Map Index map scale for the feature location:

        ' Get the Map Index map scale field index.
        Dim theMapScaleFieldIndex As Integer
        theMapScaleFieldIndex = theFeature.Fields.FindField(EditorExtension.MapIndexSettings.MapScaleField)
        If theMapScaleFieldIndex > FieldNotFoundIndex And Not IsMapIndex(obj) Then
            ' Get the Map Index map scale for the location of the new feature.
            theMapScale = GetValueViaOverlay(theGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapScaleField, EditorExtension.MapIndexSettings.MapNumberField)
            ' Set the feature map scale.
            If Len(theMapScale) > 0 Then
                theFeature.Value(theMapScaleFieldIndex) = theMapScale
            Else
                theFeature.Value(theMapScaleFieldIndex) = System.DBNull.Value
            End If
        End If
    End Sub

#End Region

#End Region

#Region "Inherited Class Members (none)"
#End Region

#Region "Implemented Interface Members"

#Region "IExtension Interface Implementation"

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

#Region "IExtensionAccelerators Interface Implementation"

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

#Region "IPersistVariant Interface Implementation"

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


