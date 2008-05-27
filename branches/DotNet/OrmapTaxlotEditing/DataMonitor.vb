#Region "Copyright 2008 ORMAP Tech Group"

' File:  DataMonitor.vb
'
' Original Author:  OPET.NET Migration Team (Shad Campbell, James Moore, 
'                   Nick Seigal)
'
' Date Created:  April 23, 2008
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
'SCC revision number: $Revision: 256 $
'Date of Last Change: $Date: 2008-04-22 19:19:27 -0700 (Tue, 22 Apr 2008) $
#End Region

#Region "Imported Namespaces"
Imports System.Configuration
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.StringUtilities
Imports OrmapTaxlotEditing.Utilities

#End Region

#Region "Class Declaration"
''' <summary>
'''  Data monitoring class for required ORMAP datasets.
''' </summary>
''' <remarks>Keeps track of presence of valid ORMAP datasets in the map document.</remarks>
Public NotInheritable Class DataMonitor

#Region "Class-Level Constants And Enumerations"

    Friend Enum ESRIClassType As Integer
        FeatureClass = 1
        ObjectClass = 2
    End Enum

#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    ''' <summary>
    ''' Private empty constructor to prevent instantiation.
    ''' </summary>
    ''' <remarks>This class follows the singleton pattern and thus has a 
    ''' private constructor and all shared members. Instances of types 
    ''' that define only shared members do not need to be created, so no
    ''' constructor should be needed. However, many compilers will 
    ''' automatically add a public default constructor if no constructor 
    ''' is specified. To prevent this an empty private constructor is 
    ''' added.</remarks>
    Private Sub New()
    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields (none)"
#End Region

#Region "Properties"

    Private Shared _taxlotFeatureLayer As IFeatureLayer
    Friend Shared ReadOnly Property TaxlotFeatureLayer() As IFeatureLayer
        Get
            Return _taxlotFeatureLayer
        End Get
    End Property
    Private Shared Sub setTaxlotFeatureLayer(ByRef value As IFeatureLayer)
        _taxlotFeatureLayer = value
    End Sub

    Private Shared _mapIndexFeatureLayer As IFeatureLayer
    Friend Shared ReadOnly Property MapIndexFeatureLayer() As IFeatureLayer
        Get
            Return _mapIndexFeatureLayer
        End Get
    End Property
    Private Shared Sub setMapIndexFeatureLayer(ByRef value As IFeatureLayer)
        _mapIndexFeatureLayer = value
    End Sub

    Private Shared _cancelledNumbersTable As IStandaloneTable
    Friend Shared ReadOnly Property CancelledNumbersTable() As IStandaloneTable
        Get
            Return _cancelledNumbersTable
        End Get
    End Property
    Private Shared Sub setCancelledNumbersTable(ByRef value As IStandaloneTable)
        _cancelledNumbersTable = value
    End Sub

    Private Shared _taxlotLinesFeatureLayer As IFeatureLayer
    Friend Shared ReadOnly Property TaxlotLinesFeatureLayer() As IFeatureLayer
        Get
            Return _taxlotLinesFeatureLayer
        End Get
    End Property
    Private Shared Sub setTaxlotLinesFeatureLayer(ByRef value As IFeatureLayer)
        _taxlotLinesFeatureLayer = value
    End Sub

    Private Shared _referenceLinesFeatureLayer As IFeatureLayer
    Friend Shared ReadOnly Property ReferenceLinesFeatureLayer() As IFeatureLayer
        Get
            Return _referenceLinesFeatureLayer
        End Get
    End Property
    Private Shared Sub setReferenceLinesFeatureLayer(ByRef value As IFeatureLayer)
        _referenceLinesFeatureLayer = value
    End Sub

    Private Shared _hasValidMapIndexData As Boolean
    Friend Shared ReadOnly Property HasValidMapIndexData() As Boolean
        Get
            Return _hasValidMapIndexData
        End Get
    End Property
    Friend Shared Sub SetHasValidMapIndexData(ByVal value As Boolean)
        _hasValidMapIndexData = value
    End Sub

    Private Shared _hasValidTaxlotData As Boolean
    Friend Shared ReadOnly Property HasValidTaxlotData() As Boolean
        Get
            Return _hasValidTaxlotData
        End Get
    End Property
    Friend Shared Sub SetHasValidTaxlotData(ByVal value As Boolean)
        _hasValidTaxlotData = value
    End Sub

    Private Shared _hasValidCancelledNumbersTableData As Boolean
    Friend Shared ReadOnly Property HasValidCancelledNumbersTableData() As Boolean
        Get
            Return _hasValidCancelledNumbersTableData
        End Get
    End Property
    Friend Shared Sub SetHasValidCancelledNumbersTableData(ByVal value As Boolean)
        _hasValidCancelledNumbersTableData = value
    End Sub

    Private Shared _hasValidTaxlotLinesData As Boolean
    Friend Shared ReadOnly Property HasValidTaxlotLinesData() As Boolean
        Get
            Return _hasValidTaxlotLinesData
        End Get
    End Property
    Friend Shared Sub SetHasValidTaxlotLinesData(ByVal value As Boolean)
        _hasValidTaxlotLinesData = value
    End Sub

    Private Shared _hasValidReferenceLinesData As Boolean
    Friend Shared ReadOnly Property HasValidReferenceLinesData() As Boolean
        Get
            Return _hasValidReferenceLinesData
        End Get
    End Property
    Friend Shared Sub SetHasValidReferenceLinesData(ByVal value As Boolean)
        _hasValidReferenceLinesData = value
    End Sub

#End Region

#Region "Event Handlers (none)"
#End Region

#Region "Methods"

    Friend Shared Sub ClearAllValidDataProperties()
        ' Taxlots
        SetHasValidTaxlotData(False)
        setTaxlotFeatureLayer(Nothing)

        ' MapIndex
        SetHasValidMapIndexData(False)
        setMapIndexFeatureLayer(Nothing)

        ' CancelledNumbers
        SetHasValidCancelledNumbersTableData(False)
        setCancelledNumbersTable(Nothing)

        ' TaxlotLines
        SetHasValidTaxlotLinesData(False)
        setTaxlotLinesFeatureLayer(Nothing)

        ' ReferenceLines
        SetHasValidReferenceLinesData(False)
        setReferenceLinesFeatureLayer(Nothing)

    End Sub

    Friend Shared Sub CheckAllValidDataProperties()
        ' Taxlot status and layer properties
        CheckValidTaxlotDataProperties()
        ' MapIndex status and layer properties
        CheckValidMapIndexDataProperties()
        ' CancelledNumbersTable status and table properties
        CheckValidCancelledNumbersTableDataProperties()
    End Sub

    Friend Shared Sub CheckValidTaxlotDataProperties()
        ' Taxlot status and layer properties
        SetHasValidTaxlotData(CheckData(ESRIClassType.FeatureClass, EditorExtension.TableNamesSettings.TaxLotFC))
        If HasValidTaxlotData Then
            If TaxlotFeatureLayer Is Nothing Then
                setTaxlotFeatureLayer(FindDataLayerInMap(EditorExtension.TableNamesSettings.TaxLotFC))
            End If
        End If
    End Sub

    Friend Shared Sub CheckValidMapIndexDataProperties()
        ' MapIndex status and layer properties
        SetHasValidMapIndexData(CheckData(ESRIClassType.FeatureClass, EditorExtension.TableNamesSettings.MapIndexFC))
        If HasValidMapIndexData Then
            If MapIndexFeatureLayer Is Nothing Then
                setMapIndexFeatureLayer(FindDataLayerInMap(EditorExtension.TableNamesSettings.MapIndexFC))
            End If
        End If
    End Sub

    Friend Shared Sub CheckValidCancelledNumbersTableDataProperties()
        ' CancelledNumbersTable status and table properties
        SetHasValidCancelledNumbersTableData(CheckData(ESRIClassType.ObjectClass, EditorExtension.TableNamesSettings.CancelledNumbersTable))
        If HasValidCancelledNumbersTableData Then
            If CancelledNumbersTable Is Nothing Then
                setCancelledNumbersTable(FindDataTableInMap(EditorExtension.TableNamesSettings.CancelledNumbersTable))
            End If
        End If
    End Sub

    Friend Shared Sub CheckValidTaxlotLinesDataProperties()
        ' TaxlotLines status and layer properties
        SetHasValidTaxlotLinesData(CheckData(ESRIClassType.FeatureClass, EditorExtension.TableNamesSettings.TaxLotLinesFC))
        If HasValidTaxlotLinesData Then
            If TaxlotLinesFeatureLayer Is Nothing Then
                setTaxlotLinesFeatureLayer(FindDataLayerInMap(EditorExtension.TableNamesSettings.TaxLotLinesFC))
            End If
        End If
    End Sub

    Friend Shared Sub CheckValidReferenceLinesDataProperties()
        ' ReferenceLines status and layer properties
        SetHasValidReferenceLinesData(CheckData(ESRIClassType.FeatureClass, EditorExtension.TableNamesSettings.ReferenceLinesFC))
        If HasValidReferenceLinesData Then
            If ReferenceLinesFeatureLayer Is Nothing Then
                setReferenceLinesFeatureLayer(FindDataLayerInMap(EditorExtension.TableNamesSettings.ReferenceLinesFC))
            End If
        End If
    End Sub

    Friend Shared Function CheckData(ByVal classType As ESRIClassType, ByVal className As String) As Boolean
        Dim foundValidData As Boolean = False
        ' Look for the data in the map
        If classType = ESRIClassType.FeatureClass Then
            foundValidData = foundValidData OrElse (Not FindDataLayerInMap(className) Is Nothing)
        Else 'If classType = ESRIClassType.ObjectClass Then
            foundValidData = foundValidData OrElse (Not FindDataTableInMap(className) Is Nothing)
        End If
        ' Load the data if not found
        foundValidData = foundValidData OrElse loadOptionSuccessful(classType, className)
        ' Validate the data if found or loaded
        foundValidData = foundValidData AndAlso validateData(classType, className)
        Return foundValidData
    End Function

    Friend Shared Function FindDataLayerInMap(ByVal featureClassName As String) As IFeatureLayer
        ' Find data layer in the current map
        Dim theFLayer As IFeatureLayer
        theFLayer = FindFeatureLayerByDSName(featureClassName)
        Return theFLayer
    End Function

    Friend Shared Function FindDataTableInMap(ByVal objectClassName As String) As IStandaloneTable
        ' Find data layer in the current map
        Dim theStandaloneTable As IStandaloneTable
        theStandaloneTable = FindStandaloneTableByDSName(objectClassName)
        Return theStandaloneTable
    End Function

    Private Shared Function loadOptionSuccessful(ByVal classType As ESRIClassType, ByVal className As String) As Boolean
        ' Offer load option
        If MessageBox.Show("Dataset " & className & " not found in the map. Load it?", "Load Data", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.OK Then
            Return loadDataIntoMap(classType, className)
        Else
            Return False
        End If
    End Function

    Private Shared Function loadDataIntoMap(ByVal classType As ESRIClassType, ByVal className As String) As Boolean
        ' Attempt to load and find the data in the map document
        If classType = ESRIClassType.FeatureClass Then
            Return LoadFCIntoMap(className)
        Else 'If classType = ESRIClassType.ObjectClass Then
            Return LoadTableIntoMap(className)
        End If
    End Function

    ''' <summary>
    ''' Checks data for valid schema.
    ''' </summary>
    ''' <param name="classType">The class type (ObjectClass or FeatureClass).</param>
    ''' <param name="className">The class to validate.</param>
    ''' <returns><c>True</c> or <c>False</c>.</returns>
    ''' <remarks>' Valid schema in this case means the right feature type 
    ''' and required fields.</remarks>
    Private Shared Function validateData(ByVal classType As ESRIClassType, ByVal className As String) As Boolean

        Dim isValid As Boolean = False 'initialize

        Try
            If classType = ESRIClassType.FeatureClass Then

                Dim theFeatureClass As IFeatureClass = FindDataLayerInMap(className).FeatureClass
                Select Case className
                    Case EditorExtension.TableNamesSettings.CartographicLinesFC
                        ' ENHANCE: [NIS] MapNumber and MapScale should be added to CartographicLinesSettings.
                        ' ENHANCE: [NIS] Add MapNumber and MapScale validation here.
                        isValid = (theFeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline)
                        With EditorExtension.CartographicLinesSettings
                            isValid = isValid AndAlso theFeatureClass.FindField(.LineTypeField) <> NotFoundIndex
                        End With

                    Case EditorExtension.TableNamesSettings.MapIndexFC()
                        isValid = (theFeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon)
                        With EditorExtension.MapIndexSettings
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapScaleField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapNumberField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.OrmapMapNumberField) <> NotFoundIndex
                            ' ENHANCE: [NIS] CityName should be added to MapIndexSettings. See ORMAP spec note from 1/13/06.
                            ' CityName (not in settings)
                            isValid = isValid AndAlso theFeatureClass.FindField(.PageNumberField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.ReliabilityCodeField) <> NotFoundIndex
                            ' ENHANCE: [NIS] County, MapSuffixType and MapSuffixNum should be removed from MapIndexSettings. See ORMAP spec note from 2/10/05.
                            'isValid = isValid AndAlso theFeatureClass.FindField(.CountyField) <> NotFoundIndex
                            'isValid = isValid AndAlso theFeatureClass.FindField(.MapSuffixTypeField) <> NotFoundIndex
                            'isValid = isValid AndAlso theFeatureClass.FindField(.MapSuffixNumberField) <> NotFoundIndex
                        End With

                    Case EditorExtension.TableNamesSettings.TaxLotFC
                        isValid = (theFeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon)
                        With EditorExtension.TaxLotSettings
                            isValid = isValid AndAlso theFeatureClass.FindField(.CountyField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.TownshipField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.TownshipPartField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.TownshipDirectionField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.RangeField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.RangePartField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.RangeDirectionField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.SectionNumberField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.QuarterSectionField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.QuarterQuarterSectionField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.OrmapMapNumberField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.OrmapTaxlotField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapNumberField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.TaxlotField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapTaxlotField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.SpecialInterestField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapSuffixNumberField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapSuffixTypeField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapAcresField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.AnomalyField) <> NotFoundIndex
                        End With

                    Case EditorExtension.TableNamesSettings.TaxLotLinesFC
                        isValid = (theFeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline)
                        With EditorExtension.TaxLotLinesSettings
                            isValid = isValid AndAlso theFeatureClass.FindField(.LineTypeField) <> NotFoundIndex
                        End With

                    Case Else
                        '[Not an ORMAP feature class for which shape type matters...]
                        isValid = True
                        '[No field names defined in settings...]
                        ' ENHANCE: [NIS] Create other settings files and connect them to the application in various places.
                        ' ENHANCE: [NIS] MapNumber and MapScale should be added to other settings files.
                        ' ENHANCE: [NIS] Add MapNumber and MapScale validation here.
                        ' ENHANCE: [NIS] Remove this with block in favor of validation based on TaxLotLinesSettings (when these settings are available).
                        With EditorExtension.MapIndexSettings
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapScaleField) <> NotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapNumberField) <> NotFoundIndex
                        End With
                End Select

            Else ' If classType = ESRIClassType.ObjectClass Then
                Dim theTable As ITable = FindDataTableInMap(className).Table
                Select Case className
                    ' ENHANCE: [NIS] Remove this with block in favor of validation based on CancelledNumbersTableSettings (when available).
                    Case EditorExtension.TableNamesSettings.CancelledNumbersTable
                        isValid = True
                        ' ENHANCE: [NIS] Create CancelledNumbersSettings file and connect it to the application in various places.
                        With EditorExtension.TaxLotSettings
                            isValid = isValid AndAlso theTable.FindField(.TaxlotField) <> NotFoundIndex
                            isValid = isValid AndAlso theTable.FindField(.MapNumberField) <> NotFoundIndex
                        End With
                    Case Else
                        '[No field names defined in settings...]
                        isValid = True
                End Select

            End If

            Return isValid

        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)

        End Try

    End Function

#End Region

#End Region

#Region "Inherited Class Members (none)"

#Region "Properties (none)"
#End Region

#Region "Methods (none)"
#End Region

#End Region

#Region "Implemented Interface Members (none)"
#End Region

#Region "Other Members (none)"
#End Region

End Class
#End Region
