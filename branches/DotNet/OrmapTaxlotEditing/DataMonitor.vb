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
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.StringUtilities
Imports OrmapTaxlotEditing.Utilities
Imports System.Configuration
Imports System.Windows.Forms
#End Region

#Region "Class Declaration"
''' <summary>
'''  Data monitoring class for required ORMAP datasets.
''' </summary>
''' <remarks>Keeps track of presence of valid ORMAP datasets in the map document.</remarks>
Public NotInheritable Class DataMonitor

#Region "Class-Level Constants And Enumerations"

    Public Enum ESRIClassType As Integer
        FeatureClass = 1
        ObjectClass = 2
    End Enum

#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    ''' <summary>
    ''' OrmapSettings constructor.
    ''' </summary>
    ''' <remarks>This class follows the singleton pattern and thus has a 
    ''' private constructor and all shared members.</remarks>
    Private Sub New()
    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields (none)"
#End Region

#Region "Properties"

    Private Shared _mapIndexFeatureLayer As IFeatureLayer

    Friend Shared ReadOnly Property MapIndexFeatureLayer() As IFeatureLayer
        Get
            Return _mapIndexFeatureLayer
        End Get
    End Property

    Private Shared Sub setMapIndexFeatureLayer(ByVal value As IFeatureLayer)
        _mapIndexFeatureLayer = value
    End Sub

    Private Shared _taxlotFeatureLayer As IFeatureLayer

    Friend Shared ReadOnly Property TaxlotFeatureLayer() As IFeatureLayer
        Get
            Return _taxlotFeatureLayer
        End Get
    End Property

    Private Shared Sub setTaxlotFeatureLayer(ByVal value As IFeatureLayer)
        _taxlotFeatureLayer = value
    End Sub

    Private Shared _hasValidMapIndexData As Boolean = False

    Friend Shared ReadOnly Property HasValidMapIndexData() As Boolean
        Get
            Return _hasValidMapIndexData
        End Get
    End Property

    Friend Shared Sub SetHasValidMapIndexData(ByVal value As Boolean)
        _hasValidMapIndexData = value
    End Sub

    Private Shared _hasValidTaxlotData As Boolean = False

    Friend Shared ReadOnly Property HasValidTaxlotData() As Boolean
        Get
            Return _hasValidTaxlotData
        End Get
    End Property

    Friend Shared Sub SetHasValidTaxlotData(ByVal value As Boolean)
        _hasValidTaxlotData = value
    End Sub

#End Region

#Region "Event Handlers (none)"
#End Region

#Region "Methods"

    Friend Shared Sub ClearValidDataProperties()
        ' MapIndex
        SetHasValidMapIndexData(False)
        setMapIndexFeatureLayer(Nothing)
        ' Taxlots
        SetHasValidTaxlotData(False)
        setTaxlotFeatureLayer(Nothing)
    End Sub

    Friend Shared Sub CheckValidDataProperties()
        ' MapIndex status and layer properties
        SetHasValidMapIndexData(CheckData(ESRIClassType.FeatureClass, EditorExtension.TableNamesSettings.MapIndexFC))
        If HasValidMapIndexData Then
            If MapIndexFeatureLayer Is Nothing Then
                setMapIndexFeatureLayer(FindDataLayerInMap(EditorExtension.TableNamesSettings.MapIndexFC))
            End If
        End If
        ' Taxlot status and layer properties
        SetHasValidTaxlotData(CheckData(ESRIClassType.FeatureClass, EditorExtension.TableNamesSettings.TaxLotFC))
        If HasValidTaxlotData Then
            If TaxlotFeatureLayer Is Nothing Then
                setTaxlotFeatureLayer(FindDataLayerInMap(EditorExtension.TableNamesSettings.TaxLotFC))
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
        If MessageBox.Show("Dataset " & className & " not found in the map. Load it?", "Load Data", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.No Then
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

        Dim isValid As Boolean = False
        
        Try
            If classType = ESRIClassType.FeatureClass Then

                Dim theFeatureClass As IFeatureClass = FindDataLayerInMap(className).FeatureClass
                Select Case className
                    Case EditorExtension.TableNamesSettings.CartographicLinesFC
                        isValid = theFeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline
                        With EditorExtension.CartographicLinesSettings
                            isValid = isValid AndAlso theFeatureClass.FindField(.LineTypeField) <> FieldNotFoundIndex
                        End With

                    Case EditorExtension.TableNamesSettings.MapIndexFC()
                        isValid = theFeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon
                        With EditorExtension.MapIndexSettings
                            isValid = isValid AndAlso theFeatureClass.FindField(.CountyField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapNumberField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.OrmapMapNumberField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapScaleField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapSuffixTypeField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapSuffixNumberField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.PageNumberField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.ReliabilityCodeField) <> FieldNotFoundIndex
                        End With

                    Case EditorExtension.TableNamesSettings.TaxLotFC
                        isValid = theFeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon
                        With EditorExtension.TaxLotSettings
                            isValid = isValid AndAlso theFeatureClass.FindField(.CountyField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.TownshipField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.TownshipPartField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.TownshipDirectionField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.RangeField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.RangePartField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.RangeDirectionField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.SectionNumberField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.QuarterSectionField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.QuarterQuarterSectionField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.OrmapMapNumberField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.OrmapTaxlotField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapNumberField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.TaxlotField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapTaxlotField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapTaxlotFormatMask) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.SpecialInterestField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapSuffixNumberField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapSuffixTypeField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.MapAcresField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theFeatureClass.FindField(.AnomalyField) <> FieldNotFoundIndex
                        End With

                    Case EditorExtension.TableNamesSettings.TaxLotLinesFC
                        isValid = theFeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon
                        With EditorExtension.TaxLotLinesSettings
                            isValid = isValid AndAlso theFeatureClass.FindField(.LineTypeField) <> FieldNotFoundIndex
                        End With

                    Case Else
                        '[Not a class for which shape type matters...]
                        '[No field names defined in settings...]
                        isValid = True
                End Select

            Else ' If classType = ESRIClassType.ObjectClass Then
                Dim theTable As ITable = FindDataTableInMap(className).Table
                Select Case className
                    Case EditorExtension.TableNamesSettings.CancelledNumbersTable
                        ' TODO: [NIS] Use EditorExtension.CancelledNumbersSettings instead (when available)
                        With EditorExtension.TaxLotSettings
                            isValid = isValid AndAlso theTable.FindField(.TaxlotField) <> FieldNotFoundIndex
                            isValid = isValid AndAlso theTable.FindField(.MapNumberField) <> FieldNotFoundIndex
                        End With
                    Case Else
                        '[No field names defined in settings...]
                        isValid = True
                End Select

            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Function

    Private Shared Function matchesFieldSettings(ByVal theFeatureClass As IClass, ByVal theSPC As SettingsPropertyCollection) As Boolean
        Try
            If theSPC.Count > 0 Then
                Dim foundAllFields As Boolean = True
                Dim theSPCEnumerator As System.Collections.IEnumerator = theSPC.GetEnumerator
                theSPCEnumerator.Reset()
                While theSPC.GetEnumerator.MoveNext()
                    Dim x As Object = theSPC.GetEnumerator.Current
                    foundAllFields = foundAllFields AndAlso (theFeatureClass.FindField(DirectCast(x, SettingsProperty).DefaultValue.ToString) <> FieldNotFoundIndex)
                End While
                Return foundAllFields
            Else
                '[There are no settings to check...]
                ' TODO: [NIS] Throw an exception - settings are missing!
            End If
            Return False

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

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
