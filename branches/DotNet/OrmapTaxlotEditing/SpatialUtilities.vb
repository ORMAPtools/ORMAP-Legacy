#Region "Copyright 2008 ORMAP Tech Group"

' File:  SpatialUtilities.vb
'
' Original Author:  OPET.NET Migration Team (Shad Campbell, James Moore, 
'                   Nick Seigal)
'
' Date Created:  20080221
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
Imports ESRI.ArcGIS
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports OrmapTaxlotEditing.StringUtilities
Imports OrmapTaxlotEditing.Utilities
Imports System.Collections.Generic
Imports System.Text
Imports System.Windows.Forms
#End Region

#Region "Class Declaration"
''' <summary>
'''  Spatial utility class.
''' </summary>
''' <remarks>Commonly used ArcObjects procedures and functions.</remarks>
Public NotInheritable Class SpatialUtilities

#Region "Custom Class Members"

#Region "Public Members"
    ''' <summary>
    ''' Add the descriptive values from each domain to the drop down comboboxes.
    ''' </summary>
    ''' <param name="fieldName">Name of the field to draw the domain from.</param>
    ''' <param name="fields">The fields collection that contains <paramref name="fieldName">fieldName</paramref>.</param>
    ''' <param name="aComboBox">The combobox to populate.</param>
    ''' <param name="currentValue">The current value of the field.</param>
    ''' <param name="allowSpace">Allow a space/null entry in the list.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddCodesToCombo(ByVal fieldName As String, ByVal fields As IFields, ByRef aComboBox As ComboBox, ByVal currentValue As Object, Optional ByVal allowSpace As Boolean = False) As Boolean
        Dim returnValue As Boolean = False
        Try
            Dim theFldIdx As Integer = fields.FindField(fieldName)
            If theFldIdx > -1 Then
                Dim thisField As IField
                thisField = fields.Field(theFldIdx)
                Dim thisDomain As IDomain
                thisDomain = thisField.Domain
                If Not (thisDomain Is Nothing) Then
                    If TypeOf thisDomain Is ICodedValueDomain Then
                        Dim thisCodedValueDomain As ICodedValueDomain
                        thisCodedValueDomain = DirectCast(thisDomain, ICodedValueDomain)
                        Dim codeCount As Integer = thisCodedValueDomain.CodeCount
                        If Not allowSpace Then
                            With aComboBox
                                If .Items.Count > 0 Then
                                    'find the blank
                                    Dim textPosition As Integer = .FindStringExact(String.Empty, -1) 'HACK: JWM this is my best guess on how to find null string
                                    If textPosition > -1 Then
                                        .Items.RemoveAt(textPosition)
                                    End If
                                End If
                            End With
                        End If
                        For i As Integer = 0 To codeCount - 1
                            aComboBox.Items.Add(thisCodedValueDomain.Name(i))
                        Next i
                        'If current value is null, add an empty string and make it active
                        If TypeOf currentValue Is String Then
                            If currentValue.Equals(String.Empty) Then
                                If allowSpace Then
                                    aComboBox.Items.Add(String.Empty)
                                    aComboBox.SelectedIndex = aComboBox.FindStringExact(String.Empty, 0)
                                Else
                                    aComboBox.SelectedIndex = 0
                                End If
                            Else 'Otherwise, select the existing value from the list
                                aComboBox.SelectedIndex = aComboBox.FindStringExact(CStr(currentValue), 0)
                            End If
                            returnValue = True
                        End If 'if a valid domain
                    End If 'field not found
                End If
            End If
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Obtains ORMapNum via overlay and calculates other field values.
    ''' </summary>
    ''' <param name="editFeature">A feature from the Taxlot feature class.</param>
    ''' <param name="mapIndexLayer">The Map Index feature layer.</param>
    ''' <remarks><para>Updates the ORMAP fields in <paramref name="editFeature"/> to 
    ''' reflect the current ORMapNum and MapNumber values in the 
    ''' overlaying <paramref name="mapIndexLayer"/> polygon.</para>
    ''' <para>The following fields are updated:</para>
    ''' <list type="table">
    '''   <listheader><term>Field</term><description>Source</description></listheader>
    '''   <item><term>County</term><description>MapIndex</description></item>
    '''   <item><term>Town</term><description>MapIndex</description></item>
    '''   <item><term>TownPart</term><description>MapIndex</description></item>
    '''   <item><term>TownDir</term><description>MapIndex</description></item>
    '''   <item><term>Range</term><description>MapIndex</description></item>
    '''   <item><term>RangePart</term><description>MapIndex</description></item>
    '''   <item><term>RangeDir</term><description>MapIndex</description></item>
    '''   <item><term>Section</term><description>MapIndex</description></item>
    '''   <item><term>Qrtr</term><description>MapIndex</description></item>
    '''   <item><term>QrtrQrtr</term><description>MapIndex</description></item>
    '''   <item><term>MapSufType</term><description>MapIndex</description></item>
    '''   <item><term>MapSufNum</term><description>MapIndex</description></item>
    '''   <item><term>Anomaly</term><description>MapIndex</description></item>
    '''   <item><term>MapNumber</term><description>MapIndex</description></item>
    '''   <item><term>ORMapNum</term><description>MapIndex</description></item>
    '''   <item><term>Taxlot</term><description>(not updated here)</description></item>
    '''   <item><term>SpcIntrst</term><description>(not updated here)</description></item>
    '''   <item><term>MapTaxlot</term><description>Combination of MapIndex.MapNum and current Taxlot</description></item>
    '''   <item><term>ORTaxlot</term><description>Combination of MapIndex.ORMapNum and current Taxlot</description></item>
    '''   <item><term>MapAcres</term><description>(feature area / 43560)</description></item>
    ''' </list>
    ''' </remarks>
    Public Shared Sub CalculateTaxlotValues(ByRef editFeature As ESRI.ArcGIS.Geodatabase.IFeature, ByRef mapIndexLayer As ESRI.ArcGIS.Carto.IFeatureLayer)

        Dim theORMapNumClass As New ORMapNum()

        Try
            Dim taxlotFClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
            taxlotFClass = DirectCast(editFeature.Class, ESRI.ArcGIS.Geodatabase.IFeatureClass)

            ' Verify the map index layer
            If mapIndexLayer Is Nothing Then
                ' Prompt the user for the location of the MapIndex feature class
                If LoadFCIntoMap(EditorExtension.TableNamesSettings.MapIndexFC, "Locate Database with Map Index") Then
                    mapIndexLayer = FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.MapIndexFC)
                End If
                ' Tests for a failure to load the MapIndex feature class
                If mapIndexLayer Is Nothing Then
                    Exit Try
                End If
            End If

            Dim theCountyFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.CountyField)
            Dim theTownFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.TownshipField)
            Dim theTownPartFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.TownshipPartField)
            Dim theTownDirFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.TownshipDirectionField)
            Dim theRangeFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.RangeField)
            Dim theRangePartFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.RangePartField)
            Dim theRangeDirFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.RangeDirectionField)
            Dim theSectionFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.SectionNumberField)
            Dim theQrtrFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.QuarterSectionField)
            Dim theQrtrQrtrFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.QuarterQuarterSectionField)
            Dim theMapSufTypeFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.MapSuffixTypeField)
            Dim theMapSufNumFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.MapSuffixNumberField)
            Dim theAnomalyFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.AnomalyField)
            Dim theMapNumberFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.MapNumberField)
            Dim theORMapNumFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.OrmapMapNumberField)
            Dim theTaxlotFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.TaxlotField)
            Dim theSpcIntrstFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.SpecialInterestField)
            Dim theMapTaxlotFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.MapTaxlotField)
            Dim theORTaxlotFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.OrmapTaxlotField)
            Dim theMapAcresFldIdx As Integer = LocateFields(taxlotFClass, EditorExtension.TaxLotSettings.MapAcresField)
            'TODO: JWM If any of these index fields are -1 we should bail according to VB6 version
            'I wonder if we multiplied these values and checked the result for a negative value. Of course if there were an even number of negative values
            'they would cancel each other out.

            '------------------------------------------
            ' Set the area
            '------------------------------------------
            Dim theArea As IArea
            theArea = DirectCast(editFeature.Shape, ESRI.ArcGIS.Geometry.IArea)

            '------------------------------------------
            ' Get the county MapNumber from the MapIndex
            ' layer and set the feature's MapNumber.
            '------------------------------------------
            Dim theCurrentMapNumber As String = GetValueViaOverlay(editFeature.ShapeCopy, mapIndexLayer.FeatureClass, EditorExtension.MapIndexSettings.MapNumberField, EditorExtension.MapIndexSettings.MapNumberField)
            If theCurrentMapNumber.Length = 0 Then
                Exit Try
            End If

            '------------------------------------------
            ' Reformat Special Interest Code
            '------------------------------------------
            Dim theCurrentSpecialInterest As String = "00000"
            If Not IsDBNull(editFeature.Value(theSpcIntrstFldIdx)) Then
                theCurrentSpecialInterest = CStr(editFeature.Value(theSpcIntrstFldIdx))
                If theCurrentSpecialInterest.Length < 5 Then
                    theCurrentSpecialInterest = theCurrentSpecialInterest.PadLeft(5, "0"c)
                ElseIf theCurrentSpecialInterest.Length > 5 Then
                    theCurrentSpecialInterest = theCurrentSpecialInterest.Substring(0, 5)
                End If
            End If

            '------------------------------------------
            ' Get the ORMapNum from the MapIndex 
            ' layer and parse it into the ORMapNum 
            ' object. Used below for field values.
            '------------------------------------------
            Dim theORMapNum As String = GetValueViaOverlay(editFeature.ShapeCopy, mapIndexLayer.FeatureClass, EditorExtension.MapIndexSettings.OrmapMapNumberField, EditorExtension.MapIndexSettings.MapNumberField)
            If Not theORMapNumClass.ParseNumber(theORMapNum) Then
                ' Exit if parse failed
                Exit Try
            End If

            '------------------------------------------
            ' Set all field values
            '------------------------------------------

            With editFeature
                .Value(theCountyFldIdx) = CShort(theORMapNumClass.County)
                .Value(theTownFldIdx) = CShort(theORMapNumClass.Township)
                .Value(theTownPartFldIdx) = CSng(theORMapNumClass.PartialTownshipCode)
                .Value(theTownDirFldIdx) = theORMapNumClass.TownshipDirectional
                .Value(theRangeFldIdx) = CShort(theORMapNumClass.Range)
                .Value(theRangePartFldIdx) = CSng(theORMapNumClass.PartialRangeCode)
                .Value(theRangeDirFldIdx) = theORMapNumClass.RangeDirectional
                .Value(theSectionFldIdx) = CShort(theORMapNumClass.Section)
                .Value(theQrtrFldIdx) = theORMapNumClass.Quarter
                .Value(theQrtrQrtrFldIdx) = theORMapNumClass.QuarterQuarter
                .Value(theMapSufTypeFldIdx) = theORMapNumClass.SuffixType
                .Value(theMapSufNumFldIdx) = CLng(theORMapNumClass.SuffixNumber)
                .Value(theAnomalyFldIdx) = theORMapNumClass.Anomaly
                .Value(theMapNumberFldIdx) = theCurrentMapNumber
                .Value(theORMapNumFldIdx) = theORMapNumClass.GetOrmapMapNumber
                'Taxlot (not updated here)
                .Value(theSpcIntrstFldIdx) = theCurrentSpecialInterest
                'MapTaxlot (see below)
                'ORTaxlot (see below)
                .Value(theMapAcresFldIdx) = theArea.Area / 43560
            End With

            '------------------------------------------
            ' Recalculate ORTaxlot
            '------------------------------------------
            If IsDBNull(editFeature.Value(theTaxlotFldIdx)) Then
                Exit Try
            End If

            ' Taxlot has actual taxlot number. ORTaxlot requires a 5-digit number, so leading zeros have to be added.
            Dim theCurrentTaxlotValue As String = CStr(editFeature.Value(theTaxlotFldIdx))
            theCurrentTaxlotValue = AddLeadingZeros(theCurrentTaxlotValue, ORMapNum.GetOrmap_TaxlotFieldLength)

            Dim theNewORTaxlot As String
            theNewORTaxlot = String.Concat(theORMapNumClass.GetORMapNum, theCurrentTaxlotValue)

            Dim theCountyCode As Short = CShort(EditorExtension.DefaultValuesSettings.County)
            Select Case theCountyCode
                Case 1 To 19, 21 To 36
                    editFeature.Value(theMapTaxlotFldIdx) = GenerateMapTaxlotValue(theNewORTaxlot, EditorExtension.TaxLotSettings.MapTaxlotFormatMask)
                Case 20
                    ' 1.  Lane County uses a 2-digit numeric identifier for ranges.
                    '     Special handling is required for east ranges, where 02E is
                    '     stored as 25, 03E as 35, etc.
                    ' 2.  ORMAP standards (OCDES (pg 13); Taxmap Data Model (pg 11)) assert that
                    '     this field should be equal to MAPNUMBER + TAXLOT. In this case, MAPNUMBER
                    '     is already in the right format, thus removing the need for the
                    '     CreateMapTaxlotValue function. Also, in this case, TAXLOT is padded
                    '     on the left with zeros to make it always a 5-digit number (see comment
                    '     above).
                    ' Trim the map number to only the left 8 characters (no spaces)
                    Dim sb As String = theCurrentMapNumber.Trim(CChar(theCurrentMapNumber.Substring(0, 8)))
                    editFeature.Value(theMapTaxlotFldIdx) = String.Concat(sb, theCurrentTaxlotValue)
            End Select

            ' Recalculate ORTaxlot
            If IsDBNull(editFeature.Value(theORTaxlotFldIdx)) Then
                Exit Try
            End If
            ' Get the current and the new ORTaxlot Numbers
            Dim theExistingORTaxlotString As String = CStr(editFeature.Value(theORTaxlotFldIdx))
            Dim theNewORTaxlotString As String = generateORMAPTaxlotNumber(theORMapNumClass.GetORMapNum, editFeature, theExistingORTaxlotString)
            'If no changes, don't save value
            If String.Compare(theExistingORTaxlotString, theNewORTaxlotString, True) <> 0 Then
                editFeature.Value(theORTaxlotFldIdx) = theNewORTaxlotString
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        Finally
            theORMapNumClass = Nothing

        End Try
    End Sub

    ''' <summary>
    ''' Converts a domain descriptive value to the stored code.
    ''' </summary>
    ''' <param name="fields">An field collection object that supports the IFields interface.</param>
    ''' <param name="fieldName">A field that exists in fields.</param>
    ''' <param name="codedValue">A coded name to convert to a coded value</param>
    ''' <returns>A string that represents the domain coded value that 
    ''' corresponds with the coded value name (<paramref name="codedValue"/>), or 
    ''' a empty string.</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvertCodeValueDomainToCode(ByVal fields As IFields, ByVal fieldName As String, ByVal codedValue As String) As String
        Try
            Dim fieldIndex As Integer
            Dim returnValue As String = ""

            If (fields Is Nothing) OrElse Not (TypeOf fields Is IFields) Then
                Return returnValue
            End If

            fieldIndex = fields.FindField(fieldName)
            If fieldIndex > -1 Then
                Dim field As IField
                field = fields.Field(fieldIndex)
                Dim domain As ICodedValueDomain
                domain = DirectCast(field.Domain, ICodedValueDomain)

                If Not (domain Is Nothing) Then
                    If TypeOf domain Is ICodedValueDomain Then
                        Dim thisCodedValueDomain As ICodedValueDomain
                        thisCodedValueDomain = domain
                        For domainIndex As Integer = 0 To thisCodedValueDomain.CodeCount - 1
                            If String.Compare(thisCodedValueDomain.Name(domainIndex), codedValue, True) = 0 Then
                                returnValue = CStr(thisCodedValueDomain.Value(domainIndex))
                            End If
                        Next domainIndex
                    Else
                        returnValue = codedValue 'if range domain return the value
                    End If
                End If 'if domain is not nothing
            End If

            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            ConvertCodeValueDomainToCode = String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Converts a code index to the domain descriptive value.
    ''' </summary>
    ''' <param name="fields">An field collection object that supports the 
    ''' <c>IFields</c> interface.</param>
    ''' <param name="fieldName">A field that exists in fields.</param>
    ''' <param name="codedValue">A coded value to covert to a coded name.</param>
    ''' <returns>A string that represents the domain coded name that 
    ''' corresponds with the coded value name (<paramref name="codedValue"/>), or 
    ''' empty string.</returns>
    ''' <remarks></remarks>
    Public Shared Function ConvertCodeValueDomainToDescription(ByVal fields As IFields, ByVal fieldName As String, ByVal codedValue As String) As String
        Try
            Dim returnValue As String = ""
            Dim fieldIndex As Integer
            fieldIndex = fields.FindField(fieldName)
            If fieldIndex > -1 Then
                Dim field As IField
                field = fields.Field(fieldIndex)

                Dim domain As ICodedValueDomain
                domain = DirectCast(field.Domain, ICodedValueDomain)
                If Not (domain Is Nothing) Then
                    If TypeOf domain Is ICodedValueDomain Then
                        Dim codedValueDomain As ICodedValueDomain
                        codedValueDomain = domain
                        For domainIndex As Integer = 0 To codedValueDomain.CodeCount - 1
                            If codedValueDomain.Name(domainIndex) = codedValue Then
                                returnValue = codedValueDomain.Name(domainIndex)
                            End If
                        Next domainIndex
                    Else
                        returnValue = codedValue 'if range domain return the value
                    End If
                End If 'if domain is valid object
            End If
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Locate a feature layer by its dataset name.
    ''' </summary>
    ''' <param name="datasetName">The name of the dataset to find.</param>
    ''' <returns>A layer object of that supports the IFeatureLayer interface.</returns>
    ''' <remarks>Searches in the TOC recursively (i.e. within group layers). 
    ''' Returns the first feature layer with a matching dataset name.</remarks>
    Public Shared Function FindFeatureLayerByDSName(ByVal datasetName As String) As ESRI.ArcGIS.Carto.IFeatureLayer

        Dim theMap As IMap = EditorExtension.Editor.Map
        Dim thisUID As New UID

        Try
            Dim returnValue As IFeatureLayer = Nothing

            ' Get a reference to the feature layers collection of the document. 
            thisUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
            'We want a EnumLayer containing all FeatureLayer objects.
            Dim theFeatureLayers As IEnumLayer
            theFeatureLayers = theMap.Layers(thisUID, True)

            theFeatureLayers.Reset()
            Dim thisFeatureLayer As IFeatureLayer
            thisFeatureLayer = DirectCast(theFeatureLayers.Next, IFeatureLayer)

            Dim thisDataSet As IDataset

            Do While Not (thisFeatureLayer Is Nothing)
                thisDataSet = DirectCast(thisFeatureLayer.FeatureClass, IDataset)
                If Not (thisDataSet Is Nothing) Then
                    If String.Compare(thisDataSet.Name, datasetName, True) = 0 Then
                        returnValue = DirectCast(thisFeatureLayer, IFeatureLayer)
                        Exit Do
                    End If
                End If
                thisFeatureLayer = DirectCast(theFeatureLayers.Next(), IFeatureLayer)
            Loop

            Return returnValue

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing

        Finally
            thisUID = Nothing
            theMap = Nothing

        End Try

    End Function

    ''' <summary>
    ''' Locate a standalone table by its dataset name.
    ''' </summary>
    ''' <param name="datasetName">The name of the table to find.</param>
    ''' <returns>A table object of that supports the ITable interface.</returns>
    ''' <remarks>Searches in the TOC recursively (i.e. within group layers). 
    ''' Returns the first table with a matching dataset name.</remarks>
    Public Shared Function FindStandaloneTableByDSName(ByVal datasetName As String) As IStandaloneTable

        Dim theMap As IMap = EditorExtension.Editor.Map

        Try
            Dim returnValue As IStandaloneTable = Nothing

            ' Get map as table collection
            Dim theStandaloneTableCollection As IStandaloneTableCollection
            theStandaloneTableCollection = DirectCast(theMap, IStandaloneTableCollection)

            Dim theStandaloneTable As IStandaloneTable
            For t As Integer = 0 To (theStandaloneTableCollection.StandaloneTableCount - 1)
                theStandaloneTable = theStandaloneTableCollection.StandaloneTable(t)
                If theStandaloneTable.Name = datasetName Then
                    returnValue = theStandaloneTable
                    Exit For
                End If
            Next

            Return returnValue

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing

        Finally
            theMap = Nothing

        End Try

    End Function

    ''' <summary>
    ''' Determine the x and y coordinates of the center of envelope, and return them as a Point object.
    ''' </summary>
    ''' <param name="envelope">An envelope object of type IEnvelope.</param>
    ''' <returns>A Point object that represents the center of the envelope.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetCenterOfEnvelope(ByRef envelope As IEnvelope) As IPoint
        Try
            Dim center As IPoint
            center = New Point
            center.X = envelope.XMin + (envelope.XMax - envelope.XMin) / 2
            center.Y = envelope.YMin + (envelope.YMax - envelope.YMin) / 2
            Return center
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Get the MapIndex feature layer.
    ''' </summary>
    ''' <returns>The MapIndex feature layer.</returns>
    ''' <remarks>This feature layer may be named something other than "MapIndex", 
    ''' depending on user settings.</remarks>
    Public Shared Function GetMapIndexFeatureLayer() As IFeatureLayer
        ' Find Map Index feature layer
        Dim theMapIndexFLayer As IFeatureLayer
        With EditorExtension.TableNamesSettings
            ' Find MapIndex feature layer
            theMapIndexFLayer = FindFeatureLayerByDSName(.MapIndexFC)
            If theMapIndexFLayer Is Nothing Then
                Return Nothing
            End If
        End With
        Return theMapIndexFLayer
    End Function

    ''' <summary>
    '''  Validate and format a map suffix number.
    ''' </summary>
    ''' <param name="theFeature">An object that supports the IFeature interface.</param>
    ''' <returns>A string the represents a properly formatted Map Suffix.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetMapSuffixNum(ByVal theFeature As IFeature) As String
        Dim returnValue As New String("0"c, 3)

        Try
            Dim theTaxlotMapSuffixFldIdx As Integer
            theTaxlotMapSuffixFldIdx = LocateFields(DirectCast(theFeature.Class, IFeatureClass), EditorExtension.TaxLotSettings.MapSuffixNumberField)
            If theTaxlotMapSuffixFldIdx > -1 Then
                If Not IsDBNull(theFeature.Value(theTaxlotMapSuffixFldIdx)) Then
                    returnValue = CStr(theFeature.Value(theTaxlotMapSuffixFldIdx))
                End If
                'verify that it is exactly 3 digits
                If returnValue.Length < 3 Then
                    returnValue.PadLeft(3, "0"c)
                End If
                If returnValue.Length > 3 Then
                    returnValue = returnValue.Substring(0, 3)
                End If
            End If
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return "000"
        End Try
    End Function

    ''' <summary>
    ''' Validate and format a map suffix type.
    ''' </summary>
    ''' <param name="theFeature">An object that supports the IFeature interface.</param>
    ''' <returns>A string that represents a properly formatted Map Suffix Type.</returns>
    ''' <remarks>A proper map suffix type is one character.</remarks>
    Public Shared Function GetMapSuffixType(ByRef theFeature As IFeature) As String
        Dim returnValue As New String("0"c, 1)
        Try
            Dim theTaxlotMapTypeFldIdx As Integer
            theTaxlotMapTypeFldIdx = LocateFields(DirectCast(theFeature.Class, IFeatureClass), EditorExtension.TaxLotSettings.MapSuffixTypeField)
            If theTaxlotMapTypeFldIdx > -1 Then
                If Not IsDBNull(theFeature.Value(theTaxlotMapTypeFldIdx)) Then
                    returnValue = CStr(theFeature.Value(theTaxlotMapTypeFldIdx))
                End If
                'verify that it is exactly one digit
                If returnValue.Length < 1 Then
                    returnValue = "0"
                End If
                If returnValue.Length > 1 Then
                    returnValue = returnValue.Substring(0, 1)
                End If
            End If
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return "0"
        End Try
    End Function

    ''' <summary>
    ''' Retrieve a feature-linked annotation feature.
    ''' </summary>
    ''' <param name="anObject">An initialized geodatabase object.</param>
    ''' <returns>An object that supports the <c>IFeature</c> interface.</returns>
    ''' <remarks><para>Finds all related objects to the feature through the 
    ''' first found relationship class, and returns the first related object as 
    ''' the return value.</para>
    ''' <para>This is optimized for annotation because there is a single 
    ''' relationship class.</para></remarks>
    Public Shared Function GetRelatedObjects(ByVal anObject As IObject) As IFeature
        Try
            Dim relationshipClassEnum As IEnumRelationshipClass

            relationshipClassEnum = anObject.Class.RelationshipClasses(esriRelRole.esriRelRoleAny)
            If relationshipClassEnum IsNot Nothing Then
                Dim thisRelationshipClass As IRelationshipClass
                thisRelationshipClass = relationshipClassEnum.Next
                If thisRelationshipClass IsNot Nothing Then
                    Dim parentSet As ISet
                    parentSet = thisRelationshipClass.GetObjectsRelatedToObject(anObject)
                    If parentSet IsNot Nothing Then
                        Dim parentFeature As IFeature
                        parentFeature = DirectCast(parentSet.Next, IFeature)
                        If parentFeature IsNot Nothing Then
                            Return parentFeature
                        End If
                    End If
                End If
            End If
            Return Nothing
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Return a feature cursor of the selected features of the specified layer.
    ''' </summary>
    ''' <param name="layer">The feature layer to return the selection from</param>
    ''' <returns>An object that supports the <c>IFeatureCursor</c> interface.</returns>
    ''' <remarks>References the currently selected features in layer, and 
    ''' returns a feature cursor with the selected features in it.</remarks>
    Public Shared Function GetSelectedFeatures(ByVal layer As IFeatureLayer) As IFeatureCursor
        Try
            If Not TypeOf layer Is IFeatureLayer Then
                Return Nothing
            End If

            Dim thisSelection As IFeatureSelection
            thisSelection = DirectCast(layer, IFeatureSelection)
            Dim returnValue As IFeatureCursor
            Dim thisCursor As ICursor = Nothing
            thisSelection.SelectionSet.Search(Nothing, False, thisCursor)
            returnValue = DirectCast(thisCursor, IFeatureCursor)
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Get the Taxlot feature layer.
    ''' </summary>
    ''' <returns>The Taxlot feature layer.</returns>
    ''' <remarks>This feature layer may be named something other than "Taxlot", 
    ''' depending on user settings.</remarks>
    Public Shared Function GetTaxlotFeatureLayer() As IFeatureLayer
        ' Find Taxlot feature layer
        Dim theTaxlotFLayer As IFeatureLayer
        With EditorExtension.TableNamesSettings
            ' Find Taxlot feature layer
            theTaxlotFLayer = FindFeatureLayerByDSName(.TaxLotFC)
            If theTaxlotFLayer Is Nothing Then
                Return Nothing
            End If
        End With
        Return theTaxlotFLayer
    End Function

    ' TODO: [NIS] Refactor (smaller modules).
    ''' <summary>
    ''' Overlay the passed in feature with a feature class to get a value from 
    ''' the specified field.
    ''' </summary>
    ''' <param name="theGeometry">The search geometry.</param>
    ''' <param name="overlayFeatureClass">Overlaying feature class.</param>
    ''' <param name="valueFieldName">Name of field to return value for.</param>
    ''' <param name="orderBestByFieldName">Name of field to order by in the 
    ''' case of a tie in area/length of intersection.</param>
    ''' <returns>Returns the value from the specified field (<paramref>valueFieldName</paramref>) as a string.</returns>
    ''' <remarks>Gets the target feature with the largest area of intersection 
    ''' with the geometry and gets its value from the field, or, if tied (unikely 
    ''' but possible), then gets the best (lowest) value from the field, based 
    ''' on the order by field value.</remarks>
    Public Shared Function GetValueViaOverlay(ByRef theGeometry As IGeometry, ByRef overlayFeatureClass As IFeatureClass, ByVal valueFieldName As String, Optional ByVal orderBestByFieldName As String = "") As String
        Try
            Dim continueThisProcess As Boolean

            continueThisProcess = True 'initialize

            If (theGeometry Is Nothing) OrElse (overlayFeatureClass Is Nothing) OrElse (valueFieldName.Length <= 0) Then
                ' TODO: [NIS] Add assertions here
                continueThisProcess = False
            End If

            Dim valueFieldIndex As Integer = FieldNotFoundIndex
            If continueThisProcess Then
                valueFieldIndex = overlayFeatureClass.Fields.FindField(valueFieldName)
                If valueFieldIndex = FieldNotFoundIndex Then
                    continueThisProcess = False
                End If
            End If

            Dim orderBestByFieldIndex As Integer = FieldNotFoundIndex
            If continueThisProcess Then
                If orderBestByFieldName.Length = 0 Then
                    ' Use the value field as the order-by field
                    orderBestByFieldIndex = valueFieldIndex
                Else
                    orderBestByFieldIndex = overlayFeatureClass.Fields.FindField(orderBestByFieldName)
                    If orderBestByFieldIndex = FieldNotFoundIndex Then
                        ' Field not found. Try the OID field
                        If overlayFeatureClass.HasOID Then
                            orderBestByFieldIndex = overlayFeatureClass.Fields.FindField(overlayFeatureClass.OIDFieldName)
                            Dim msg As String = "Field " & orderBestByFieldName & " not found in " & overlayFeatureClass.AliasName & vbNewLine & _
                                    ". Using " & overlayFeatureClass.OIDFieldName
                            Debug.WriteLine(msg)
                            Trace.WriteLine(msg)
                        Else
                            continueThisProcess = False
                        End If
                    End If
                End If
            End If

            Dim thisPolygon As IPolygon
            Dim thisArea As IArea
            Dim thisCurve As ICurve
            Dim thisEnvelope As IEnvelope
            Dim intersectFuzzAmount As Double
            Const fuzzFactor As Double = 0.05 ' TODO: [NIS] Re-implement as user setting?

            If continueThisProcess Then
                Select Case theGeometry.GeometryType
                    Case esriGeometryType.esriGeometryPolygon
                        thisPolygon = DirectCast(theGeometry, IPolygon)
                        thisArea = DirectCast(thisPolygon, IArea)
                        intersectFuzzAmount = thisArea.Area * fuzzFactor

                    Case esriGeometryType.esriGeometryPolyline, esriGeometryType.esriGeometryLine, esriGeometryType.esriGeometryBezier3Curve, esriGeometryType.esriGeometryCircularArc, esriGeometryType.esriGeometryEllipticArc, esriGeometryType.esriGeometryPath
                        thisCurve = DirectCast(theGeometry, ICurve)
                        intersectFuzzAmount = thisCurve.Length * fuzzFactor

                    Case esriGeometryType.esriGeometryEnvelope
                        thisEnvelope = DirectCast(theGeometry, IEnvelope)
                        thisPolygon = envelopeToPolygon(thisEnvelope)
                        thisArea = DirectCast(thisPolygon, IArea)
                        intersectFuzzAmount = thisArea.Area * fuzzFactor

                    Case Else
                        continueThisProcess = False

                End Select
            End If

            Dim theOverlayFeatureCursor As IFeatureCursor
            Dim anOverlayFeature As IFeature
            Dim dictCandidates As New Dictionary(Of Integer, Double) '(key::value) OID::area/length

            If continueThisProcess Then

                Dim largestIntersectArea As Double = 0
                Dim longestIntersectLength As Double = 0

                theOverlayFeatureCursor = doSpatialQuery(overlayFeatureClass, theGeometry, esriSpatialRelEnum.esriSpatialRelIntersects)
                If theOverlayFeatureCursor IsNot Nothing Then
                    anOverlayFeature = theOverlayFeatureCursor.NextFeature
                    Dim topoOperator As ITopologicalOperator
                    Dim intersectGeometry As IGeometry
                    Dim intersectArea As Double = 0
                    dictCandidates = New Dictionary(Of Integer, Double)
                    While Not (anOverlayFeature Is Nothing)

                        topoOperator = DirectCast(anOverlayFeature.Shape, ITopologicalOperator)
                        If Not topoOperator.IsSimple Then
                            topoOperator.Simplify()
                        End If

                        Select Case theGeometry.GeometryType

                            Case esriGeometryType.esriGeometryEnvelope, esriGeometryType.esriGeometryPolygon
                                ' Determine if the target feature has the largest area of intersection with the
                                ' current source feature. Set flags used below.
                                intersectGeometry = topoOperator.Intersect(theGeometry, esriGeometryDimension.esriGeometry2Dimension)
                                If Not intersectGeometry.IsEmpty Then
                                    thisPolygon = DirectCast(intersectGeometry, IPolygon)
                                    thisArea = DirectCast(thisPolygon, IArea)
                                    intersectArea = thisArea.Area
                                    If System.Math.Abs(intersectArea - largestIntersectArea) > intersectFuzzAmount Then
                                        '[Difference greater than fuzz amount, so not a "fuzzy tie"...]
                                        ' Determine if this is the feature with the largest area of intersection.
                                        If intersectArea > largestIntersectArea Then
                                            largestIntersectArea = intersectArea
                                            '[New largest intersection...]
                                            ' Get the value only for this feature.
                                            dictCandidates.Clear()
                                            dictCandidates.Add(anOverlayFeature.OID, largestIntersectArea)
                                        Else
                                            '[Smaller intersection...]
                                            ' Don't get the value for this feature.
                                        End If
                                    Else
                                        '[Difference not greater than fuzz amount, so a "fuzzy tie"...]
                                        ' Evaluate this feature against other tied candidates (don't clear dictionary).
                                        dictCandidates.Add(anOverlayFeature.OID, largestIntersectArea)
                                    End If
                                Else
                                    '[Empty intersection geometry...]
                                    ' Don't get the value for this feature
                                End If

                            Case esriGeometryType.esriGeometryLine, esriGeometryType.esriGeometryPolyline
                                ' Determine if the target feature has the longest length of intersection with the
                                ' current source feature. Set flags used below.
                                intersectGeometry = topoOperator.Intersect(theGeometry, esriGeometryDimension.esriGeometry1Dimension)
                                If Not intersectGeometry.IsEmpty Then
                                    thisCurve = DirectCast(intersectGeometry, ICurve)
                                    Dim intersectLength As Double
                                    intersectLength = thisCurve.Length
                                    If System.Math.Abs(intersectLength - longestIntersectLength) > intersectFuzzAmount Then
                                        '[Difference greater than fuzz amount, so not a "fuzzy tie"...]
                                        ' Determine if this is the feature with the longest length of intersection
                                        If intersectLength > longestIntersectLength Then
                                            longestIntersectLength = intersectLength
                                            dictCandidates.Clear()
                                            dictCandidates.Add(anOverlayFeature.OID, longestIntersectLength)
                                        Else
                                            '[Smaller intersection...]
                                            ' Don't get the value for this feature
                                        End If
                                    Else
                                        '[Difference not greater than fuzz amount, so a "fuzzy tie"...]
                                        'Evaluate this feature against other tied candidates (don't clear dictionary).
                                        dictCandidates.Add(anOverlayFeature.OID, longestIntersectLength)
                                    End If
                                Else
                                    '[Empty intersection geometry...]
                                    ' Don't get the value for this feature
                                End If

                            Case esriGeometryType.esriGeometryPoint
                                '[Tied by definition (0-dimension = zero area & length = tie)...]
                                ' Evaluate this feature against other tied candidates (don't clear dictionary).
                                dictCandidates.Add(anOverlayFeature.OID, 0)

                            Case Else
                                continueThisProcess = False

                        End Select
                        anOverlayFeature = theOverlayFeatureCursor.NextFeature
                    End While
                End If

            End If

            Dim theBestValue As String = ""

            If continueThisProcess Then

                If dictCandidates.Count > 0 Then
                    Dim aValue As String = ""
                    Dim theBestOrderByValue As String = ""

                    Dim whereClause As String = String.Concat(overlayFeatureClass.OIDFieldName, " in (", candidateKeysToDelimitedString(dictCandidates), ")")
                    theOverlayFeatureCursor = doSpatialQuery(overlayFeatureClass, theGeometry, esriSpatialRelEnum.esriSpatialRelIntersects, whereClause)
                    If Not (theOverlayFeatureCursor Is Nothing) Then
                        anOverlayFeature = theOverlayFeatureCursor.NextFeature
                        While Not (anOverlayFeature Is Nothing)
                            If dictCandidates.Count > 1 Then
                                '[Candidates tied for area/length of intersection (unikely but possible)...]
                                ' Get the value only if the candidate source feature has the best (lowest) order-by value
                                Dim anOrderByValue As String = CStr(anOverlayFeature.Value(orderBestByFieldIndex))
                                ' Note: When you compare strings, the string expressions are evaluated based on their 
                                ' alphabetical sort order (e.g. "A" < "B" and "A" < "Ax"), which depends on the Option 
                                ' Compare setting.
                                If theBestOrderByValue.Length = 0 OrElse anOrderByValue < theBestOrderByValue Then
                                    '[Any new value beats an empty value...]
                                    '[lower in the sort order is assumed to be better...]
                                    theBestOrderByValue = anOrderByValue
                                    If Not IsDBNull(anOverlayFeature.Value(valueFieldIndex)) Then
                                        aValue = CStr(anOverlayFeature.Value(valueFieldIndex))
                                        theBestValue = aValue
                                    Else
                                        MessageBox.Show("The field " & anOverlayFeature.Fields.Field(valueFieldIndex).Name & " contains a Null" & vbNewLine & "for the feature with OID=" & anOverlayFeature.OID & ".", "Error: GetValueViaOverlay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                    End If
                                End If
                            Else
                                '[Candidates not tied (only one)...]
                                ' Get the value for the only candidate feature
                                If Not IsDBNull(anOverlayFeature.Value(valueFieldIndex)) Then
                                    aValue = CStr(anOverlayFeature.Value(valueFieldIndex))
                                    theBestValue = aValue
                                Else
                                    MessageBox.Show("The field " & anOverlayFeature.Fields.Field(valueFieldIndex).Name & " contains a Null" & vbNewLine & "for the feature with OID=" & anOverlayFeature.OID & ".", "Error: GetValueViaOverlay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                End If
                                Exit While
                            End If 'dictCandidates.count > 1
                        End While
                    End If 'featurecursor is not nothing
                Else
                    MessageBox.Show("No " & overlayFeatureClass.AliasName & "features found " & vbNewLine & " which intersect this feature.", "Error: GetValueViaOverlay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If 'dictionary count > zero
            End If 'continueThisProcess 

            Return theBestValue

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Finds a feature layer in the table of contents corresponding to the given feature class name
    ''' and determines if the feature class has the listed fields.
    ''' </summary>
    ''' <param name="featureClassName">The name of the feature class to inspect (if found).</param>
    ''' <param name="fieldNames">A collection of field names to search for.</param>
    ''' <param name="loadData">A boolean flag. If true, if the feature class is not found, 
    ''' an attempt is made to find the data and create a layer in the map.</param>
    ''' <returns><c>True</c> or <c>False</c>.</returns>
    ''' <remarks></remarks>
    Public Shared Function FeatureClassHasRequiredFields(ByVal featureClassName As String, ByVal fieldNames As List(Of String), ByVal loadData As Boolean) As Boolean

        Dim returnValue As Boolean = True 'initial assumption

        ' Confirm data layer is present in current map
        Dim theFLayer As IFeatureLayer

        theFLayer = FindFeatureLayerByDSName(featureClassName)

        If theFLayer Is Nothing Then
            If loadData Then
                '[Load option accepted...]
                ' Attempt to load and find the taxlot layer in the map document
                If LoadFCIntoMap(featureClassName) Then
                    '[Layer loaded...]
                    theFLayer = FindFeatureLayerByDSName(featureClassName)
                Else
                    '[Unable to load the layer...]
                    returnValue = False
                End If
            Else
                '[Data not present and load option refused...]
                returnValue = False
            End If
        End If

        If theFLayer IsNot Nothing Then
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
        End If

        Return returnValue

    End Function

    Public Shared Function TableHasRequiredFields(ByVal tableName As String, ByVal fieldNames As List(Of String), ByVal loadData As Boolean) As Boolean

        ' TODO: [NIS] Test this...

        Dim returnValue As Boolean = True 'initial assumption

        ' Confirm data layer is present in current map
        Dim theTable As ITable

        theTable = FindStandaloneTableByDSName(tableName).Table

        If theTable Is Nothing Then
            If loadData Then
                '[Load option accepted...]
                ' Attempt to load and find the table in the map document
                If LoadTableIntoMap(tableName) Then
                    '[Table loaded...]
                    theTable = FindStandaloneTableByDSName(tableName).Table
                Else
                    '[Unable to load the table...]
                    returnValue = False
                End If
            Else
                '[Data not present and load option refused...]
                returnValue = False
            End If
        End If

        If theTable IsNot Nothing Then
            ' Confirm fields are present
            Dim foundAllFields As Boolean = True  'initial assumption
            Dim fieldIndex As Integer
            For Each fn As String In fieldNames
                fieldIndex = theTable.FindField(fn)
                If fieldIndex <> FieldNotFoundIndex Then
                    foundAllFields = True
                Else
                    foundAllFields = False
                    Exit For
                End If
            Next fn
            returnValue = foundAllFields
        End If

        Return returnValue

    End Function

    ''' <summary>
    ''' Determines if the feature layer has a selection
    ''' </summary>
    ''' <param name="layer">An object that supports the IFeatureLayer</param>
    ''' <returns><c>True</c> or <c>False</c>.</returns>
    ''' <remarks>Checking the selection set of layer, determine if one, many, or no features are selected.</remarks>
    Public Shared Function HasSelectedFeatures(ByVal layer As IFeatureLayer) As Boolean
        Try
            Dim returnValue As Boolean = False
            If (layer Is Nothing) Or Not (TypeOf layer Is IFeatureLayer) Then
                Return returnValue
            End If

            ' How many are selected?
            Dim featuresSelected As IFeatureSelection
            featuresSelected = DirectCast(layer, IFeatureSelection)
            returnValue = (featuresSelected.SelectionSet.Count > 0)
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determine if a feature is annotation.
    ''' </summary>
    ''' <param name="thisObject">A valid initialized geodatabase object.</param>
    ''' <returns><c>True</c> or <c>False</c>.</returns>
    ''' <remarks>Compares the feature type of <paramref name="thisObject"/> with 
    ''' that of annotation and return the truth value of the comparison.</remarks>
    Public Shared Function IsAnno(ByVal thisObject As IObject) As Boolean
        Try
            Dim thisObjectClass As IObjectClass
            thisObjectClass = thisObject.Class

            If TypeOf thisObject Is IFeature Then
                Dim thisFeatureClass As IFeatureClass
                thisFeatureClass = DirectCast(thisObjectClass, IFeatureClass)
                If thisFeatureClass.FeatureType = esriFeatureType.esriFTAnnotation Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determine if the feature belongs to the MapIndex feature class.
    ''' </summary>
    ''' <param name="thisObject">A valid initialized geodatabase object.</param>
    ''' <returns><c>True</c> or <c>False</c>.</returns>
    ''' <remarks>Compares the name of the dataset of the <paramref name="thisObject"/> 
    ''' feature class to the Map Index layer name in order to determine if it 
    ''' is the MapIndex feature class.</remarks>
    Public Shared Function IsMapIndex(ByVal thisObject As IObject) As Boolean
        Try
            Dim thisObjectClass As IObjectClass
            Dim thisDataset As IDataset
            thisObjectClass = thisObject.Class
            thisDataset = DirectCast(thisObjectClass, IDataset)
            If String.Compare(thisDataset.Name, EditorExtension.TableNamesSettings.MapIndexFC, True) = 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determines if a feature class is part of the ORMAP design.
    ''' </summary>
    ''' <param name="theObject">A valid initialized geodatabase object</param>
    ''' <returns>True or False</returns>
    ''' <remarks></remarks>
    Public Shared Function IsORMAPFeature(ByRef theObject As IObject) As Boolean
        Try
            Dim returnValue As Boolean = False
            Dim thisObjectClass As IObjectClass
            thisObjectClass = theObject.Class
            'TODO: JWM REFACTOR RETURN ENUM OF TYPE OF FEATURE CLASS
            Dim thisDataset As IDataset
            thisDataset = DirectCast(thisObjectClass, IDataset)
            Dim datasetName As String = thisDataset.Name
            Const StringMatch As Integer = 0
            ' Check for a match against any of the ORMAP feature classes.
            returnValue = (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.Anno0010scaleFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.Anno0020scaleFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.Anno0030scaleFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.Anno0040scaleFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.Anno0050scaleFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.Anno0100scaleFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.Anno0200scaleFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.Anno0400scaleFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.Anno0800scaleFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.Anno2000scaleFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.TaxCodeAnnoFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.TaxlotAcreageAnnoFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.AnnoTableNamesSettings.TaxlotNumberAnnoFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.TableNamesSettings.CartographicLinesFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.TableNamesSettings.TaxLotFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.TableNamesSettings.MapIndexFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.TableNamesSettings.PlatsFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.TableNamesSettings.ReferenceLinesFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.TableNamesSettings.TaxCodeFC, True) = StringMatch)
            returnValue = returnValue OrElse (String.Compare(datasetName, EditorExtension.TableNamesSettings.TaxLotLinesFC, True) = StringMatch)
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determine if the feature belongs to the Taxlot feature class.
    ''' </summary>
    ''' <param name="thisObject">A valid initialized geodatabase object.</param>
    ''' <returns><c>True</c> or <c>False</c>.</returns>
    ''' <remarks>Determine if thisObject belongs to the Taxlot feature class by 
    ''' checking the name of the dataset of thisObject feature class against 
    ''' the Taxlot Feature Class constant.</remarks>
    Public Shared Function IsTaxlot(ByVal thisObject As IObject) As Boolean
        Try
            Dim thisObjectClass As IObjectClass
            Dim thisDataset As IDataset
            thisObjectClass = thisObject.Class
            thisDataset = DirectCast(thisObjectClass, IDataset)
            If String.Compare(thisDataset.Name, EditorExtension.TableNamesSettings.TaxLotFC, True) = 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Loads the specified feature class into the current map as a feature layer.
    ''' </summary>
    ''' <param name="featureClassName">The feature class to find.</param>
    ''' <param name="title">An alternate title for the file dialog box.</param>
    ''' <returns><c>True</c> for found and loaded, <c>False</c> for not found and loaded.</returns>
    ''' <remarks>Show a dialog box with title that allows the user to select the 
    ''' personal geodatabase that the <paramref name="featureClassName"/> resides in. 
    ''' The feature class is then loaded in the current map from the chosen personal 
    ''' geodatabase.</remarks>
    Public Shared Function LoadFCIntoMap(ByVal featureClassName As String, Optional ByVal title As String = "") As Boolean
        Try
            Dim thisFileDialog As CatalogFileDialog
            thisFileDialog = New CatalogFileDialog()

            With thisFileDialog
                .SetAllowMultiSelect(True)
                .SetButtonCaption("Select")
                If title.Length > 0 Then
                    .SetTitle(title)
                Else
                    .SetTitle(String.Concat("Find feature class ", featureClassName, "..."))
                End If
                .SetFilter(New ESRI.ArcGIS.Catalog.GxFilterPersonalGeodatabases, True, True)
                .ShowOpen()
            End With

            'exit if there is nothing selected
            If thisFileDialog.SelectedObject(1) Is Nothing Then
                Return False
            End If

            Dim thisWorkspaceFactory As IWorkspaceFactory2
            thisWorkspaceFactory = New AccessWorkspaceFactory
            Dim thisWorkSpace As IWorkspace
            thisWorkSpace = thisWorkspaceFactory.OpenFromFile(CStr(thisFileDialog.SelectedObject(1)), 0)

            Dim thisFeatureWorkspace As IFeatureWorkspace
            thisFeatureWorkspace = DirectCast(thisWorkSpace, IFeatureWorkspace)

            Dim thisFeatureClass As IFeatureClass
            thisFeatureClass = thisFeatureWorkspace.OpenFeatureClass(featureClassName)

            Dim thisFeatureLayer As New ESRI.ArcGIS.Carto.FeatureLayer
            thisFeatureLayer.FeatureClass = thisFeatureClass

            Dim thisDataSet As IDataset
            thisDataSet = DirectCast(thisFeatureClass, IDataset)
            thisFeatureLayer.Name = thisDataSet.Name

            Dim thisMap As ESRI.ArcGIS.Carto.IMap
            thisMap = EditorExtension.Editor.Map
            thisMap.AddLayer(thisFeatureLayer)

            Dim thisArcMapDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
            thisArcMapDoc = DirectCast(EditorExtension.Application.Document, IMxDocument)
            thisArcMapDoc.CurrentContentsView.Refresh(0)

            Return True

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False

        End Try

    End Function

    ''' <summary>
    ''' Loads the specified object class (table) into the current map.
    ''' </summary>
    ''' <param name="objectClassName">The object class (table) to find.</param>
    ''' <param name="title">An alternate title for the file dialog box.</param>
    ''' <returns><c>True</c> for found and loaded, <c>False</c> for not found and loaded.</returns>
    ''' <remarks>Show a dialog box with title that allows the user to select the 
    ''' personal geodatabase that the <paramref name="objectClassName"/> resides in. 
    ''' The object class is then loaded in the current map from the chosen personal 
    ''' geodatabase.</remarks>
    Public Shared Function LoadTableIntoMap(ByVal objectClassName As String, Optional ByVal title As String = "") As Boolean

        ' TODO: [NIS] TEST THIS...

        Try
            Dim theFileDialog As CatalogFileDialog
            theFileDialog = New CatalogFileDialog()

            With theFileDialog
                .SetAllowMultiSelect(True)
                .SetButtonCaption("Select")
                If title.Length > 0 Then
                    .SetTitle(title)
                Else
                    .SetTitle(String.Concat("Find object class (table) ", objectClassName, "..."))
                End If
                .SetFilter(New ESRI.ArcGIS.Catalog.GxFilterPersonalGeodatabases, True, True)
                .ShowOpen()
            End With

            ' Exit if there is nothing selected
            If theFileDialog.SelectedObject(1) Is Nothing Then
                Return False
            End If

            Dim theWorkspaceFactory As IWorkspaceFactory2
            theWorkspaceFactory = New AccessWorkspaceFactory
            Dim theWorkSpace As IWorkspace
            theWorkSpace = theWorkspaceFactory.OpenFromFile(CStr(theFileDialog.SelectedObject(1)), 0)

            Dim theFeatureWorkspace As IFeatureWorkspace
            theFeatureWorkspace = DirectCast(theWorkSpace, IFeatureWorkspace)

            ' Open the table
            Dim theTable As Geodatabase.ITable
            theTable = theFeatureWorkspace.OpenTable(objectClassName)

            Dim theMap As ESRI.ArcGIS.Carto.IMap
            Dim theArcMapDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
            theMap = EditorExtension.Editor.Map
            theArcMapDoc = DirectCast(EditorExtension.Application.Document, IMxDocument)

            ' Create a table collection and assign the new table to it
            Dim theStandaloneTable As IStandaloneTable
            Dim theStandaloneTableCollection As IStandaloneTableCollection
            theStandaloneTable = New StandaloneTable
            theStandaloneTable.Table = theTable
            theStandaloneTableCollection = DirectCast(theMap, IStandaloneTableCollection)
            theStandaloneTableCollection.AddStandaloneTable(theStandaloneTable)

            ' Create a new table window for the table
            Dim theTableWindow As ITableWindow
            theTableWindow = New TableWindow
            theTableWindow.Table = theTable
            theTableWindow.ShowAliasNamesInColumnHeadings = True
            theTableWindow.Application = EditorExtension.Application

            ' Update the document
            theArcMapDoc.UpdateContents()

            Return True

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False

        End Try

    End Function

    ''' <summary>
    ''' Find the index of a field in a feature class.
    ''' </summary>
    ''' <param name="featureClass">The feature class to locate a field in.</param>
    ''' <param name="fieldName">The name of the field to find.</param>
    ''' <returns>Index of field or -1.</returns>
    ''' <remarks>This function may return zero because that is a valid index, but -1 is not. The return value of -1 means the field was not found.</remarks>
    Public Shared Function LocateFields(ByRef featureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass, ByRef fieldName As String) As Integer
        Try
            Dim returnValue As Integer
            returnValue = featureClass.Fields.FindField(fieldName)
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Reads a value from a row, given a field name.
    ''' </summary>
    ''' <param name="aRow">An object that implements the IRow interface.</param>
    ''' <param name="fieldName">A field that exists in row.</param>
    ''' <param name="dataType">A string value indicating data type of the field.</param>
    ''' <returns></returns>
    ''' <remarks>Reads the value of a field with a domain and translates the value from the coded value to the coded name.</remarks>
    Public Shared Function ReadValue(ByRef aRow As IRow, ByVal fieldName As String, Optional ByVal dataType As String = "") As String
        Try
            Dim fieldIndex As Integer
            Dim returnValue As String = ""

            fieldIndex = aRow.Fields.FindField(fieldName)
            If fieldIndex > -1 Then
                If String.Compare(dataType, "date", True) = 0 Then
                    If IsDBNull(aRow.Value(fieldIndex)) Then
                        returnValue = CStr(System.DateTime.Today)
                    Else
                        returnValue = CStr(aRow.Value(fieldIndex))
                    End If
                Else
                    If IsDBNull(aRow.Value(fieldIndex)) Then
                        returnValue = String.Empty
                    Else
                        returnValue = CStr(aRow.Value(fieldIndex))
                    End If
                End If
                'Determine if a Domain Field
                Dim field As IField
                field = aRow.Fields.Field(fieldIndex)
                Dim domain As IDomain
                domain = field.Domain
                If Not domain Is Nothing Then
                    If domain.Type = esriDomainType.esriDTCodedValue Then
                        'If TypeOf domain Is ICodedValueDomain Then
                        Dim thisCodedValueDomain As ICodedValueDomain
                        thisCodedValueDomain = DirectCast(domain, ICodedValueDomain)
                        Dim domainValue As Object
                        domainValue = aRow.Value(fieldIndex)
                        'search domain for the code
                        For domainIndex As Integer = 0 To thisCodedValueDomain.CodeCount - 1
                            If thisCodedValueDomain.Value(domainIndex).ToString = domainValue.ToString Then 'TODO: [NIS] Confirm that ToString will work here
                                returnValue = thisCodedValueDomain.Name(domainIndex)
                                Exit For
                            End If
                        Next domainIndex
                    End If
                End If
            End If

            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Update/Initialize feature linked annotation size.
    ''' </summary>
    ''' <param name="theObject">A valid initialized geodatabase object.</param>
    ''' <param name="theFeature">The feature associated with the annotation.</param>
    ''' <remarks> Given an object, <paramref name="theObject">theObject</paramref>, and a feature, <paramref name="theFeature">theFeature</paramref>.
    ''' Determines if <paramref name="theObject">theObject</paramref> is an annotation feature, derives the map number for the Map Index polygon overlaying <paramref name="theFeature">theFeature</paramref>, and
    ''' resets the annotation feature size in <paramref name="theObject">theObject</paramref></remarks>
    Public Shared Sub SetAnnoSize(ByRef theObject As IObject, ByRef theFeature As IFeature)
        Try
            Dim annoObjectClass As IObjectClass
            annoObjectClass = theObject.Class
            Dim annoFeature As IFeature
            annoFeature = DirectCast(theObject, IFeature)

            'Capture MapNumber for each anno feature created
            Dim annoMapNumField As Integer = LocateFields(DirectCast(theObject.Class, IFeatureClass), EditorExtension.TaxLotSettings.MapNumberField)
            If annoMapNumField = FieldNotFoundIndex Then
                Exit Try
            End If

            Dim fieldIndex As Integer = annoFeature.Fields.FindField("TextString")
            If fieldIndex = FieldNotFoundIndex Then
                MessageBox.Show("Unable to locate text string field in annotation class. Cannot set size", "Cannot set size", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Try
            End If

            Dim thisValue As Object
            thisValue = annoFeature.Value(fieldIndex)
            If IsDBNull(thisValue) Then
                Exit Try
            End If

            theFeature = DirectCast(theObject, IFeature)
            Dim thisGeometry As IGeometry
            thisGeometry = theFeature.Shape
            If thisGeometry.IsEmpty Then
                Exit Try
            End If

            Dim thisEnvelope As IEnvelope
            thisEnvelope = thisGeometry.Envelope

            'Dim thisCenter As IPoint
            'thisCenter = GetCenterOfEnvelope(thisEnvelope)

            Dim mapIndexFeatureLayer As IFeatureLayer
            mapIndexFeatureLayer = FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.MapIndexFC)
            If mapIndexFeatureLayer Is Nothing Then
                Exit Try
            End If

            Dim mapIndexFeatureClass As IFeatureClass
            mapIndexFeatureClass = mapIndexFeatureLayer.FeatureClass
            'original vb6 code placed the point object as the first parameter. 
            Dim mapNumber As String = GetValueViaOverlay(thisGeometry, mapIndexFeatureClass, EditorExtension.MapIndexSettings.MapNumberField, EditorExtension.MapIndexSettings.MapNumberField)

            ' Allow existing anno to be moved without changing MapNumber
            ' Some anno will reside in another Taxlot, but labels the neighboring taxlot
            If String.Compare(mapNumber, CStr(theObject.Value(annoMapNumField)), True) = 0 Then
                ' Sets the value of the annotation map number field
                theObject.Value(annoMapNumField) = mapNumber
                ' Update the size to reflect current mapscale
                Dim mapScale As String = GetValueViaOverlay(thisGeometry, mapIndexFeatureClass, EditorExtension.MapIndexSettings.MapScaleField, EditorExtension.MapIndexSettings.MapNumberField)
                If mapScale.Length = 0 Then
                    Exit Try
                End If
                ' Determine which annotation class this is
                Dim annoClass As IObjectClass
                annoClass = theObject.Class

                Dim annoDataSet As IDataset
                annoDataSet = DirectCast(annoClass, IDataset)

                'If other anno, don't continue
                If String.Compare(annoDataSet.Name, EditorExtension.AnnoTableNamesSettings.TaxlotNumberAnnoFC) <> 0 And String.Compare(annoDataSet.Name, EditorExtension.AnnoTableNamesSettings.TaxlotAcreageAnnoFC, True) <> 0 Then
                    Exit Try
                End If

                ' Gets the size of the annotation from the scale of the annotation dataset
                Dim annotationSize As Double = getAnnoSizeByScale(annoDataSet.Name, CInt(mapScale))

                Dim annoFeature2 As IAnnotationFeature
                annoFeature2 = DirectCast(theObject, IAnnotationFeature)

                Dim annoElement As IAnnotationElement
                annoElement = DirectCast(annoFeature2.Annotation, IAnnotationElement)

                Dim cartoElement As IElement
                cartoElement = DirectCast(annoElement, IElement)

                Dim thisTextElement As ITextElement
                thisTextElement = DirectCast(cartoElement, ITextElement)

                Dim thisTextSymbol As ESRI.ArcGIS.Display.ITextSymbol
                thisTextSymbol = thisTextElement.Symbol

                thisTextSymbol.Size = annotationSize
                thisTextElement.Symbol = thisTextSymbol
                cartoElement = DirectCast(thisTextElement, IElement)
                annoElement = DirectCast(cartoElement, IAnnotationElement)
                annoFeature2.Annotation = DirectCast(annoElement, IElement)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Updates Auto fields in a feature class.
    ''' </summary>
    ''' <param name="feature">An object that implements the Ifeature interface.</param>
    ''' <remarks>Update the AutoWho and the AutoDate fields with the current username and date/time, respectively.</remarks>
    Public Shared Sub UpdateMinimumAutoFields(ByRef feature As IFeature)
        Try
            If feature Is Nothing Then
                Exit Try
            End If

            Dim theAutoDateFldIdx As Integer
            theAutoDateFldIdx = feature.Fields.FindField(EditorExtension.AllTablesSettings.AutoDateField)
            If theAutoDateFldIdx > FieldNotFoundIndex Then
                feature.Value(theAutoDateFldIdx) = System.DateTime.Now
            End If

            Dim theAutoWhoFldIdx As Integer
            theAutoWhoFldIdx = feature.Fields.FindField(EditorExtension.AllTablesSettings.AutoWhoField)
            If theAutoWhoFldIdx > FieldNotFoundIndex Then
                feature.Value(theAutoWhoFldIdx) = GetUserName()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
    End Sub

    ''' <summary>
    ''' Determine the validity of a taxlot number.
    ''' </summary>
    ''' <param name="taxlotNumber">The taxlot value to validate.</param>
    ''' <param name="thisGeometry">The geometry of the feature to check.</param>
    ''' <returns>True or False</returns>
    ''' <remarks>Determine if the feature represented by thisGeometry with taxlot taxlotNumber is a unique and therefore valid.</remarks>
    Public Shared Function ValidateTaxlotNumber(ByVal taxlotNumber As String, ByRef thisGeometry As IGeometry) As Boolean
        Try
            Dim returnValue As Boolean = False

            'check for existence of Taxlot layer
            Dim thisTaxlotFeatureLayer As IFeatureLayer
            thisTaxlotFeatureLayer = FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.TaxLotFC)
            If thisTaxlotFeatureLayer Is Nothing Then 'TODO: JWM Place strings in resource file and may use for different type of notification
                MessageBox.Show("Unable to locate the Taxlot layer in Table of Contents." & vbNewLine & _
                                "This process requires a feature class called " & EditorExtension.TableNamesSettings.TaxLotFC & ".", _
                                String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Return returnValue
            End If

            'check for existence of Map index layer
            Dim thisTaxlotFeatureClass As IFeatureClass
            thisTaxlotFeatureClass = thisTaxlotFeatureLayer.FeatureClass

            Dim mapIndexFeatureLayer As IFeatureLayer
            mapIndexFeatureLayer = FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.MapIndexFC)
            If mapIndexFeatureLayer Is Nothing Then
                MessageBox.Show("Unable to locate the MapIndex layer in Table of Contents." & vbNewLine & _
                                "This process requires a feature class called " & EditorExtension.TableNamesSettings.MapIndexFC & ".", _
                                String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Return returnValue
            End If

            Dim mapIndexFeatureClass As IFeatureClass
            mapIndexFeatureClass = mapIndexFeatureLayer.FeatureClass

            ' Checks for the existence of a current ORMAP Number and Taxlot number
            Dim mapIndexORMAPValue As String = String.Empty
            'mapIndexORMAPValue= getValueViaOverlay() 'TODO: JWM Flesh getValueViaOverlay function out
            If mapIndexORMAPValue.Length = 0 Then
                returnValue = True
            End If

            'Make sure this number is unique within taxlots with this OM number
            'TODO: JWM check these EditorExtension values
            Dim whereClause As String = String.Concat(EditorExtension.TaxLotSettings.MapNumberField, "='", mapIndexORMAPValue, "' AND ", EditorExtension.TaxLotSettings.TaxlotField, " = '", taxlotNumber, "'")
            Dim cursor As ICursor
            cursor = attributeQuery(DirectCast(thisTaxlotFeatureClass, ITable), whereClause)
            If Not (cursor Is Nothing) AndAlso (returnValue = False) Then
                Dim row As IRow
                row = cursor.NextRow
                If row Is Nothing Then
                    returnValue = True
                End If
            End If

            Return returnValue

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False
        End Try
    End Function

    <ObsoleteAttribute("Use the ZoomToEnvelope() function instead.", True)> _
    Public Shared Sub ZoomToExtent(ByRef pEnv As ESRI.ArcGIS.Geometry.IEnvelope, ByRef pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument)
        Dim pMap As ESRI.ArcGIS.Carto.IMap
        Dim pActiveView As ESRI.ArcGIS.Carto.IActiveView

        ' Gets a reference to the current view window
        pMap = pMxDoc.FocusMap
        pActiveView = DirectCast(pMap, IActiveView)

        ' Updates the view's extent
        pActiveView.Extent = pEnv
        pActiveView.Refresh()
    End Sub

    ''' <summary>
    ''' Zooms to the given envenlope.
    ''' </summary>
    ''' <param name="theEnvelope">The envelope to zoom to.</param>
    ''' <remarks>Replaces ZoomToExtent sub.  This sub removes the unneeded pMxDoc parameter.</remarks>
    ''' 
    Public Shared Sub ZoomToEnvelope(ByRef theEnvelope As IEnvelope)

        Dim theArcMapDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
        Dim theMap As IMap = theArcMapDoc.FocusMap
        Dim theActiveView As IActiveView = DirectCast(theMap, IActiveView)

        '-- Updates the view's extent
        theActiveView.Extent = theEnvelope
        theActiveView.Refresh()

    End Sub

    ''' <summary>
    ''' Selects a single feature.
    ''' </summary>
    ''' <param name="theFeatureLayer">The feature layer containing the feature.</param>
    ''' <param name="theFeature">The feature to zoom to.</param> 
    ''' <remarks></remarks>
    ''' 
    Public Shared Sub SetSelectedFeature(ByRef theFeatureLayer As IFeatureLayer, ByRef theFeature As IFeature)

        Dim theArcMapDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
        Dim theMap As IMap = theArcMapDoc.FocusMap
        Dim theActiveView As IActiveView = DirectCast(theMap, IActiveView)

        '-- Select the feature
        theMap.ClearSelection()
        theMap.SelectFeature(theFeatureLayer, theFeature)
        theActiveView.Refresh()

    End Sub





#End Region

#Region "Private Members"

    ''' <summary>
    ''' Return a cursor that represents the results of an attribute query.
    ''' </summary>
    ''' <param name="table">An object that supports the ITable interface.</param>
    ''' <param name="whereClause">A Sql Where clause without the WHERE.</param>
    ''' <returns>Return a cursor that represents the results of an attribute query.</returns>
    ''' <remarks>Creates a cursor from table that contains all feature records that meet the criteria in <paramref name="whereClause">whereClause</paramref>.</remarks>
    Private Shared Function attributeQuery(ByRef table As ITable, Optional ByRef whereClause As String = "") As ICursor
        Try
            Dim thisQueryFilter As IQueryFilter
            thisQueryFilter = New QueryFilter
            thisQueryFilter.WhereClause = whereClause
            Dim thisCursor As ICursor
            thisCursor = table.Search(thisQueryFilter, False)
            If table.RowCount(thisQueryFilter) = 0 Then
                Return Nothing
            Else
                Return thisCursor
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Calculate ORMAP Taxlot Number when one if its components has changed.
    ''' </summary>
    ''' <param name="existingORMapNum">An ORMAP Number.</param>
    ''' <param name="theFeature">An object that supports the IFeature interface.</param>
    ''' <param name="taxlotValue">A taxlot number.</param>
    ''' <returns>A string that represents an ORMAP number updated with the value from theFeature and taxlotValue.</returns>
    ''' <remarks>Given an ORMAP Number, <paramref name="existingORMapNum">existingORMapNum</paramref>, and feature, <paramref name="theFeature">theFeature</paramref>,
    ''' and a taxlot value,<paramref name="taxlotValue">taxlotValue</paramref>.
    ''' Remove the existing map suffix type and number from <paramref name="existingORMapNum">existingORMapNum</paramref> and replace them with the new values in <paramref name="theFeature">theFeature</paramref> and
    ''' append <paramref name="taxlotValue">taxlotValue</paramref> to form the return value.</remarks>
    Private Shared Function generateORMAPTaxlotNumber(ByVal existingORMapNum As String, ByRef theFeature As IFeature, ByVal taxlotValue As String) As String
        Try
            Dim shortORMapNum As String = existingORMapNum.Substring(0, 19) 'replaces the ShortenOMTLNum function 
            Dim taxlotMapSufNumberValue As String = GetMapSuffixNum(theFeature)
            Dim taxlotMapSufTypeValue As String = GetMapSuffixType(theFeature)
            ' Recreate and return the ORMAP Taxlot number
            Return String.Concat(shortORMapNum, taxlotMapSufTypeValue, taxlotMapSufNumberValue, taxlotValue)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Converts the list of candidate dictionary keys to a string in the 
    ''' format of a comma-delimited list.
    ''' </summary>
    ''' <param name="candidatesDictionary">A dictionary of integer candidate IDs for keys.</param>
    ''' <returns>A string in the format of a comma-delimited list.</returns>
    ''' <remarks>An empty dictionary will return as an empty string.</remarks>
    Private Shared Function candidateKeysToDelimitedString(ByRef candidatesDictionary As Dictionary(Of Integer, Double)) As String
        Dim keyCollection As ICollection(Of Integer) = DirectCast(candidatesDictionary.Keys, ICollection(Of Integer))
        Dim returnValue As String = String.Empty

        If keyCollection.Count > 0 Then
            ' The elements of the keyCollection are strongly typed
            ' with the type that was specified for dictionary values.
            For Each n As Integer In keyCollection
                ' Append keys as strings and commas.
                returnValue &= (CStr(n) & ",")
            Next n
            ' Trim off the last comma.
            returnValue = Left(returnValue, returnValue.Length - 1)
        Else
            ' Will return the default empty string.
        End If
        Return returnValue

    End Function

    ''' <summary>
    ''' Return a feature cursor based on the results of a spatial query.
    ''' </summary>
    ''' <param name="inFeatureClass">Feature class to search.</param>
    ''' <param name="searchGeometry">Geometry to search in relation to spatialRelation.</param>
    ''' <param name="spatialRelation">Geometry relationship to searchGeometry.</param>
    ''' <param name="whereClause">SQL Where clause.</param>
    ''' <param name="isUpdateable">Read/Write state of the return cursor.</param>
    ''' <returns>Returns a feature cursor that represents the results of the spatial query.</returns>
    ''' <remarks>Given a feature class, <paramref name="inFeatureClass">inFeatureClass</paramref>, 
    ''' a search geometry,<paramref name=" searchGeometry">searchGeometry</paramref> , a spatial relationship,<paramref name=" spatialRelation">spatialRelation</paramref> , 
    ''' an Sql search statement, <paramref name=" whereClause">whereClause</paramref>,
    ''' and whether or not the returned cursor should be updateable, <paramref name=" isUpdateable">IsUpdateable</paramref>.
    ''' Perform a spatial query <paramref name="infeatureClass">inFeatureClass</paramref> where feature
    ''' which meet criteria whereClause have a relationship of spatialRelation to searchGeometry. The returned cursor is updatable if IsUpdateable is True.</remarks>
    Private Shared Function doSpatialQuery(ByRef inFeatureClass As IFeatureClass, ByRef searchGeometry As IGeometry, ByRef spatialRelation As ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum, Optional ByRef whereClause As String = "", Optional ByVal isUpdateable As Boolean = False) As IFeatureCursor
        Try
            Dim thisSpatialFilter As ISpatialFilter
            thisSpatialFilter = New SpatialFilter
            thisSpatialFilter.Geometry = searchGeometry

            Dim shapeFieldName As String
            shapeFieldName = inFeatureClass.ShapeFieldName
            thisSpatialFilter.GeometryField = shapeFieldName

            thisSpatialFilter.SpatialRel = spatialRelation
            thisSpatialFilter.WhereClause = whereClause

            Dim thisQueryFilter As IQueryFilter
            thisQueryFilter = thisSpatialFilter
            Dim thisFeatureCursor As IFeatureCursor
            If isUpdateable Then
                thisFeatureCursor = inFeatureClass.Update(thisQueryFilter, False)
            Else
                thisFeatureCursor = inFeatureClass.Search(thisQueryFilter, False)
            End If
            Return thisFeatureCursor

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Copy envelope points to polygon.
    ''' </summary>
    ''' <param name="envelope"></param>
    ''' <returns>A Polygon.</returns>
    ''' <remarks></remarks>
    Private Shared Function envelopeToPolygon(ByRef envelope As IEnvelope) As IPolygon
        Try
            Dim thisPolygon As IPolygon
            thisPolygon = New Polygon
            Dim pointCollection As IPointCollection
            pointCollection = DirectCast(thisPolygon, IPointCollection)
            With pointCollection
                .AddPoint(envelope.UpperRight)
                .AddPoint(envelope.LowerRight)
                .AddPoint(envelope.LowerLeft)
                .AddPoint(envelope.UpperLeft)
            End With
            Return thisPolygon
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Determine the annotation size based on scale.
    ''' </summary>
    ''' <param name="thisFeatureClassName">Feature class to find the proper annotation size for.</param>
    ''' <param name="scale">The scale of the feature class.</param>
    ''' <returns>A double that represents the proper scale factor.</returns>
    ''' <remarks>Determines the proper size for the text in thisFeatureClassName.
    ''' Defaults at size 5 is the <paramref name="scale">scale</paramref> is invalid, and size 10 if
    ''' <paramref name="thisFeatureClassName">thisFeatureClassName</paramref> is not Taxlot Acreage Annotation or Taxlot Number Annotation.</remarks>
    Private Shared Function getAnnoSizeByScale(ByVal thisFeatureClassName As String, ByVal scale As Integer) As Double
        Try 'TODO: JWM verify the table names that we are comparing
            Dim size As String
            If String.Compare(thisFeatureClassName, EditorExtension.AnnoTableNamesSettings.TaxlotAcreageAnnoFC, True) = 0 Then
                Select Case scale
                    Case 120
                        size = EditorExtension.TaxlotAcreageAnnoSettings.TextSize00120Scale
                    Case 240
                        size = EditorExtension.TaxlotAcreageAnnoSettings.TextSize00240Scale
                    Case 360
                        size = EditorExtension.TaxlotAcreageAnnoSettings.TextSize00360Scale
                    Case 480
                        size = EditorExtension.TaxlotAcreageAnnoSettings.TextSize00480Scale
                    Case 600
                        size = EditorExtension.TaxlotAcreageAnnoSettings.TextSize00600Scale
                    Case 1200
                        size = EditorExtension.TaxlotAcreageAnnoSettings.TextSize01200Scale
                    Case 2400
                        size = EditorExtension.TaxlotAcreageAnnoSettings.TextSize02400Scale
                    Case 4800
                        size = EditorExtension.TaxlotAcreageAnnoSettings.TextSize04800Scale
                    Case 9600
                        size = EditorExtension.TaxlotAcreageAnnoSettings.TextSize09600Scale
                    Case 24000
                        size = EditorExtension.TaxlotAcreageAnnoSettings.TextSize24000Scale
                    Case Else
                        ' Default size
                        size = "5"
                End Select
            ElseIf String.Compare(thisFeatureClassName, EditorExtension.AnnoTableNamesSettings.TaxlotNumberAnnoFC, True) = 0 Then
                Select Case scale
                    Case 120
                        size = EditorExtension.TaxlotNumberAnnoSettings.TextSize00120Scale
                    Case 240
                        size = EditorExtension.TaxlotNumberAnnoSettings.TextSize00240Scale
                    Case 360
                        size = EditorExtension.TaxlotNumberAnnoSettings.TextSize00360Scale
                    Case 480
                        size = EditorExtension.TaxlotNumberAnnoSettings.TextSize00480Scale
                    Case 600
                        size = EditorExtension.TaxlotNumberAnnoSettings.TextSize00600Scale
                    Case 1200
                        size = EditorExtension.TaxlotNumberAnnoSettings.TextSize01200Scale
                    Case 2400
                        size = EditorExtension.TaxlotNumberAnnoSettings.TextSize02400Scale
                    Case 4800
                        size = EditorExtension.TaxlotNumberAnnoSettings.TextSize04800Scale
                    Case 9600
                        size = EditorExtension.TaxlotNumberAnnoSettings.TextSize09600Scale
                    Case 24000
                        size = EditorExtension.TaxlotNumberAnnoSettings.TextSize24000Scale
                    Case Else
                        size = "5"
                End Select
            Else
                size = "10" 'default
            End If
            Return CDbl(size)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return 10 'default

        End Try
    End Function

    ''' <summary>
    '''Private empty constructor to prevent instantiation.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()
    End Sub

#End Region

#End Region

End Class
#End Region
