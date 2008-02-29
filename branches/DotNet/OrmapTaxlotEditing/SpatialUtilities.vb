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

#Region "Subversion Keyword expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision$
'Date of Last Change: $Date$
#End Region

#Region "Imported namespace statements"
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.Carto
#End Region

#Region "Class Declaration"
''' <summary>
'''  General utility class
''' </summary>
''' <remarks>Commonly used procedures and functions</remarks>
Public NotInheritable Class SpatialUtilities
#Region "Custom Class Members"
#Region "Public Members"

    ''' <summary>
    ''' Calculates Taxlot values from ORMAPMapnum
    ''' </summary>
    ''' <param name="feature">A feature from the Taxlot feature class</param>
    ''' <param name="mapIndexLayer">The Map Index feature layer</param>
    ''' <remarks>Update the ORMAP fields in a_pFeat to the reflect the current ORMAP Number and Map Number elements in the overlaying Map Index polygon.</remarks>
    Public Shared Sub CalculateTaxlotValues(ByRef feature As IFeature, ByRef mapIndexLayer As IFeatureLayer)
        'TODO: JWM flesh out
        Try

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Converts a domain descriptive value to the stored code
    ''' </summary>
    ''' <param name="fields">An object that supports the IFields interface</param>
    ''' <param name="fieldName">A field that exists in fields</param>
    ''' <param name="codedValue">A coded name to convert to a coded value</param>
    ''' <returns>A string that represents the domain coded value that corresponds with the coded name, codedValue, or a empty string.</returns>
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
            MessageBox.Show(ex.Message)
            ConvertCodeValueDomainToCode = String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Converts a code index to the  domain descriptive value
    ''' </summary>
    ''' <param name="fields">An object that supports the IFields interface</param>
    ''' <param name="fieldName">A field that exists in fields</param>
    ''' <param name="codedValue">A coded value to covert to a coded name</param>
    ''' <returns>A string that represents the domain coded name that corresponds with the coded value, codedValue, or a zero-length string.</returns>
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
            MessageBox.Show(ex.Message)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Locate a feature layer in the TOC
    ''' </summary>
    ''' <param name="dataSetName">The name of the dataset to find</param>
    ''' <returns>A layer object of that supports the IFeatureLayer interface.</returns>
    ''' <remarks>Return the Feature Layer based on its dataset name, recursively</remarks>
    Public Shared Function FindFeatureLayerByDSName(ByVal dataSetName As String) As ESRI.ArcGIS.Carto.IFeatureLayer
        Try
            Dim returnValue As IFeatureLayer

            Dim thisDoc As Object 'TODO: JWM need a pointer/handle to the ARcMap document
            Dim theMap As IMap = thisDoc.focusMap
            Dim thisUID As UID
            ' Get a reference to the feature layers collection of the document. 
            thisUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
            'We want a EnumLayer containing all FeatureLayer objects.
            Dim theLayers As IEnumLayer
            theLayers = theMap.Layers(thisUID, True)

            theLayers.Reset()
            Dim thisLayer As ILayer
            thisLayer = theLayers.Next()

            Dim thisFeatureLayer As IFeatureLayer
            Dim thisDataSet As IDataset

            Do While Not (thisLayer Is Nothing)
                thisFeatureLayer = DirectCast(thisLayer, IFeatureLayer)
                thisDataSet = DirectCast(thisFeatureLayer.FeatureClass, IDataset)
                If Not (thisDataSet Is Nothing) Then
                    If String.Compare(thisDataSet.Name, dataSetName, True) = 0 Then
                        returnValue = DirectCast(thisLayer, IFeatureLayer)
                    End If
                End If
                thisLayer = theLayers.Next
            Loop

            If (returnValue Is Nothing) Then 'TODO: JWM Is this how we should return objects?
                returnValue = New FeatureLayer
                Return returnValue
            Else
                Return returnValue
            End If

            thisUID = Nothing
            theMap = Nothing
            thisDoc = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Retrieve a feature-linked annotation feature
    ''' </summary>
    ''' <param name="anObject">An initialized geodatabase object</param>
    ''' <returns>An object that supports the IFeature interface.</returns>
    ''' <remarks>Finds all related objects to the feature through the first found relationship class, and returns the first related object as the return value.
    ''' This is optimized for annotation because there is a single relationship class.</remarks>
    Public Shared Function GetRelatedObjects(ByVal anObject As IObject) As IFeature
        Try
            Dim relationshipClassEnum As IEnumRelationshipClass
            Dim parentSet As ISet

            relationshipClassEnum = anObject.Class.RelationshipClasses(esriRelRole.esriRelRoleAny)
            If Not relationshipClassEnum Is Nothing Then
                Dim thisRelationshipClass As IRelationshipClass
                thisRelationshipClass = relationshipClassEnum.Next
                If Not thisRelationshipClass Is Nothing Then
                    parentSet = thisRelationshipClass.GetObjectsRelatedToObject(anObject)
                End If
            Else
                Return Nothing
            End If
            ' Returns the first related feature
            If Not parentSet Is Nothing Then
                Dim parentFeature As IFeature
                parentFeature = DirectCast(parentSet.Next, IFeature)
                If Not parentFeature Is Nothing Then
                    Return parentFeature
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Return a cursor for the selected features
    ''' </summary>
    ''' <param name="layer">The feature layer to return the selection from</param>
    ''' <returns>An object that supports the IFeatureCursor interface</returns>
    ''' <remarks>References the currently selected features in layer, and return a cursor with the feature in it.</remarks>
    Public Shared Function GetSelectedFeatures(ByVal layer As IFeatureLayer) As IFeatureCursor
        Try
            If Not TypeOf layer Is IFeatureLayer Then
                Return Nothing
            End If

            Dim thisSelection As IFeatureSelection
            thisSelection = DirectCast(layer, IFeatureSelection)
            Dim returnValue As IFeatureCursor
            ' Returns a cursor with the selected items from the specified feature layer
            thisSelection.SelectionSet.Search(Nothing, False, DirectCast(returnValue, ICursor))
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Overlay the passed in feature with a feature class to get a value from fieldName
    ''' </summary>
    ''' <param name="theGeometry">The search geometry</param>
    ''' <param name="ovrlayFeatureClass">Overlying feature class</param>
    ''' <param name="fieldName">Name of field to return value for</param>
    ''' <param name="orderBestByFldName">Name of field to order by in the case of a tie in area/length of intersection</param>
    ''' <returns>Returns the value from the specified field (fieldName) as a string</returns>
    ''' <remarks>Get the target feature with the largest area of intersection with pGeometry and get its value from the field, or, if tied
    '''(unikely but possible), then get the best (lowest) value from the field.</remarks>
    Public Shared Function GetValueViaOverlay(ByRef theGeometry As IGeometry, ByRef ovrlayFeatureClass As IFeatureClass, ByVal fieldName As String, Optional ByVal orderBestByFldName As String = "") As String
        Try
            Dim continueThisProcess As Boolean

            continueThisProcess = True 'initialize

            If (theGeometry Is Nothing) OrElse (ovrlayFeatureClass Is Nothing) OrElse (Len(fieldName) <= 0) Then
                continueThisProcess = False
            End If

            Dim fieldIndex As Integer
            fieldIndex = ovrlayFeatureClass.Fields.FindField(fieldName)
            If fieldIndex < 0 Then
                continueThisProcess = False
            End If

            Dim indexOrderByField As Integer
            If orderBestByFldName.Length = 0 Then
                ' Use the value field as the order-by field
                indexOrderByField = fieldIndex
            Else
                indexOrderByField = ovrlayFeatureClass.Fields.FindField(orderBestByFldName)
                If indexOrderByField < 0 Then
                    'field not found. Try the OID field
                    If ovrlayFeatureClass.HasOID Then
                        indexOrderByField = ovrlayFeatureClass.Fields.FindField(ovrlayFeatureClass.OIDFieldName)
                        MessageBox.Show("Field " & orderBestByFldName & " not found in " & ovrlayFeatureClass.AliasName & ". Using " & ovrlayFeatureClass.OIDFieldName, "Taxlot Editing - Get Value Via Overlay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Else
                        continueThisProcess = False
                    End If
                End If
            End If

            Dim thisPolygon As IPolygon
            Dim thisArea As IArea
            Dim thisCurve As ICurve
            Dim thisEnvelope As IEnvelope
            Dim aPointCollection As IPointCollection
            Dim workingFuzzAmount As Double
            Dim dictCandidates As Dictionary(Of Integer, Double)
            Const fuzzFactor As Double = 0.05
            Dim theBestValue As String = ""

            If continueThisProcess Then
                Select Case theGeometry.GeometryType
                    Case esriGeometryType.esriGeometryPolygon
                        thisPolygon = DirectCast(theGeometry, IPolygon)
                        thisArea = DirectCast(thisPolygon, IArea)
                        workingFuzzAmount = thisArea.Area * fuzzFactor

                    Case esriGeometryType.esriGeometryPolyline, esriGeometryType.esriGeometryLine, esriGeometryType.esriGeometryBezier3Curve, esriGeometryType.esriGeometryCircularArc, esriGeometryType.esriGeometryEllipticArc, esriGeometryType.esriGeometryPath
                        thisCurve = DirectCast(theGeometry, ICurve)
                        workingFuzzAmount = thisCurve.Length * fuzzFactor

                    Case esriGeometryType.esriGeometryEnvelope
                        thisEnvelope = DirectCast(theGeometry, IEnvelope)
                        thisPolygon = EnvelopeToPolygon(thisEnvelope)
                    Case Else
                        Exit Try 'HACK JWM
                End Select

                Dim largestInternalArea As Double = 0
                Dim longestInternalLength As Double = 0

                Dim aFeatureCursor As IFeatureCursor
                Dim aFeature As IFeature
                aFeatureCursor = DoSpatialQuery(ovrlayFeatureClass, theGeometry, esriSpatialRelEnum.esriSpatialRelIntersects)
                If Not (aFeatureCursor Is Nothing) Then
                    aFeature = aFeatureCursor.NextFeature
                    Dim topoOperator As ITopologicalOperator
                    Dim interiorGeometry As IGeometry
                    Dim interiorArea As Double = 0
                    dictCandidates = New Dictionary(Of Integer, Double)
                    While Not (aFeature Is Nothing)

                        topoOperator = DirectCast(aFeature.Shape, ITopologicalOperator)
                        If Not topoOperator.IsSimple Then
                            topoOperator.Simplify()
                        End If


                        Select Case theGeometry.GeometryType

                            Case esriGeometryType.esriGeometryEnvelope, esriGeometryType.esriGeometryPolygon
                                ' Determine if the target feature has the largest area of intersection with the
                                ' current source feature. Set flags used below.
                                interiorGeometry = topoOperator.Intersect(theGeometry, esriGeometryDimension.esriGeometry2Dimension)
                                If Not interiorGeometry.IsEmpty Then
                                    thisPolygon = DirectCast(interiorGeometry, IPolygon)
                                    thisArea = DirectCast(thisPolygon, IArea)
                                    interiorArea = thisArea.Area
                                    If System.Math.Abs(interiorArea - largestInternalArea) > workingFuzzAmount Then
                                        '[Difference greater than fuzz amount, so not a "fuzzy tie"...]

                                        ' Determine if this is the feature with the largest area of intersection
                                        If interiorArea > largestInternalArea Then
                                            largestInternalArea = interiorArea
                                            '[New longest intersection...]
                                            ' Get the value only for this feature
                                            'TODO: JWM Need to empty collection then add
                                            dictCandidates.Add(aFeature.OID, largestInternalArea)
                                        End If
                                    Else
                                        '[Difference not greater than fuzz amount, so a "fuzzy tie"...]

                                        ' Evaluate this feature against other tied candidates
                                        dictCandidates.Add(aFeature.OID, largestInternalArea)
                                    End If
                                End If

                            Case esriGeometryType.esriGeometryLine, esriGeometryType.esriGeometryPolyline
                                ' Determine if the target feature has the longest length of intersection with the
                                ' current source feature. Set flags used below.
                                interiorGeometry = topoOperator.Intersect(theGeometry, esriGeometryDimension.esriGeometry1Dimension)
                                If Not interiorGeometry.IsEmpty Then
                                    thisCurve = DirectCast(interiorGeometry, ICurve)
                                    Dim interiorLength As Double
                                    interiorLength = thisCurve.Length
                                    If System.Math.Abs(interiorLength - longestInternalLength) > workingFuzzAmount Then
                                        '[Difference greater than fuzz amount, so not a "fuzzy tie"...]
                                        ' Determine if this is the feature with the longest length of intersection
                                        If interiorLength > longestInternalLength Then
                                            longestInternalLength = interiorLength
                                            'TODO: JWM Need to empty collection then add
                                            dictCandidates.Add(aFeature.OID, longestInternalLength)
                                        End If
                                    Else
                                        '[Difference not greater than fuzz amount, so a "fuzzy tie"...]
                                        'Evaluate this feature against other tied candidates
                                        dictCandidates.Add(aFeature.OID, longestInternalLength)
                                    End If
                                End If
                            Case esriGeometryType.esriGeometryPoint
                                '[Tied by definition (0-dimension = zero area & length = tie)...]
                                ' Evaluate this feature against other tied candidates
                                dictCandidates.Add(aFeature.OID, 0)
                            Case Else
                                Exit Try 'HACK jwm
                        End Select
                        aFeature = aFeatureCursor.NextFeature
                    End While
                End If
                If dictCandidates.Count > 0 Then
                    Dim theValue As String = ""

                    Dim theBestOrderByValue As String = ""

                    Dim whereClause As String = String.Concat(ovrlayFeatureClass.OIDFieldName, " in (", Join(dictCandidates.KeyCollection, ","), ")") 'TODO JWM Fix this assignment
                    aFeatureCursor = DoSpatialQuery(ovrlayFeatureClass, theGeometry, esriSpatialRelEnum.esriSpatialRelIntersects, whereClause)
                    If Not (aFeatureCursor Is Nothing) Then
                        aFeature = aFeatureCursor.NextFeature
                        While Not (aFeature Is Nothing)
                            If dictCandidates.Count > 1 Then
                                '[Candidates tied for area/length of intersection (unikely but possible)...]
                                ' Get the value only if the candidate source feature has the best (lowest) order-by value
                                Dim orderByValue As String = CStr(aFeature.Value(indexOrderByField))
                                If theBestOrderByValue.Length = 0 Or (orderByValue < theBestOrderByValue) Then 'TODO: JWM Ask Nick about this comparison
                                    theBestOrderByValue = orderByValue
                                    If Not IsDBNull(aFeature.Value(fieldIndex)) Then
                                        theValue = CStr(aFeature.Value(fieldIndex))
                                        theBestValue = theValue
                                    Else
                                        MessageBox.Show("The field " & aFeature.Fields.Field(fieldIndex).Name & " contains a Null" & vbNewLine & "for the feature with OID=" & aFeature.OID & ".", "Error: GetValueViaOverlay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                    End If
                                End If
                            Else
                                '[Candidates not tied (only one)...]
                                ' Get the value for the only candidate feature
                                If Not IsDBNull(aFeature.Value(fieldIndex)) Then
                                    theValue = CStr(aFeature.Value(fieldIndex))
                                    theBestValue = theValue
                                Else
                                    MessageBox.Show("The field " & aFeature.Fields.Field(fieldIndex).Name & " contains a Null" & vbNewLine & "for the feature with OID=" & aFeature.OID & ".", "Error: GetValueViaOverlay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                End If
                            End If 'dictCandidates.count >1
                        End While
                    End If 'featurecursor is not nothing
                Else
                    MessageBox.Show("No " & ovrlayFeatureClass.AliasName & "features found " & vbNewLine & " which intersect this feature.", "Error: GetValueViaOverlay", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If 'dictionary count > zero
            End If 'continueThisProcess 

            Return theBestValue

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Determines if the feature layer has a selection
    ''' </summary>
    ''' <param name="layer">An object that supports the IFeatureLayer2</param>
    ''' <returns>True or False</returns>
    ''' <remarks>Checking the selection set of layer, determine if one, many, or no features are selected.</remarks>
    Public Shared Function HasSelectedFeatures(ByVal layer As IFeatureLayer2) As Boolean
        Try
            Dim returnValue As Boolean = False
            If (layer Is Nothing) Or Not (TypeOf layer Is IFeatureLayer) Then
                Return False
            End If
            'how many are selected
            Dim featuresSelected As IFeatureSelection
            featuresSelected = DirectCast(layer, IFeatureSelection)
            Dim thisFeatureCursor As IFeatureCursor
            featuresSelected.SelectionSet.Search(Nothing, False, DirectCast(thisFeatureCursor, ICursor))
            If Not thisFeatureCursor Is Nothing Then
                Dim thisFeature As IFeature
                thisFeature = thisFeatureCursor.NextFeature
                If thisFeature Is Nothing Then
                    returnValue = True
                End If
            End If
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determine if the feature belongs to the Taxlot feature class
    ''' </summary>
    ''' <param name="thisObject">A valid initialized geodatabase object</param>
    ''' <returns>True or False</returns>
    ''' <remarks>Determine if thisObject belongs to the Taxlot feature class by checking the name of the dataset of thisObject feature class
    ''' againts the Taxlot Feature Class constant.</remarks>
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
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determine if the feature belongs to the MapIndex feature class
    ''' </summary>
    ''' <param name="thisObject">A valid initialized geodatabase object</param>
    ''' <returns>True or False</returns>
    ''' <remarks>Compares the name of the dataset of thisObject feature class to the Map Index layer name in order to determine if thisObject
    ''' belongs to the MapIndex feature class.</remarks>
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
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determine if a feature is annotation
    ''' </summary>
    ''' <param name="thisObject">A valid initialized geodatabase object</param>
    ''' <returns>True or False</returns>
    ''' <remarks>Compares the feature type of thisObjec with that of annotation
    ''' and return the truth value of the comparison</remarks>
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
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Loads a feature class into the current map
    ''' </summary>
    ''' <param name="featureClassName">The feature class to find</param>
    ''' <param name="title">An alternate title for the file dialog box</param>
    ''' <returns>True for loaded, False for not loaded</returns>
    ''' <remarks>Show a dialog box with title; title that allows the user to select the personal geodatabase that featureClassName resides in.
    ''' The feature class featureClassName is then loaded from the chosen personal geodatabase.</remarks>
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
                .ShowOPen()
            End With

            'exit if there is nothing selected
            If Len(thisFileDialog.SelectedObject(1)) = 0 Then
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

            Dim thisArcMapDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
            thisArcMapDoc= object 'TODO: Need a pointer/handle to the application
            Dim thisMap As ESRI.ArcGIS.Carto.IMap
            thisMap = thisArcMapDoc.FocusMap

            thisMap.AddLayer(thisFeatureLayer)
            thisArcMapDoc.CurrentContentsView.Refresh(0)
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Reads a value from a row, given a field name.
    ''' </summary>
    ''' <param name="row">An object that implements the IRow interface</param>
    ''' <param name="fieldName">A field that exists in row</param>
    ''' <param name="dataType">A string value indicating data type of the field</param>
    ''' <returns></returns>
    ''' <remarks>Reads the value of a field with a domain and translates the value from the coded value to the coded name.</remarks>
    Public Shared Function ReadValue(ByRef row As IRow, ByVal fieldName As String, Optional ByVal dataType As String = "") As String
        Try
            Dim fieldIndex As Integer
            Dim returnValue As String = ""

            fieldIndex = row.Fields.FindField(fieldName)
            If fieldIndex > -1 Then
                If String.Compare(dataType, "date", True) = 0 Then
                    If IsDBNull(row.Value(fieldIndex)) Then
                        returnValue = CStr(System.DateTime.Today)
                    Else
                        returnValue = CStr(row.Value(fieldIndex))
                    End If
                Else
                    If IsDBNull(row.Value(fieldIndex)) Then
                        returnValue = String.Empty
                    Else
                        returnValue = CStr(row.Value(fieldIndex))
                    End If
                End If
                'Determine if a Domain Field
                Dim field As IField
                field = row.Fields.Field(fieldIndex)
                Dim domain As IDomain
                domain = field.Domain
                If Not domain Is Nothing Then
                    If domain.Type = esriDomainType.esriDTCodedValue Then
                        'If TypeOf domain Is ICodedValueDomain Then
                        Dim thisCodedValueDomain As ICodedValueDomain
                        thisCodedValueDomain = DirectCast(domain, ICodedValueDomain)
                        Dim domainValue As Object
                        domainValue = row.Value(fieldIndex)
                        'search domain for the code
                        For domainIndex As Integer = 0 To thisCodedValueDomain.CodeCount - 1
                            If thisCodedValueDomain.Value(domainIndex) = domainValue Then 'TODO: JWM Need to resolve this comparison
                                returnValue = thisCodedValueDomain.Name(domainIndex)
                                Exit For
                            End If
                        Next domainIndex
                    End If
                End If
            End If

            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Update Auto fields in a feature class.
    ''' </summary>
    ''' <param name="feature"> An object that implements the Ifeature interface</param>
    ''' <remarks>Update the AutoWho and the AutoDate fields with the current username and date/time, respectively.</remarks>
    Public Shared Sub UpdateAutoFields(ByRef feature As IFeature)
        Try
            If feature Is Nothing Then
                Exit Try
            End If
            Dim indexAutoDateField As Integer

            indexAutoDateField = feature.Fields.FindField(EditorExtension.AllTablesSettings.AutoDateField)

            If indexAutoDateField > -1 Then
                feature.Value(indexAutoDateField) = System.DateTime.Now
            End If
            Dim indexAutoWhoField As Integer
            indexAutoWhoField = feature.Fields.FindField(EditorExtension.AllTablesSettings.AutoWhoField)
            If indexAutoWhoField > -1 Then
                feature.Value(indexAutoWhoField) = Utilities.GetUserName
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Determine the validity of a taxlot number
    ''' </summary>
    ''' <param name="taxlotNumber">The taxlot value to validate</param>
    ''' <param name="thisGeometry">The geometry of the feature to check</param>
    ''' <returns>True or False</returns>
    ''' <remarks>Determine if the feature represented by thisGeometry with taxlot taxlotNumber is a unique and therefore valid.</remarks>
    Public Shared Function ValidateTaxlotNumber(ByVal taxlotNumber As String, ByRef thisGeometry As IGeometry) As Boolean
        Try
            Dim returnValue As Boolean = False
            'check for existence of Taxlot layer
            Dim thisTaxlotFeatureLayer As IFeatureLayer
            thisTaxlotFeatureLayer = FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.MapIndexFC)
            If thisTaxlotFeatureLayer Is Nothing Then 'TODO: JWM Place strings in resource file and use may use different type of notification
                MessageBox.Show("Unable to locate Taxlot layer in Table of Contents. This process requires a feature class called " & EditorExtension.TableNamesSettings.TaxLotFC)
                Return returnValue
            End If

            'check for existence of Map index layer
            Dim thisTaxlotFeatureClass As IFeatureClass
            thisTaxlotFeatureClass = thisTaxlotFeatureLayer.FeatureClass

            Dim mapIndexFeatureLayer As IFeatureLayer
            mapIndexFeatureLayer = FindFeatureLayerByDSName(EditorExtension.TableNamesSettings.MapIndexFC)
            If mapIndexFeatureLayer Is Nothing Then
                MessageBox.Show("Unable to locate MapIndex layer in Table of Contents. This process requires a feature class called " & EditorExtension.TableNamesSettings.MapIndexFC)
                Return returnValue
            End If

            Dim mapIndexFeatureClass As IFeatureClass
            mapIndexFeatureClass = mapIndexFeatureLayer.FeatureClass

            ' Checks for the existence of a current ORMAP Number and Taxlot number
            Dim mapIndexORMAPValue As String
            'mapIndexORMAPValue= getvalueViaOverlay()'TODO: JWM Flesh this function out
            If mapIndexORMAPValue.Length = 0 Then
                returnValue = True
            End If

            'Make sure this number is unique within taxlots with this OM number
            'TODO: JWM check these EditorExtension values
            Dim whereClause As String = String.Concat(EditorExtension.TaxLotSettings.MapNumberField, "='", mapIndexORMAPValue, "' AND ", EditorExtension.TaxLotSettings.TaxlotField, " = '", taxlotNumber, "'")
            Dim cursor As ICursor
            cursor = AttributeQuery(DirectCast(thisTaxlotFeatureClass, ITable), whereClause)
            If Not (cursor Is Nothing) And (returnValue = False) Then
                Dim row As IRow
                row = cursor.NextRow
                If row Is Nothing Then
                    returnValue = True
                End If
            End If

            Return returnValue

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function


#End Region

#Region "Private Members"

    ''' <summary>
    ''' Return a cursor that represents the results of an attribute query
    ''' </summary>
    ''' <param name="table">An object that supports the ITable interface</param>
    ''' <param name="whereClause">An Sql Where clause</param>
    ''' <returns>Return a cursor that represents the results of an attribute query</returns>
    ''' <remarks>Creates a cursor from table that contains all feature records that meet the criteria in whereClause</remarks>
    Private Shared Function AttributeQuery(ByRef table As ITable, Optional ByRef whereClause As String = "") As ICursor
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
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Return a feature cursor based on the results of a spatial query
    ''' </summary>
    ''' <param name="inFeatureClass">Feature class to search</param>
    ''' <param name="searchGeometry">Geometry to search in relation to spatialRelation</param>
    ''' <param name="spatialRelation">Geometry relationship to searchGeometry</param>
    ''' <param name="whereClause">SQL Where clause</param>
    ''' <param name="isUpdateable">Read/Write state of the return cursor</param>
    ''' <returns>Returns a feature cursor that represents the results of the spatial query</returns>
    ''' <remarks>Given a feature class, pFeatureClassIn, a search geometry, pSearchGeometry, a spatial relationship, lSpatialRelation,
    ''' an Sql search statement, sWhereClause, and whether or not
    ''' the returned cursor should be updateable, bUpdateable.
    ''' Perform a spatial query on pFeatureClassIn where feature
    ''' which meet criteria sWhereClause have a relationship of lSpatialRelation to pSearchGeometry. The returned cursor is updatable if bUpdateable is True.</remarks>
    Private Shared Function DoSpatialQuery(ByRef inFeatureClass As IFeatureClass, ByRef searchGeometry As IGeometry, ByRef spatialRelation As ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum, Optional ByRef whereClause As String = "", Optional ByVal isUpdateable As Boolean = False) As IFeatureCursor
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
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Copy envelope points to polygon
    ''' </summary>
    ''' <param name="envelope"></param>
    ''' <returns>A Polygon</returns>
    ''' <remarks></remarks>
    Private Shared Function EnvelopeToPolygon(ByRef envelope As IEnvelope) As IPolygon
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
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Determine the annotation size based on scale
    ''' </summary>
    ''' <param name="thisFeatureClassName">Feature class to find the proper annotation size for</param>
    ''' <param name="scale">The scale of the feature class</param>
    ''' <returns>A double that represents the proper scale factor</returns>
    ''' <remarks>Determines the proper size for the text in thisFeatureClassName. Defaults at size 5 is the scale is invalid, and size 10 if
    ''' thisFeatureClassName is not Taxlot Acreage Annotation or Taxlot Number Annotation</remarks>
    Private Shared Function GetAnnoSizeByScale(ByVal thisFeatureClassName As String, ByVal scale As Integer) As Double
        Try 'TODO: JWM verify the table names that we are comparing
            Dim size As String
            If String.Compare(thisFeatureClassName, EditorExtension.AnnoTableNamesSettings.TaxlotAcreageAnno, True) = 0 Then
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
            MessageBox.Show(ex.Message)
            Return 10 'default
        End Try
    End Function

    ''' <summary>
    ''' Determine the x and y coordinates of the center of envelope, and return them as a Point object
    ''' </summary>
    ''' <param name="envelope">An envelope object of type IEnvelope</param>
    ''' <returns>A Point object that represents the center of the envelope</returns>
    ''' <remarks></remarks>
    Private Shared Function GetCenterOfEnvelope(ByRef envelope As IEnvelope) As IPoint
        Try
            Dim center As IPoint
            center = New Point
            center.X = envelope.XMin + (envelope.XMax - envelope.XMin) / 2
            center.Y = envelope.YMin + (envelope.YMax - envelope.YMin) / 2
            Return center
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    '''Private empty constructor to prevent instantiation
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()
    End Sub

#End Region
#End Region
End Class
#End Region