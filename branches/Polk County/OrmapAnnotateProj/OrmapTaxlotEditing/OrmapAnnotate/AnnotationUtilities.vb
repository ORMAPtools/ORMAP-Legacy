#Region "Copyright 2008 ORMAP Tech Group"

' File:  AnnotationUtilities.vb
'
' Original Author:  Robert Gumtow
'
' Date Created:  June 9, 2010
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
Imports System.Windows.Forms
Imports System.Environment
Imports System.Globalization
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.EditorExtension
Imports OrmapTaxlotEditing.Utilities
#End Region

#Region "Class Declaration"
''' <summary>
''' Annotation utilities class (singleton).
''' </summary>
''' <remarks>
''' <para>Commonly used annotation methods and functions.</para>
''' </remarks>

Public NotInheritable Class AnnotationUtilities

#Region "Class-Level Constants and Enumerations"
    '------------------------------------------
    ' Some constants could be redefined as settings
    '------------------------------------------

    Public Const AnnotationClassName As String = "34"
    Public Const Pi As Double = 3.1415926535897931
    Public Const WideLine As Integer = 60

    'Annotation placement constants
    Public Const DistanceBothSides As Double = 2.55
    Public Const DistanceBothAbove As Double = 1.75
    Public Const DistanceBothBelow As Double = 1.99

    Public Enum AnnotationPlacement As Integer
        BothSides
        BothSidesWide
        BothAbove
        BothBelow
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

#Region "Public Members"

    ''' <summary>
    ''' Gets the name of the Annotation Feature Class based on the Map Scale.
    ''' </summary>
    ''' <param name="theMapScale">The Map Scale from which to determine the Annotation Feature Class name.</param>
    ''' <returns>The Annotation Feature Class name as a string.</returns>
    ''' <remarks>Uses ORMAP Settings to find name of appropriate annotation feature class.</remarks>
    Public Shared Function GetAnnoFCName(ByVal theMapScale As String) As String
        'TODO: (RG) Need to put Case Else and exception handler in here
        Dim theAnnoFCName As String = String.Empty
        Select Case CInt(theMapScale)
            Case 120
                theAnnoFCName = EditorExtension.AnnoTableNamesSettings.Anno0010scaleFC
            Case 240
                theAnnoFCName = EditorExtension.AnnoTableNamesSettings.Anno0020scaleFC
            Case 360
                theAnnoFCName = EditorExtension.AnnoTableNamesSettings.Anno0030scaleFC
            Case 480
                theAnnoFCName = EditorExtension.AnnoTableNamesSettings.Anno0040scaleFC
            Case 600
                theAnnoFCName = EditorExtension.AnnoTableNamesSettings.Anno0050scaleFC
            Case 1200
                theAnnoFCName = EditorExtension.AnnoTableNamesSettings.Anno0100scaleFC
            Case 2400
                theAnnoFCName = EditorExtension.AnnoTableNamesSettings.Anno0200scaleFC
            Case 4800
                theAnnoFCName = EditorExtension.AnnoTableNamesSettings.Anno0400scaleFC
            Case 9600
                theAnnoFCName = EditorExtension.AnnoTableNamesSettings.Anno0800scaleFC
            Case 24000
                theAnnoFCName = EditorExtension.AnnoTableNamesSettings.Anno2000scaleFC
        End Select
        Return theAnnoFCName

    End Function

    ''' <summary>
    ''' Gets the Subtype Code for the desired Subtype name. 
    ''' </summary>
    ''' <param name="theAnnoFeatureClass">Name of the field to draw the domain from.</param>
    ''' <param name="theSubtypeName">Name of the Subtype for which you need the Subtype Code.</param>
    ''' <returns>An Symbol Id as an integer.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetSubtypeCode(ByVal theAnnoFeatureClass As IFeatureClass, ByVal theSubtypeName As String, _
                                          Optional ByVal stringLength As Integer = 0) As Integer
        'Cannot use SpatialUtilities.ConvertCodeValueDomainToCode here since this is an annotation subtype, and not a domain.
        Dim theReturnValue As Integer = Nothing
        If stringLength = 0 Then
            stringLength = Len(theSubtypeName)
        End If

        'TODO:  (RG) Need exception handling
        Dim theSubtypes As ISubtypes = DirectCast(theAnnoFeatureClass, ISubtypes)
        Dim theEnumSubtypes As IEnumSubtype = theSubtypes.Subtypes
        Dim thisSubtypeName As String
        Dim thisSubtypeCode As Integer
        theEnumSubtypes.Reset()
        Do Until theEnumSubtypes Is Nothing
            thisSubtypeName = theEnumSubtypes.Next(thisSubtypeCode)
            If String.Compare(thisSubtypeName, 0, theSubtypeName, 0, stringLength, CultureInfo.CurrentCulture, CompareOptions.IgnoreCase) >= 0 Then
                theReturnValue = thisSubtypeCode
                Exit Do
            End If
        Loop
        Return theReturnValue
    End Function

    ''' <summary>
    ''' Gets the Symbol ID value based on the Symbol Name
    ''' </summary>
    ''' <param name="theAnnoFeatureClass">The Annotation Feature Class containing the appropriate symbol set.</param>
    ''' <param name="theSymbolName">The name of the symbol for which the symbol id is being sought.</param>
    ''' <returns>The Symbol Id as an integer.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetSymbolId(ByVal theAnnoFeatureClass As IFeatureClass, ByVal theSymbolName As String) As Integer
        'TODO:  (RG) Need exception handling
        Dim theSymbolCollection As ISymbolCollection2 = New SymbolCollectionClass()
        Dim thisSymbolIdentifier As ISymbolIdentifier2 = New SymbolIdentifierClass()
        Dim theAnnoClass As IAnnoClass
        Dim theReturnValue As Integer = Nothing

        '------------------------------------------
        'Get the Symbol ID 
        '------------------------------------------
        'Defined in the anno feature class for Symbol Name  
        'representing Distance/Bearing (should be "34")
        theAnnoClass = DirectCast(theAnnoFeatureClass.Extension, IAnnoClass)
        theSymbolCollection = DirectCast(theAnnoClass.SymbolCollection, ISymbolCollection2)
        theSymbolCollection.Reset()
        thisSymbolIdentifier = DirectCast(theSymbolCollection.Next, ISymbolIdentifier2)
        Do Until thisSymbolIdentifier Is Nothing
            If String.Compare(thisSymbolIdentifier.Name, theSymbolName, True, CultureInfo.CurrentCulture) = 0 Then
                theReturnValue = thisSymbolIdentifier.ID
                Exit Do
            End If
            thisSymbolIdentifier = DirectCast(theSymbolCollection.Next, ISymbolIdentifier2)
        Loop
        Return theReturnValue
    End Function

    ''' <summary>
    ''' Gets the maximum Object ID value in the Annotation Feature Class
    ''' </summary>
    ''' <param name="theAnnoFeatureClass">The Annotation Feature Class from which to get the maximum Object ID.</param>
    ''' <returns>The Annotation Feature Class's maximum Object Id as an integer.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetMaxOidByAnnoFC(ByVal theAnnoFeatureClass As IFeatureClass) As Integer
        'TODO:  (RG) Should decouple this from annotation and have it return the max OID from any feature class or table
        Dim theMaxOid As Integer = Nothing
        Dim theAnnoDataset As IDataset = DirectCast(theAnnoFeatureClass, IDataset)
        Dim theAnnoWorkspace As IFeatureWorkspace = DirectCast(theAnnoDataset.Workspace, IFeatureWorkspace)
        Dim theAnnoWorkspaceEditControl As IWorkspaceEditControl = DirectCast(theAnnoWorkspace, IWorkspaceEditControl)

        '------------------------------------------
        'Get the max ID from this anno layer 
        '------------------------------------------
        '(it will be the anno just created by the label to anno converter)
        Dim theQueryDef As IQueryDef
        Dim thisRow As IRow
        Dim theIdCursor As ICursor
        theQueryDef = theAnnoWorkspace.CreateQueryDef
        theQueryDef.Tables = theAnnoDataset.Name
        theQueryDef.SubFields = "MAX(OBJECTID)"
        theIdCursor = theQueryDef.Evaluate
        If Not theIdCursor Is Nothing Then
            thisRow = theIdCursor.NextRow
            'TODO:  (RG) Should really do a find for the field index to the OID... its always 0, but should check anyways... 
            theMaxOid = CInt(thisRow.Value(0))
        End If
        Return theMaxOid
    End Function

    ''' <summary>
    ''' Gets the Annotation Feature Class based on the name.
    ''' </summary>
    ''' <param name="theAnnoFCName">The name of the desired Annotation Feature Class.</param>
    ''' <returns>An object which implements IFeatureClass.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetAnnoFeatureClass(ByVal theAnnoFCName As String) As IFeatureClass
        Dim theAnnoEnumLayer As IEnumLayer = SpatialUtilities.GetTOCLayersEnumerator(EsriLayerTypes.FDOGraphicsLayer)
        Dim thisAnnoFeatureClass As IFeatureClass = Nothing
        Dim thisAnnoLayer As IFeatureLayer
        thisAnnoLayer = DirectCast(theAnnoEnumLayer.Next, IFeatureLayer)

        '------------------------------------------
        'Find anno feature class (based on name)
        '------------------------------------------
        Do While Not (thisAnnoLayer Is Nothing)
            If String.Compare(thisAnnoLayer.Name, theAnnoFCName, True, CultureInfo.CurrentCulture) = 0 Then
                thisAnnoFeatureClass = thisAnnoLayer.FeatureClass
                Exit Do
            End If
            thisAnnoLayer = DirectCast(theAnnoEnumLayer.Next, IFeatureLayer)
        Loop

        Return thisAnnoFeatureClass
    End Function

    ''' <summary>
    ''' Moves an annotation element based on its TO and FROM point values
    ''' </summary>
    ''' <param name="theToPoint">The point to be used as the "to" point of the move vector.</param>
    ''' <param name="theFromPoint">The point to be used as the "from" point of the move vector.</param>
    ''' <param name="theGraphicsContainer">The graphics container which holds the annotation element.</param>
    ''' <param name="theElement">The annotation element to be moved.</param>
    ''' <remarks>Moves the element based on "to" and "from" points</remarks>
    Public Overloads Shared Sub MoveElement(ByVal theToPoint As IPoint, ByVal theFromPoint As IPoint, ByVal theGraphicsContainer As IGraphicsContainer, ByVal theElement As IElement)
        '------------------------------------------
        'Create a vector line between the two points 
        '------------------------------------------
        Dim theMoveVector As ILine = New ESRI.ArcGIS.Geometry.Line
        theMoveVector.PutCoords(theToPoint, theFromPoint)

        '------------------------------------------
        'Call the overloaded MoveElement function with the vector
        '------------------------------------------
        MoveElement(theMoveVector, theGraphicsContainer, theElement)

    End Sub

    ''' <summary>
    ''' Moves an annotation element based on a move vector
    ''' </summary>
    ''' <param name="theMoveVector">The move vector which defines the element's movement.</param>
    ''' <param name="theGraphicsContainer">The graphics container which holds the annotation element.</param>
    ''' <param name="theElement">The annotation element to be moved.</param>
    ''' <remarks>Moves the element based on a move vector.</remarks>
    Public Overloads Shared Sub MoveElement(ByVal theMoveVector As ILine, ByVal theGraphicsContainer As IGraphicsContainer, ByVal theElement As IElement)
        '------------------------------------------
        'Move the annotation using ITansform2D
        '------------------------------------------
        Dim transform2D As ITransform2D = DirectCast(theElement, ITransform2D)
        transform2D.MoveVector(theMoveVector)

        '------------------------------------------
        'Transformation complete
        '------------------------------------------
        'Update the source objects with the new geometry
        finishTransform(theGraphicsContainer, theElement, transform2D)

    End Sub

    ''' <summary>
    ''' Rotates an annotation element around a specified point.
    ''' </summary>
    ''' <param name="theRotationPoint">The move vector which defines the element's movement.</param>
    ''' <param name="theGraphicsContainer">The graphics container which holds the annotation element.</param>
    ''' <param name="theElement">The annotation element to be rotated.</param>
    ''' <remarks>Rotates the annotation element 180 degrees (pi) around a specified point.</remarks>
    Public Shared Sub RotateElement(ByVal theRotationPoint As IPoint, ByVal theGraphicsContainer As IGraphicsContainer, ByVal theElement As IElement)
        '------------------------------------------
        'Rotate the annotation 180 degrees (pi)
        '------------------------------------------
        Dim transform2D As ITransform2D = DirectCast(theElement, ITransform2D)
        transform2D.Rotate(theRotationPoint, Pi)

        '------------------------------------------
        'Transformation complete, now update the source objects with the new geometry
        '------------------------------------------
        finishTransform(theGraphicsContainer, theElement, transform2D)

    End Sub

    ''' <summary>
    ''' Rotates an annotation element around a specified point.
    ''' </summary>
    ''' <param name="theDisplay">The move vector which defines the element's movement.</param>
    ''' <param name="theElement">The annotation element to be rotated.</param>
    ''' <returns>An object implementing IEnvelope.</returns>
    ''' <remarks>Rotates the annotation element 180 degrees (pi) around a specified point.</remarks>
    Public Shared Function GetAnnoEnvelope(ByVal theDisplay As IDisplay, ByVal theElement As IElement) As IEnvelope
        Dim theBoundaryEnvelope As IEnvelope = New Envelope
        theElement.QueryBounds(theDisplay, theBoundaryEnvelope)
        Return theBoundaryEnvelope
    End Function

    Public Shared Sub MoveAnnotationElements(ByVal isInverted As Boolean, ByVal isTransposed As Boolean, ByVal isMoveUp As Boolean, _
                                                ByVal isMoveDown As Boolean, ByVal isStandardSpace As Boolean, ByVal isWideSpace As Boolean)
        '------------------------------------------
        ' -- DESIGN COMMENT --
        '   NOTE=>  Annotation movement only works for annotation created and placed by the CreateAnnotation class. This class 
        '           places annotation based on a set of rules which comply with annotation spacing, fonts, and font sizes used
        '           in Polk County taxmaps. 
        '
        '   Moving distance and bearing annotation is accomplished by describing mathematical relationships relating to a "virtual"  
        '   line (the taxlot line). Distances between the "upper" and "lower" annotation varies depending on its relative location 
        '   to this line. Annotation on different sides of the line, both above the line, or both below the line have different 
        '   distances between them. These distances are calculated relative to the centroid of the annotation's envelopes, and are 
        '   dependent on font size. These distances also vary if the line is "standard" or "big" (wide). Annotation may be placed 
        '   with wide spacing if it will be located next to map index boundaries, taxcode lines, subdivision boundaries, etc. 
        '
        '   Constants were calculated by checking envelope centroid distances for a number of annotation sets of different scales 
        '   placed above, below, and on both sides of the line at 100, 200, and 400 scales. These constants were calculated 
        '   from the formula:
        '
        '       K = Distance / FontSize
        '
        '   and work out to the following values:
        '
        '           Anno Location           K
        '           --------------         ----
        '           Both Sides             2.55   
        '           Both Above             1.75
        '           Both Below             1.99
        '
        '------------------------------------------

        Dim theMxDoc As IMxDocument
        Dim theMap As IMap
        theMxDoc = DirectCast(EditorExtension.Application.Document, IMxDocument)
        theMap = theMxDoc.FocusMap

        Dim theActiveView As IActiveView = DirectCast(theMap, IActiveView)
        Dim theMapExtent As IEnvelope = theActiveView.Extent

        DataMonitor.CheckValidMapIndexDataProperties()

        '------------------------------------------
        'Get the annotation layer
        '------------------------------------------
        Dim theAnnoEnumLayer As IEnumLayer = SpatialUtilities.GetTOCLayersEnumerator(EsriLayerTypes.FDOGraphicsLayer)
        Dim theAnnoFeatureClass As IFeatureClass = Nothing
        Dim thisAnnoLayer As IFeatureLayer
        thisAnnoLayer = DirectCast(theAnnoEnumLayer.Next, IFeatureLayer)

        '------------------------------------------
        'Now go through anno feature class enum looking for selected features
        '------------------------------------------
        Do While Not (thisAnnoLayer Is Nothing)
            Dim theAnnoSelection As IFeatureSelection = DirectCast(thisAnnoLayer, IFeatureSelection)
            If theAnnoSelection.SelectionSet.Count Mod 2 > 0 Then
                MessageBox.Show("Cannot Move Annotation: Odd number of annotation items selected... " & NewLine & _
                                "This tool works with pairs of annotation, so you must select them in " & NewLine & _
                                "sets of two (within each annotation feature class).", _
                                "Move Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            End If
            Dim theSelectedAnnoCursor As IFeatureCursor = SpatialUtilities.GetSelectedFeatures(thisAnnoLayer)
            If Not theSelectedAnnoCursor Is Nothing Then
                'NOTE:  Each annotation feature class will have all of its features moved as a single edit operation.
                EditorExtension.Editor.StartOperation()

                '------------------------------------------
                'Get the graphics container 
                '------------------------------------------
                'Without this, the transform2D will transform the envelope, but NOT the annotation itself!!!!
                Dim theGraphicsContainer As IGraphicsContainer = DirectCast(theMxDoc.ActivatedView, IGraphicsContainer)

                '------------------------------------------
                'Now loop through the selection set
                '------------------------------------------
                'Since actions work on pairs of anno, only do half the loops
                Dim i As Integer
                For i = 1 To theAnnoSelection.SelectionSet.Count Step 2
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' -- DESIGN COMMENT --
                    '   There are a number of ways this next section can be written, but due to the fairly convoluted concept of 
                    '   moving two pieces of annotation, I decided to leave it this way. I almost never use object names which imply
                    '   ordinal relationships (such as 1st, 2nd, etc.), but in this case it clarifies the underlying process since
                    '   this code must work on annotation pairs and actions are based entirely on which element is selected first.
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '------------------------------------------
                    'Get the selected features 
                    '------------------------------------------
                    'Needed for the transform2d since it accesses the underlying geometry of the annotation
                    Dim the1stSelectedFeature As IFeature = theSelectedAnnoCursor.NextFeature
                    Dim the2ndSelectedFeature As IFeature = theSelectedAnnoCursor.NextFeature

                    If the1stSelectedFeature.OID = the2ndSelectedFeature.OID Then
                        MessageBox.Show("Cannot Move Annotation: ESRI BUG- All Annotation Classes " & NewLine & _
                                        "MUST be turned on in the TOC. Turn all Annotation Classes  " & NewLine & _
                                        "back on to bypass this bug. If this does not resolve issue,  " & NewLine & _
                                        "use [Editor > Options  > ORMAP Taxlot Editor and click" & NewLine & _
                                        "[ Report Bug or Request New Feature ].", _
                                        "Move Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        Exit Sub
                    End If

                    '------------------------------------------
                    'Get the selected annotation features annotation elements 
                    '------------------------------------------
                    'Needed to do transform and update the graphics container
                    Dim the1stSelectedAnnoFeature As IAnnotationFeature = DirectCast(the1stSelectedFeature, IAnnotationFeature)
                    Dim the2ndSelectedAnnoFeature As IAnnotationFeature = DirectCast(the2ndSelectedFeature, IAnnotationFeature)

                    If isInverted Then
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ' -- DESIGN COMMENT --
                        '   The envelope used to calculate the rotation point (centroid) must be retrieved from an element's QueryBounds method. 
                        '   This method takes into account the annotation's text area, whereas the element's envelope is a polyline. Using
                        '   its centroid will actually offset the location of the annotation since the polyline's MaxY is just beneath the text.
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '------------------------------------------
                        'To invert the annotation as a pair, must rotate each element then transpose them
                        '------------------------------------------
                        RotateElement(GetCenterOfEnvelope(GetAnnoEnvelope(theActiveView.ScreenDisplay, the1stSelectedAnnoFeature.Annotation)), _
                                        theGraphicsContainer, the1stSelectedAnnoFeature.Annotation)
                        RotateElement(GetCenterOfEnvelope(GetAnnoEnvelope(theActiveView.ScreenDisplay, the2ndSelectedAnnoFeature.Annotation)), _
                                        theGraphicsContainer, the2ndSelectedAnnoFeature.Annotation)
                    End If

                    '------------------------------------------
                    'Calculate a move vector  
                    '------------------------------------------
                    'from the X, Y pairs of the 1st and 2nd element's envelopes (mean of lower left & upper right corners)
                    Dim theMoveVector As ILine = New ESRI.ArcGIS.Geometry.Line

                    theMoveVector.PutCoords(GetCenterOfEnvelope(GetAnnoEnvelope(theActiveView.ScreenDisplay, the1stSelectedAnnoFeature.Annotation)), _
                                            GetCenterOfEnvelope(GetAnnoEnvelope(theActiveView.ScreenDisplay, the2ndSelectedAnnoFeature.Annotation)))

                    Dim theTextElement As ITextElement = DirectCast(the1stSelectedAnnoFeature.Annotation, ITextElement)
                    Dim theAnnoPlacement As AnnotationPlacement = GetAnnoPlacement(theTextElement.Symbol.Size, theMoveVector.Length)
                    If theAnnoPlacement = -1 And Not isTransposed Then
                        MessageBox.Show("Cannot Move Annotation: Selected annotation is at non-standard" & NewLine & _
                                        "placement (was not created by Create Annotation tool or has" & NewLine & _
                                        "been moved). Placement tools cannot reposition annotation.", _
                                        "Move Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        Exit For
                    End If
                    Dim theToDistance As Double
                    Dim theNewVector As ILine = New ESRI.ArcGIS.Geometry.Line

                    'NOTE=> All movement is from theMoveVector's 'From' end to its 'To' end, so resized vectors utilize
                    '       the esriExtendAtFrom constant with a recalculated 'To' point in the QueryTangent method

                    'TODO:  (RG)- A lot of redundant code here... should be refactored

                    '------------------------------------------
                    'Move the element
                    '------------------------------------------
                    'Each action below is exclusionary, so placed in nested If-Then-ElseIf blocks
                    If isTransposed Then
                        MoveElement(theMoveVector, theGraphicsContainer, the1stSelectedAnnoFeature.Annotation)
                        theMoveVector.ReverseOrientation()
                        MoveElement(theMoveVector, theGraphicsContainer, the2ndSelectedAnnoFeature.Annotation)
                    ElseIf isMoveDown Then
                        MoveElement(theMoveVector, theGraphicsContainer, the1stSelectedAnnoFeature.Annotation)
                        If isStandardSpace Then
                            theToDistance = DistanceBothSides * theTextElement.Symbol.Size
                            theMoveVector.QueryTangent(esriSegmentExtension.esriExtendAtFrom, 0, False, theToDistance, theNewVector)
                            theMoveVector = theNewVector
                        ElseIf isWideSpace Then
                            theToDistance = DistanceBothSides * theTextElement.Symbol.Size + WideLine
                            theMoveVector.QueryTangent(esriSegmentExtension.esriExtendAtFrom, 0, False, theToDistance, theNewVector)
                            theMoveVector = theNewVector
                        ElseIf theAnnoPlacement = AnnotationPlacement.BothSides Or theAnnoPlacement = AnnotationPlacement.BothSidesWide Then
                            theToDistance = DistanceBothBelow * theTextElement.Symbol.Size
                            theMoveVector.QueryTangent(esriSegmentExtension.esriExtendAtFrom, 0, False, theToDistance, theNewVector)
                            theMoveVector = theNewVector
                        End If
                        MoveElement(theMoveVector, theGraphicsContainer, the2ndSelectedAnnoFeature.Annotation)
                    ElseIf isMoveUp Then
                        theMoveVector.ReverseOrientation()
                        MoveElement(theMoveVector, theGraphicsContainer, the2ndSelectedAnnoFeature.Annotation)
                        If isStandardSpace Then
                            theToDistance = DistanceBothSides * theTextElement.Symbol.Size
                            theMoveVector.QueryTangent(esriSegmentExtension.esriExtendAtFrom, 0, False, theToDistance, theNewVector)
                            theMoveVector = theNewVector
                        ElseIf isWideSpace Then
                            theToDistance = DistanceBothSides * theTextElement.Symbol.Size + WideLine
                            theMoveVector.QueryTangent(esriSegmentExtension.esriExtendAtFrom, 0, False, theToDistance, theNewVector)
                            theMoveVector = theNewVector
                        ElseIf theAnnoPlacement = AnnotationPlacement.BothSides Or theAnnoPlacement = AnnotationPlacement.BothSidesWide Then
                            theToDistance = DistanceBothAbove * theTextElement.Symbol.Size
                            theMoveVector.QueryTangent(esriSegmentExtension.esriExtendAtFrom, 0, False, theToDistance, theNewVector)
                            theMoveVector = theNewVector
                        End If
                        MoveElement(theMoveVector, theGraphicsContainer, the1stSelectedAnnoFeature.Annotation)
                    End If

                Next
                theActiveView.Refresh()
                '------------------------------------------
                'Label and close the edit operation
                '------------------------------------------
                If isInverted Then
                    EditorExtension.Editor.StopOperation("Rotate Annotation")
                ElseIf isTransposed Then
                    EditorExtension.Editor.StopOperation("Flip Annotation")
                ElseIf isMoveUp Then
                    EditorExtension.Editor.StopOperation("Move Annotation Up")
                ElseIf isMoveDown Then
                    EditorExtension.Editor.StopOperation("Move Annotation Down")
                End If
            End If
            thisAnnoLayer = DirectCast(theAnnoEnumLayer.Next, IFeatureLayer)
        Loop
    End Sub

    Public Shared Function GetAnnoPlacement(ByVal theFontSize As Double, ByVal theCalculatedDistance As Double) As AnnotationPlacement
        Dim thisAnnoPlacement As AnnotationPlacement
        '------------------------------------------
        'Checks annotation placement 
        '------------------------------------------
        'Based on the scale constants and the font size for the annotation feature class
        If DistanceBothSides = calculateAnnoSpacing(theCalculatedDistance, theFontSize) Then
            thisAnnoPlacement = AnnotationPlacement.BothSides
        ElseIf DistanceBothAbove = calculateAnnoSpacing(theCalculatedDistance, theFontSize) Then
            thisAnnoPlacement = AnnotationPlacement.BothAbove
        ElseIf DistanceBothBelow = calculateAnnoSpacing(theCalculatedDistance, theFontSize) Then
            thisAnnoPlacement = AnnotationPlacement.BothBelow
        ElseIf DistanceBothSides = calculateAnnoSpacing(theCalculatedDistance - WideLine, theFontSize) Then
            thisAnnoPlacement = AnnotationPlacement.BothSidesWide
        Else
            thisAnnoPlacement = DirectCast(-1, AnnotationPlacement)
        End If
        Return thisAnnoPlacement
    End Function

    Public Shared Function GetMoveDistance(ByVal isStandardSpace As Boolean, ByVal isBothSides As Boolean, ByVal theTextElement As ITextElement) As Double
        '------------------------------------------
        'Currently unused... 
        '------------------------------------------
        'This code was written to handle moving annotation when cartographer is using asymmetric offsets (top or bottom only)
        'and may be needed for the next phase of the OrmapAnnotation project

        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '!!!  NOTE=> Code NOT COMPLETE  !!!
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

        Dim theLineWidth As Double = New Double

        'Polk County uses Font Sizes = MapScale/240
        'The cartographers use an additional map unit offset for "big" lines 
        '    100     3
        '    200     6
        '    400    12
        '    800    24
        '   1000    48
        '   2000    60
        If isStandardSpace Then
            If isBothSides Then
                theLineWidth = theTextElement.Symbol.Size * 2
            ElseIf Not isBothSides Then
                theLineWidth = theTextElement.Symbol.Size * 1.5
            End If
        ElseIf Not isStandardSpace Then
            Dim theScaleOffset As Integer = CInt(theTextElement.Symbol.Size Mod 1.5)
            If isBothSides Then
                theLineWidth = theTextElement.Symbol.Size * 2 + theScaleOffset
            ElseIf Not isBothSides Then
                theLineWidth = theTextElement.Symbol.Size * 1.5 + theScaleOffset
            End If
        End If
        Return theLineWidth
    End Function

#End Region

#Region "Private Members"

    Private Shared Sub finishTransform(ByVal theGraphicsContainer As IGraphicsContainer, ByVal theElement As IElement, ByVal theTransform As ITransform2D)
        '------------------------------------------
        'Update the source objects with the new geometry
        '------------------------------------------
        theElement = DirectCast(theTransform, IElement)
        theGraphicsContainer.UpdateElement(theElement)
    End Sub

    Private Shared Function calculateAnnoSpacing(ByVal theDistance As Double, ByVal theFontSize As Double) As Double
        Dim theAnnoSpacing As Double
        theAnnoSpacing = Math.Round(theDistance / theFontSize, 2, MidpointRounding.AwayFromZero)
        Return theAnnoSpacing
    End Function

#End Region

End Class
#End Region
