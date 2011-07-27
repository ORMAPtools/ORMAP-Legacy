#Region "Copyright 2008 ORMAP Tech Group"

' File:  CreateAnnotation.vb
'
' Original Author:  Robert Gumtow
'
' Date Created:  May 11, 2010
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
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Environment
Imports System.Globalization
Imports System.Drawing.Text
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.DataSourcesGDB
Imports OrmapTaxlotEditingAddIn.DataMonitor
Imports OrmapTaxlotEditingAddIn.SpatialUtilities
Imports OrmapTaxlotEditingAddIn.EditorExtension
Imports OrmapTaxlotEditingAddIn.Utilities
Imports OrmapTaxlotEditingAddIn.AnnotationUtilities

#End Region

''' <summary>
''' Allows creation of Distance and Direction annotation from Line Feature Classes that contain
''' appropriately formatted Distance and Direction attributes.
''' </summary>
''' <remarks>
''' Annotation Feature Classes must have definitions for an Annotation Class named "34" and a
''' Symbol named "34" (with appropriately defined font size and type information). The tool will 
''' use these definitions for new annotation.
''' <para><seealso cref="InvertAnnotation"/></para>
''' <para><seealso cref="TransposeAnnotation"/></para>
''' <para><seealso cref="MoveUp"/></para>
''' <para><seealso cref="MoveDown"/></para>
''' </remarks>
Public NotInheritable Class CreateAnnotation
    Implements IDisposable

#Region "Class-Level Constants and Enumerations"

    ''' <summary>
    ''' Identifies which annotation element of the pair is placed "on top"
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum TopPosition
        ''' <summary>
        ''' The distance value (from the Distance field of the line feature class)
        ''' </summary>
        ''' <remarks></remarks>
        Distance
        ''' <summary>
        ''' The direction (bearing) value (from the Direction field of the line feature class)
        ''' </summary>
        ''' <remarks></remarks>
        Direction
    End Enum

#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        'MyBase.m_category = "OrmapAnnotate"  'localizable text 
        'MyBase.m_caption = "CreateAnnotation"   'localizable text 
        'MyBase.m_message = "Creates Distance and Direction annotation with user-selected placement preferences"   'localizable text 
        'MyBase.m_toolTip = "Create Distance && Direction annotation" 'localizable text 
        'MyBase.m_name = MyBase.m_category & "_CreateAnnotation"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        'Try
        '    Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
        '    MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        'Catch ex As Exception
        '    System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        'End Try


    End Sub
#End Region

#End Region

#Region "Custom Class Members"

#Region "Properties"
    '------------------------------------------
    ' Translate the form controls into Properties
    '------------------------------------------
    ' This is where the Controller (or Presenter) couples the form (View) and the class (Model) 

    Private WithEvents _partnerCreateAnnotationForm As CreateAnnotationForm

    Friend ReadOnly Property PartnerCreateAnnotationForm() As CreateAnnotationForm
        Get
            If _partnerCreateAnnotationForm Is Nothing OrElse _partnerCreateAnnotationForm.IsDisposed Then
                setPartnerCreateAnnotationForm(New CreateAnnotationForm())
            End If
            Return _partnerCreateAnnotationForm
        End Get
    End Property



    ''' <summary>
    ''' Form Class Name setting.
    ''' </summary>
    Private _annoClassName As String
    ''' <summary>
    ''' Gets the name of the anno class.
    ''' </summary>
    ''' <value>
    ''' The name of the anno class.
    ''' </value>
    Public ReadOnly Property AnnoClassName() As String
        Get
            ' Value from constant... could eventually be a setting
            _annoClassName = AnnotationUtilities.AnnotationClassName
            Return _annoClassName
        End Get
    End Property

    ''' <summary>
    ''' Form Curved setting.
    ''' </summary>
    Private _isCurved As Boolean
    ''' <summary>
    ''' Gets a value indicating whether this instance is curved.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if this instance is curved; otherwise, <c>false</c>.
    ''' </value>
    Public ReadOnly Property IsCurved() As Boolean
        Get
            _isCurved = PartnerCreateAnnotationForm.uxCurved.Checked
            Return _isCurved
        End Get
    End Property

    ''' <summary>
    ''' Form Parallel setting.
    ''' </summary>
    Private _isParallel As Boolean
    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is parallel.
    ''' </summary>
    ''' <value>
    ''' <c>true</c> if this instance is parallel; otherwise, <c>false</c>.
    ''' </value>
    Public Property IsParallel() As Boolean
        Get
            _isParallel = PartnerCreateAnnotationForm.uxParallel.Checked
            Return _isParallel
        End Get
        Set(ByVal value As Boolean)
            _isParallel = value
        End Set
    End Property

    ''' <summary>
    ''' Form Horizontal setting.
    ''' </summary>
    Private _isHorizontal As Boolean
    ''' <summary>
    ''' Gets a value indicating whether this instance is horizontal.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is horizontal; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsHorizontal() As Boolean
        Get
            _isHorizontal = PartnerCreateAnnotationForm.uxHorizontal.Checked
            Return _isHorizontal
        End Get
    End Property

    ''' <summary>
    ''' Form  Perpendicular setting.
    ''' </summary>
    Private _isPerpendicular As Boolean
    ''' <summary>
    ''' Gets a value indicating whether this instance is perpendicular.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is perpendicular; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsPerpendicular() As Boolean
        Get
            _isPerpendicular = PartnerCreateAnnotationForm.uxPerpendicular.Checked
            Return _isPerpendicular
        End Get
    End Property

    ''' <summary>
    ''' Form Above setting
    ''' </summary>
    Private _isAbove As Boolean
    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is above.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if this instance is above; otherwise, <c>false</c>.
    ''' </value>
    Public Property IsAbove() As Boolean
        ' Property is set programmatically
        Get
            Return _isAbove
        End Get
        Set(ByVal value As Boolean)
            _isAbove = value
        End Set
    End Property

    ''' <summary>
    ''' Form Both Below setting.
    ''' </summary>
    Private _isBelow As Boolean
    ''' <summary>
    ''' Gets or sets a value indicating whether this instance is below.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if this instance is below; otherwise, <c>false</c>.
    ''' </value>
    Public Property IsBelow() As Boolean
        ' Property is set programmatically
        Get
            Return _isBelow
        End Get
        Set(ByVal value As Boolean)
            _isBelow = value
        End Set
    End Property

    ''' <summary>
    ''' Form Both Sides setting.
    ''' </summary>
    Private _isBothSides As Boolean
    ''' <summary>
    ''' Gets a value indicating whether this instance is on both sides of line.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is on both sides; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsBothSides() As Boolean
        Get
            _isBothSides = PartnerCreateAnnotationForm.uxBothSides.Checked
            Return _isBothSides
        End Get
    End Property

    ''' <summary>
    ''' Form Both Above setting
    ''' </summary>
    Private _isBothAbove As Boolean
    ''' <summary>
    ''' Gets a value indicating whether annotation pairs are both above line.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is both above; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsBothAbove() As Boolean
        Get
            _isBothAbove = PartnerCreateAnnotationForm.uxBothAbove.Checked
            Return _isBothAbove
        End Get
    End Property

    ''' <summary>
    ''' Form Both Below setting
    ''' </summary>
    Private _isBothBelow As Boolean
    ''' <summary>
    ''' Gets a value indicating whether annotation pairs are both below the line.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is both below; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsBothBelow() As Boolean
        Get
            _isBothBelow = PartnerCreateAnnotationForm.uxBothBelow.Checked
            Return _isBothBelow
        End Get
    End Property

    ''' <summary>
    ''' Form Standard Above setting.
    ''' </summary>
    Private _isStandardAbove As Boolean
    ''' <summary>
    ''' Gets a value indicating whether annotation is standard offset above line.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is above standard; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsStandardAbove() As Boolean
        Get
            _isStandardAbove = PartnerCreateAnnotationForm.uxStandardAbove.Checked
            Return _isStandardAbove
        End Get
    End Property

    ''' <summary>
    ''' Form Double Above setting.
    ''' </summary>
    Private _isDoubleAbove As Boolean

    ''' <summary>
    ''' Gets a value indicating whether this instance is double offset above line.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is double above; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsDoubleAbove() As Boolean
        Get
            _isDoubleAbove = PartnerCreateAnnotationForm.uxDoubleAbove.Checked
            Return _isDoubleAbove
        End Get
    End Property

    ''' <summary>
    ''' Form Standard Below setting.
    ''' </summary>
    Private _isStandardBelow As Boolean
    ''' <summary>
    ''' Gets a value indicating whether annotation is standard offset below line.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is below standard; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsStandardBelow() As Boolean
        Get
            _isStandardBelow = PartnerCreateAnnotationForm.uxStandardBelow.Checked
            Return _isStandardBelow
        End Get
    End Property

    ''' <summary>
    ''' Form Double Below setting.
    ''' </summary>
    Private _isDoubleBelow As Boolean
    ''' <summary>
    ''' Gets a value indicating whether this instance is double offset below line.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is double below; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsDoubleBelow() As Boolean
        Get
            _isDoubleBelow = PartnerCreateAnnotationForm.uxDoubleBelow.Checked
            Return _isDoubleBelow
        End Get
    End Property

    ''' <summary>
    ''' Form Standard line setting.
    ''' </summary>
    Private _isStandardLine As Boolean
    ''' <summary>
    ''' Gets a value indicating whether this instance is standard line.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is standard line; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsStandardLine() As Boolean
        Get
            _isStandardLine = PartnerCreateAnnotationForm.uxStandardLine.Checked
            Return _isStandardLine
        End Get
    End Property

    ''' <summary>
    ''' Form Wide line setting.
    ''' </summary>
    Private _wideLine As Boolean
    ''' <summary>
    ''' Gets a value indicating whether this instance is wide line.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is wide line; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsWideLine() As Boolean
        Get
            _wideLine = PartnerCreateAnnotationForm.uxWideLine.Checked
            Return _wideLine
        End Get
    End Property

    ''' <summary>
    ''' Form "Direction" setting.
    ''' </summary>
    Private _isDirection As Boolean
    ''' <summary>
    ''' Gets a value indicating whether this instance is "Direction".
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is Direction; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsDirection() As Boolean
        Get
            _isDirection = PartnerCreateAnnotationForm.uxDirection.Checked
            Return _isDirection
        End Get
    End Property

    ''' <summary>
    ''' Form "Distance" setting.
    ''' </summary>
    Private _isDistance As Boolean
    ''' <summary>
    ''' Gets a value indicating whether this instance is Distance.
    ''' </summary>
    ''' <value>
    ''' 
    ''' <c>true</c> if this instance is Distance; otherwise, <c>false</c>.
    ''' 
    ''' </value>
    Public ReadOnly Property IsDistance() As Boolean
        Get
            _isDistance = PartnerCreateAnnotationForm.uxDistance.Checked
            Return _isDistance
        End Get
    End Property

    ''' <summary>
    ''' Form Reference scale value.
    ''' </summary>
    Private _referenceScale As Integer
    ''' <summary>
    ''' Gets the reference scale.
    ''' </summary>
    ''' 
    Public ReadOnly Property ReferenceScale() As Integer
        Get
            _referenceScale = CInt(PartnerCreateAnnotationForm.uxReferenceScale.Text)
            Return _referenceScale
        End Get
    End Property

    ''' <summary>
    ''' Form Annotation Feature Layer value.
    ''' </summary>
    Private _annoFeatureLayer As IFeatureLayer
    Public Property AnnoFeatureLayer() As IFeatureLayer
        Get
            Return _annoFeatureLayer
        End Get
        Set(ByVal value As IFeatureLayer)
            _annoFeatureLayer = value
        End Set
    End Property

    ''' <summary>
    ''' Form Upper Value setting
    ''' </summary>
    Private _upperValue As TopPosition
    ''' <summary>
    ''' Gets the upper value.
    ''' </summary>
    ''' 
    Public ReadOnly Property UpperValue() As TopPosition
        ' Property is set programmatically; enum allows for easier processing later
        Get
            If IsDistance Then
                _upperValue = TopPosition.Distance
            ElseIf IsDirection Then
                _upperValue = TopPosition.Direction
            End If
            Return _upperValue
        End Get
    End Property

    'Private _progressBar As System.Windows.Forms.ProgressBar
    'Public Property ProgressBar() As System.Windows.Forms.ProgressBar
    '    Get
    '        Return _progressBar
    '    End Get
    '    Set(ByVal value As System.Windows.Forms.ProgressBar)
    '        _progressBar = value
    '    End Set
    'End Property

#End Region

#Region "Event Handlers"

    ''' <summary>
    ''' Handles [Cancel] button click
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub uxCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerCreateAnnotationForm.Cancel.Click
        PartnerCreateAnnotationForm.Close()
    End Sub

    ''' <summary>
    ''' Handles [Help] button click
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' Opens the Ormap Annotate help documentation
    ''' </remarks>
    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim theRTFStream As System.IO.Stream = Me.GetType().Assembly.GetManifestResourceStream("OrmapTaxlotEditingAddIn.CreateAnnotation_help.rtf")
        OpenHelp("Create Annotation Help", theRTFStream)
    End Sub

    ''' <summary>
    ''' Handles the [Create Annotation] button click event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' Implements the IConvertLabelsToAnnotation interface to collect Distance and Direction values from the line
    ''' feature class and convert them to properly scaled and symbolized ORMAP annotation. Conversion occurs in a
    ''' scratch file geodatabased, and then is copied into the appropriate annotation tables
    ''' </remarks>
    Private Sub uxCreateAnno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim theScratchWorkspaceFactory As IScratchWorkspaceFactory2 = New FileGDBScratchWorkspaceFactoryClass
        Dim theScratchAnnoWorkspace As IFeatureWorkspace = CType(theScratchWorkspaceFactory.CreateNewScratchWorkspace, IFeatureWorkspace)

        Dim theGeometry As IGeometry
        Dim theAnnoFCName As String

        'TODO:  Put a wait cursor in here... 
        'PartnerCreateAnnotationForm.UseWaitCursor = False

        Dim theMxDoc As IMxDocument = My.Document
        Dim theMap As IMap = theMxDoc.FocusMap

        'Set the reference (Line offsets are based on 1:1200 scale default)
        theMap.MapScale = ReferenceScale

        Dim theActiveView As IActiveView = DirectCast(theMap, IActiveView)
        'Dim theMapExtent As IEnvelope = theActiveView.Extent
        Dim theEnumLayer As IEnumLayer = SpatialUtilities.GetTOCLayersEnumerator(EsriLayerTypes.FeatureLayer)

        'Store all selected features in a collection and iterate through the collection
        'selecting each feature in the collection using the SpatialUtilities.SetSelectedFeature method
        Dim theEnumFeature As IEnumFeature = DirectCast(theMap.FeatureSelection, IEnumFeature)
        Dim theEnumFeatureSetup As IEnumFeatureSetup = DirectCast(theEnumFeature, IEnumFeatureSetup)
        theEnumFeatureSetup.AllFields = True
        Dim thisLineFeature As IFeature = theEnumFeature.Next
        Dim theLineFeatureCollection As Collection = New Collection
        Do While Not thisLineFeature Is Nothing
            theLineFeatureCollection.Add(thisLineFeature)
            thisLineFeature = theEnumFeature.Next
        Loop

        PartnerCreateAnnotationForm.uxProgressBar.Maximum = 1000
        PartnerCreateAnnotationForm.uxProgressBar.Minimum = 0
        PartnerCreateAnnotationForm.uxProgressBar.Step = CInt((1000 / theLineFeatureCollection.Count) / 5)
        PartnerCreateAnnotationForm.UseWaitCursor = True
        PartnerCreateAnnotationForm.Cursor = Cursors.WaitCursor

        clearall(theMxDoc)
        theActiveView.Refresh()

        Dim theAnnoFcCollection As Collection = New Collection

        'Prime the layer for looping
        Dim theLayer As IFeatureLayer = DirectCast(theEnumLayer.Next, IFeatureLayer)
        Dim theConverterFeatureSelection As IFeatureSelection = CType(theLayer, IFeatureSelection)

        Dim thisFeatureIndex As Integer = 1

        Do Until theLineFeatureCollection.Count = 0
            thisLineFeature = CType(theLineFeatureCollection.Item(thisFeatureIndex), IFeature)

            'Check to see if the current feature is in the current layer. If not, reset the layer enum
            If Not theLayer.FeatureClass.Equals(thisLineFeature.Class) Then
                theEnumLayer.Reset()
                theLayer = DirectCast(theEnumLayer.Next, IFeatureLayer)
            End If
            'Loop through the layers
            Do Until (theLayer Is Nothing)
                If theLayer.FeatureClass.Equals(thisLineFeature.Class) Then
                    'Set the selected feature from theFeatureCollection
                    SpatialUtilities.SetSelectedFeature(theLayer, thisLineFeature, True, True)

                    'Verify that Distance and Direction attributes are present 
                    If thisLineFeature.Fields.FindField("Direction") < 0 Or thisLineFeature.Fields.FindField("Distance") < 0 Then
                        MessageBox.Show("Missing data: Direction and/or Distance attributes are missing" & NewLine & _
                                        "from the selected feature in " + theLayer.Name + ".", _
                                        "Create Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        Exit Sub
                    End If

                    Dim theGeoFeatureLayer As IGeoFeatureLayer = DirectCast(theLayer, IGeoFeatureLayer)
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '--DESIGN COMMENT--
                    'Clear the selected features (done for the conversion engine... can only use one of three options):
                    '   1- Convert all features in the layer
                    '   2- Convert all selected features in the layer
                    '   3- Convert all features in the current extent
                    '
                    '   Can't convert all selected features because they may be from different MapScales and need to be placed into
                    '   different Annotation Feature Classes. Can't rely on current extent because it almost always includes pieces
                    '   of other features. So clear the selections, then reselect them feature-by-feature from the cursor
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    theConverterFeatureSelection = CType(theLayer, IFeatureSelection)
                    Dim theFeatureClass As IFeatureClass = CType(thisLineFeature.Class, IFeatureClass)

                    Dim distanceIndex As Integer = theFeatureClass.FindField("Distance")
                    Dim directionIndex As Integer = theFeatureClass.FindField("Direction")

                    'Verify selected feature is a line feature type
                    If Not thisLineFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolyline And _
                        Not thisLineFeature.Shape.GeometryType = esriGeometryType.esriGeometryLine Then
                        MessageBox.Show("Wrong Type: A feature was selected which is NOT a line feature." & NewLine & _
                                        "Only line features can be used for Distance and Bearing annotation." & NewLine & _
                                        "Annotation from this feature will be skipped.", _
                                        "Create Annotation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        theConverterFeatureSelection.Clear()
                        theConverterFeatureSelection.SelectionChanged()
                        theLineFeatureCollection.Remove(thisFeatureIndex)
                    ElseIf IsDBNull(thisLineFeature.Value(directionIndex)) Or IsDBNull(thisLineFeature.Value(distanceIndex)) Then
                        MessageBox.Show("Missing data: Direction and/or Distance attribute is <Null> or" & NewLine & _
                                        "missing from the selected feature. Feature " & thisLineFeature.OID & " will not be processed.", _
                                        "Save Cogo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        theConverterFeatureSelection.Clear()
                        theConverterFeatureSelection.SelectionChanged()
                        theLineFeatureCollection.Remove(thisFeatureIndex)
                    Else
                        theGeometry = thisLineFeature.Shape

                        Dim theMapScale As String = GetValue(theGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapScaleField, EditorExtension.MapIndexSettings.MapNumberField)
                        Dim theMapNumber As String = GetValue(theGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapNumberField, EditorExtension.MapIndexSettings.MapNumberField)

                        'Get annoFC based on MapScale
                        theAnnoFCName = GetAnnoFeatureClassName(theMapScale)

                        Dim theAnnoFeatureClass As IFeatureClass = GetAnnoFeatureClass(theAnnoFCName)
                        Dim theAnnoDataSet As IDataset = CType(theAnnoFeatureClass, IDataset)

                        If theAnnoFeatureClass Is Nothing Then
                            MessageBox.Show("Missing data: The annotation feature class" & NewLine & _
                                theAnnoFCName & " is not loaded." & NewLine & _
                                "Please load this dataset into your map.", _
                                "Create Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                            Exit Sub
                        End If
                        Dim theFeatureLayerPropsCollection As IAnnotateLayerPropertiesCollection2
                        theFeatureLayerPropsCollection = CType(theGeoFeatureLayer.AnnotationProperties, IAnnotateLayerPropertiesCollection2)
                        theFeatureLayerPropsCollection.Clear()

                        setLabelProperties(theFeatureLayerPropsCollection, theAnnoFeatureClass, "Direction", theMapNumber)
                        setLabelProperties(theFeatureLayerPropsCollection, theAnnoFeatureClass, "Distance", theMapNumber)
                        convertLabelsToAnnotation(theMap, theGeoFeatureLayer, theAnnoFeatureClass, False, theScratchAnnoWorkspace)

                        PartnerCreateAnnotationForm.uxProgressBar.PerformStep()

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '--DESIGN COMMENT--
                        'Call to processNewAnnotation was originally set here, but opening and closing the edit session destroyed
                        'theSelectedFeaturesCursor pointer even though that cursor is completely unrelated to the insert cursor 
                        'created in processNewAnnotation (NextFeature would fail, sometimes returning a record twice, and then
                        'skipping the final record). Spent a lot of time messing around with selection sets as well, but ended up
                        'with same problem. 

                        'To solve, had to build a dictionary of the annotation feature classes being updated by the conversion
                        'engine. This is used later for processing all of the new annotation added to each anno feature class.
                        'Becuase the annotation feature classes retain deleted OID values, the collection stores the OID minus 1 (because
                        'the last two OIDs were inserted by convertLabelsToGDBAnnotation for Distance and Direction which uses the
                        'next Oid from the sequence (even if earlier Oid's were deleted). 

                        'UPDATE 7/20/2010- Spoke to programmer at ESRI IUC in San Diego. Was told that issue relates to symbol definition.
                        'Basically, a symbol must be defined in the Line Feature Class that EXACTLY matches the symbol defined in the
                        'Annotation Feature Class for the specified annotation class. This is not practical for this situation since
                        'no symbols are defined in the Line Feature Class since the symbol is scale dependent. So this workaround is
                        'still the best solution. 

                        'UPDATE 8/2/2010- Moving processNewAnnotation outside of selected feature loop caused issue with converter
                        'creating sequenced sets of Distance and Direction annotation classes and symbols (E.g., Distance, Distance_1,
                        'Distance_2, ... , Distance_999). At Distance_999, system would crash. Created exception to verify less than
                        '999 features have been selected (technically, it could be 999 features per annotation feature class, but
                        'for now left at 999 selected features total). Also, created new method cleanNewAnnotation which will remove
                        'all Distance and Direction subtypes, annotation classes, and symbols. 

                        'UPDATE 8/19/2010- Encountered issues with different code behavior in Personal, File, and SDE geodatabases.
                        'In SDE, none of the auto fields were being populated (create event wasn't being fired...) so had to re-
                        'structure by moving processNewAnnotation back into feature loop (see note above). The edit session in 
                        'processNewAnnotation clears the selection set, so had to first capture all features selected by the user
                        '(put in allSelectedFeatures before first edit session called). After annotation creation, the feature just
                        'processed is removed from the selection, processNewAnnotion is called (clearing all selected features), and
                        'allSelectedFeatures.SelectFeatures restores all of the originally selected features (except the one just
                        'processed). 

                        'UPDATE 8/27/2010- CreateAnnotation class was reworked extensively to get annotation working in Personal, File,
                        'and SDE databases. This has been completed, as well as tuning the cleanNewAnnotation method. Note that this
                        'class needs serious refactoring, but that will have to be done at a later date. It works, appears to be fairly
                        'stable (although its slower than the earlier version since it has to run every selected feature through the
                        'layer enumeration to reset it...)

                        'UPDATE 4/5/2011- Got code from Jeff Schenke (LCOG) showing how the Symbol Collection works... using the method
                        'solved issues relating to bad subtypes and symbols created by 
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        If theAnnoFcCollection.Count = 0 Or Not theAnnoFcCollection.Contains(theAnnoDataSet.Name) Then
                            theAnnoFcCollection.Add(theAnnoFeatureClass, theAnnoDataSet.Name)
                        End If
                        theConverterFeatureSelection.Clear()
                        theConverterFeatureSelection.SelectionChanged()
                        theLineFeatureCollection.Remove(thisFeatureIndex)
                    End If
                    Exit Do
                End If
                theLayer = CType(theEnumLayer.Next, IFeatureLayer)
            Loop
        Loop

        processNewAnnotation(theAnnoFcCollection, theScratchAnnoWorkspace)
        PartnerCreateAnnotationForm.Cursor = Cursors.Default

        PartnerCreateAnnotationForm.UseWaitCursor = False

        theActiveView.Refresh()

    End Sub

#End Region

#Region "Methods"

    ''' <summary>
    ''' The form entrance operation called by the AddIn button OnClick event
    ''' </summary>
    Friend Sub DoButtonOperation()
        Try
            'Dim theMxDoc As IMxDocument = DirectCast(EditorExtension.Application.Document, IMxDocument)
            Dim theMxDoc As IMxDocument = My.Document
            Dim theMap As IMap = theMxDoc.FocusMap
            DataMonitor.CheckAllValidDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & NewLine & _
                                "Please load this dataset into your map.", _
                                "Create Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            ElseIf theMap.SelectionCount < 1 Then
                MessageBox.Show("Missing data: No line features have been selected." & NewLine & _
                                "Please load this dataset into your map." & NewLine & _
                                "Distance and Direction attributes.", _
                                "Create Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            ElseIf Not HasValidTaxlotData Then
                MessageBox.Show("Missing data: Valid ORMAP Taxlot layer not found in the map." & NewLine & _
                                "Please select at least one line feature which has." & NewLine & _
                                "Distance and Direction attributes.", _
                                "Create Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            ElseIf theMap.SelectionCount > 999 Then
                MessageBox.Show("Too much data: You have selected more than 999 features. The" & NewLine & _
                                "system cannot process more than 999 line features at a time", _
                                "Create Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            End If

            PartnerCreateAnnotationForm.ShowDialog()

        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try

    End Sub

    ''' <summary>
    ''' Creates the controller for the Create Annotation form and sets up event handler delegates
    ''' </summary>
    ''' <param name="value">A form of type CreateAnnotationForm</param>
    Private Sub setPartnerCreateAnnotationForm(ByVal value As CreateAnnotationForm)
        If value IsNot Nothing Then
            _partnerCreateAnnotationForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerCreateAnnotationForm.uxCreateAnno.Click, AddressOf uxCreateAnno_Click
            AddHandler _partnerCreateAnnotationForm.uxOptionsCancel.Click, AddressOf uxCancel_Click
            AddHandler _partnerCreateAnnotationForm.uxHelp.Click, AddressOf uxHelp_Click
            'AddHandler _partnerCreateAnnotationForm.uxProgressBar
        Else
            ' Unsubscribe to partner form events.
            RemoveHandler _partnerCreateAnnotationForm.uxCreateAnno.Click, AddressOf uxCreateAnno_Click
            RemoveHandler _partnerCreateAnnotationForm.uxOptionsCancel.Click, AddressOf uxCancel_Click
            RemoveHandler _partnerCreateAnnotationForm.uxHelp.Click, AddressOf uxHelp_Click
        End If
    End Sub

    ''' <summary>
    ''' Collects all needed ORMAP data for the new annotation and copies it from the scratch workspace into
    ''' the appropriate ORMAP annotation feature class
    ''' </summary>
    ''' <param name="theAnnoFcCollection">The collection of selected annotation feature classes</param>
    ''' <param name="theScratchAnnoWorkspace">The scratch file geodatabase workspace</param>
    ''' <remarks>
    ''' The EditorExtension On_Create event will overlay the anno with the map index and determine the MapNumber. This is incorrect, 
    ''' as the MapNumber is based on overlay of map index and line feature from which the anno is derived. Correct MapNumber is 
    ''' calculated during anno conversion and carried in TextString of Direction anno. the second newFeature.Store() method below 
    ''' assigns the correct MapNumber, and since its already fired (On_Create), it now trips the EditorExtension On_Update event, 
    ''' which won't perform the overlay and calculate the incorrect MapNumber.
    ''' </remarks>
    Private Sub processNewAnnotation(ByVal theAnnoFcCollection As Collection, ByVal theScratchAnnoWorkspace As IFeatureWorkspace)

        'Now process all the new annotation (since the converter creates new SymbolIDs, adds FeatureIDs, and does not place any of the 
        'ORMAP-required pieces such as MapNumber, user, date, etc).

        Dim i As Integer
        For i = 1 To theAnnoFcCollection.Count
            '------------------------------------------
            ' Append data from scratch workspace to correct anno feature class
            '------------------------------------------
            Dim theAnnoFeatureClass As IFeatureClass = CType(theAnnoFcCollection.Item(i), IFeatureClass)
            Dim theAnnoDataset As IDataset = CType(theAnnoFeatureClass, IDataset)
            Dim theScratchAnnoFCName As String = Nothing

            If theAnnoDataset.Category = "SDE Feature Class" Then
                Dim sqlSyntax As ISQLSyntax = CType(theAnnoDataset.Workspace, ISQLSyntax)
                Dim dbName As String = Nothing
                Dim schemaName As String = Nothing
                sqlSyntax.ParseTableName(theAnnoDataset.Name, dbName, schemaName, theScratchAnnoFCName)
            Else
                theScratchAnnoFCName = theAnnoDataset.Name
            End If

            theScratchAnnoWorkspace.OpenFeatureClass(theScratchAnnoFCName)

            EditorExtension.Editor.StartEditing(theAnnoDataset.Workspace)
            EditorExtension.Editor.StartOperation()

            Dim theQueryFilter As IQueryFilter2 = New QueryFilter
            Dim theScratchAnnoFC As IFeatureClass = theScratchAnnoWorkspace.OpenFeatureClass(theScratchAnnoFCName)
            Dim theScratchCursor As IFeatureCursor = CType(theScratchAnnoFC.Search(theQueryFilter, False), IFeatureCursor)
            Dim thisScratchFeature As IFeature
            Dim newFeature As IFeature
            'Dim numberRows As Integer = -1
            Dim theMapNumber As String = Nothing
            PartnerCreateAnnotationForm.uxProgressBarText.Text = "Creating new annotation in " + theAnnoDataset.Name

            thisScratchFeature = theScratchCursor.NextFeature()
            Do While Not thisScratchFeature Is Nothing
                'PartnerCreateAnnotationForm.uxProgressBar.PerformStep()
                newFeature = theAnnoFeatureClass.CreateFeature()
                Dim scratchAnnoFeature As IAnnotationFeature = DirectCast(thisScratchFeature, IAnnotationFeature)
                Dim newAnnoFeature As IAnnotationFeature = DirectCast(newFeature, IAnnotationFeature)
                Dim scratchSCE As ISymbolCollectionElement = CType(scratchAnnoFeature.Annotation, ISymbolCollectionElement)
                Dim newSCE As ISymbolCollectionElement = New TextElementClass
                '------------------------------------------
                ' Set up field indexes
                '------------------------------------------
                Dim theFeatureIdIndex As Integer = theAnnoFeatureClass.FindField("FeatureID")
                Dim theAutoMethodIndex As Integer = theAnnoFeatureClass.FindField("AutoMethod")
                Dim theAnnoClassIdIndex As Integer = theAnnoFeatureClass.FindField("AnnotationClassID")
                'Dim theSymbolIdIndex As Integer = theAnnoFeatureClass.FindField("SymbolID")
                Dim theMapNumberIndex As Integer = theAnnoFeatureClass.FindField("MapNumber")
                Dim theSymbolName As String = AnnoClassName
                Dim theAnnoClassName As String = AnnoClassName
                Dim theSymbolId As Integer = GetSymbolId(theAnnoFeatureClass, theSymbolName)
                If theSymbolId = -1 Then
                    MessageBox.Show("Missing data: No Symbol defined " & NewLine & _
                        "for " + theAnnoDataset.Name & "." & NewLine & _
                        "Please define this symbol as " + theSymbolName, _
                        "Create Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Else
                    Dim ace As IAnnotationClassExtension = CType(newFeature.Class.Extension, IAnnotationClassExtension)
                    ace.SymbolCollection.Reset()
                    Dim thisSymbolIdentifier As ISymbolIdentifier = ace.SymbolCollection.Next()
                    Do Until thisSymbolIdentifier Is Nothing
                        If thisSymbolIdentifier.ID = theSymbolId Then
                            Exit Do
                        End If
                        thisSymbolIdentifier = ace.SymbolCollection.Next()
                    Loop

                    newSCE.SharedSymbol(thisSymbolIdentifier.ID) = thisSymbolIdentifier.Symbol
                    newSCE.AnchorPoint = scratchSCE.AnchorPoint
                    newSCE.CharacterSpacing = scratchSCE.CharacterSpacing
                    newSCE.CharacterWidth = scratchSCE.CharacterWidth
                    newSCE.Geometry = scratchSCE.Geometry
                    newSCE.Leading = scratchSCE.Leading
                    newSCE.TextPath = scratchSCE.TextPath
                    newSCE.WordSpacing = scratchSCE.WordSpacing
                    newSCE.Text = scratchSCE.Text

                    Dim theValue As String = scratchSCE.Text
                    ' MapNumber is tacked on end of Direction anno by converter... this peels it off and 
                    ' puts it into theMapNumber to be assigned to the anno feature later
                    Dim theIndex As Integer = theValue.IndexOf("%", System.StringComparison.CurrentCulture)
                    Dim theDirection As String = Nothing
                    If theIndex > 0 Then
                        theMapNumber = Right(theValue, Len(theValue) - theIndex - 3)
                        theDirection = Left(theValue, theIndex)
                        newSCE.Text = theDirection
                    End If

                    newAnnoFeature.Annotation = DirectCast(newSCE, IElement)

                    newFeature.Value(theFeatureIdIndex) = System.DBNull.Value
                    newFeature.Value(theAutoMethodIndex) = "CON"
                    newFeature.Value(theAnnoClassIdIndex) = GetSubtypeCode(theAnnoFeatureClass, theAnnoClassName)

                    ' On_Create event will overlay the anno with the map index and determine the MapNumber. This is in-
                    ' correct, as the MapNumber is based on overlay of map index and line feature from which the anno
                    ' is derived. Correct MapNumber is calculated during anno converstion and carried in TextString of 
                    ' Direction anno). the second newFeature.Store() method below assigns the correct MapNumber, and 
                    ' since its already fired (On_Create), it now trips the On_Update event, which won't update MapNumber
                    newFeature.Store()
                    newFeature.Value(theMapNumberIndex) = theMapNumber
                    newFeature.Store()

                    PartnerCreateAnnotationForm.uxProgressBar.PerformStep()
                End If
                thisScratchFeature = theScratchCursor.NextFeature()

                PartnerCreateAnnotationForm.uxProgressBar.PerformStep()

            Loop

            EditorExtension.Editor.StopOperation("Create new annotation")
            EditorExtension.Editor.StopEditing(True)

            PartnerCreateAnnotationForm.uxProgressBar.PerformStep()
        Next i
        PartnerCreateAnnotationForm.uxProgressBar.Value = 0
        PartnerCreateAnnotationForm.uxProgressBarText.Text = Nothing
        theAnnoFcCollection = Nothing
    End Sub

    ''' <summary>
    ''' Clears all selected features in the current map document
    ''' </summary>
    ''' <param name="ThisDocument">The active map document.</param>
    ''' <remarks>There is probably a way to do this using ArcObjects, but gave up trying to find it.</remarks>
    Private Shared Sub clearall(ByVal ThisDocument As IMxDocument)
        ' clears all selections in the focused map
        Dim pMxDoc As IMxDocument
        pMxDoc = ThisDocument
        Dim pFSel As IFeatureSelection, pCompLayer As ICompositeLayer
        Dim l As Long
        For l = 0 To pMxDoc.FocusMap.LayerCount - 1
            If TypeOf pMxDoc.FocusMap.Layer(CInt(l)) Is IFeatureSelection Then
                pFSel = CType(pMxDoc.FocusMap.Layer(CInt(l)), IFeatureSelection)
                pFSel.Clear()
            ElseIf TypeOf pMxDoc.FocusMap.Layer(CInt(l)) Is IGroupLayer Then
                pCompLayer = CType(pMxDoc.FocusMap.Layer(CInt(l)), ICompositeLayer)
                Dim k As Long
                For k = 0 To pCompLayer.Count - 1
                    If TypeOf pCompLayer.Layer(CInt(k)) Is IFeatureSelection Then
                        pFSel = CType(pCompLayer.Layer(CInt(k)), IFeatureSelection)
                        pFSel.Clear()
                    End If
                Next k
            End If
        Next l
    End Sub

    ''' <summary>
    ''' Implements the IConvertLabelsToAnnotation interface to create new bearing and distance 
    ''' annotation from the selected line(s).
    ''' </summary>
    ''' <param name="theMap">The currently selected map.</param>
    ''' <param name="theGeoFeatureLayer">The geo feature layer of the selected source line feature.</param>
    ''' <param name="theAnnoFeatureClass">The appropriate annotation feature class based on scale.</param>
    ''' <param name="isFeatureLinked">Boolean value for defining whether annotation is feature linked (not supported in this version).</param>
    ''' <param name="theScratchAnnoWorkspace">The file geodatabase scratch workspace into which new annotation will be placed.</param>
    ''' <remarks></remarks>
    Private Sub convertLabelsToAnnotation(ByVal theMap As IMap, ByVal theGeoFeatureLayer As IGeoFeatureLayer, ByVal theAnnoFeatureClass As IFeatureClass, _
                                          ByVal isFeatureLinked As Boolean, ByVal theScratchAnnoWorkspace As IFeatureWorkspace)
        Dim theConvertLabelsToAnnotation As IConvertLabelsToAnnotation = New ConvertLabelsToAnnotationClass()
        Dim theTrackCancel As ITrackCancel = New CancelTrackerClass()
        Dim annoFCName As String = Nothing

        Try
            '------------------------------------------
            ' Initialize the converter
            '------------------------------------------
            ' Set the map, the annotation storage type, which features to label, generation of unplaced anno to 'True',
            ' assign the cancel tracker, and do not assign an error event handler (TODO: (RG) Look into how to set this up...
            theConvertLabelsToAnnotation.Initialize(theMap, esriAnnotationStorageType.esriDatabaseAnnotation, _
                                                  esriLabelWhichFeatures.esriSelectedFeatures, True, theTrackCancel, Nothing)

            Dim theAnnoDataset As IDataset = DirectCast(theAnnoFeatureClass, IDataset)
            If Not theGeoFeatureLayer Is Nothing Then
                If theGeoFeatureLayer.DataSourceType = "SDE Feature Class" Then
                    'Dim theAnnoDataset As IDataset = DirectCast(theAnnoFeatureClass, IDataset)
                    Dim sqlSyntax As ISQLSyntax = CType(theAnnoDataset.Workspace, ISQLSyntax)
                    Dim dbName As String = Nothing
                    Dim schemaName As String = Nothing

                    sqlSyntax.ParseTableName(theAnnoDataset.Name, dbName, schemaName, annoFCName)
                Else
                    annoFCName = theAnnoDataset.Name
                End If
                '------------------------------------------
                ' Add the feature to the converter
                '------------------------------------------
                Dim append As Boolean = False

                Dim theScratchWorkspace As IWorkspace2 = DirectCast(theScratchAnnoWorkspace, IWorkspace2)
                If theScratchWorkspace.NameExists(esriDatasetType.esriDTFeatureClass, annoFCName) Then
                    append = True
                Else
                    append = False
                End If

                'Try
                '    Dim theScratchAnno As IFeatureClass = theScratchAnnoWorkspace.OpenFeatureClass(annoFCName)
                '    append = True
                'Catch ex As Exception
                '    append = False
                'End Try

                theConvertLabelsToAnnotation.AddFeatureLayer(theGeoFeatureLayer, _
                                                           annoFCName, _
                                                           theScratchAnnoWorkspace, _
                                                           , _
                                                           isFeatureLinked, append, False, False, False, "")

                '------------------------------------------
                ' Turn on labels, convert, and turn labels off
                '------------------------------------------
                theGeoFeatureLayer.DisplayAnnotation = True
                theConvertLabelsToAnnotation.ConvertLabels()
                theGeoFeatureLayer.DisplayAnnotation = False
            End If

        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try

    End Sub

    ''' <summary>
    ''' Sets up the correct label placements based on type (Distance or Direction) and form settings.
    ''' </summary>
    ''' <param name="theFeatureLayerPropsCollection">The annotation properties collection source line feature.</param>
    ''' <param name="theAnnoFeatureClass">The appropriate annotation feature class (based MapIndex).</param>
    ''' <param name="theAnnoClassName">The annotation class name (should be "34").</param>
    ''' <remarks>
    ''' The map number must be embedded in the label now since annotation can be created in various map indices
    ''' simultaneously. Once annotation is created, there is no way of knowing what its map number is, and this
    ''' must be used later to populate the MapNumber field of the annotation feature class.
    ''' </remarks>
    Private Sub setLabelProperties(ByVal theFeatureLayerPropsCollection As IAnnotateLayerPropertiesCollection2, ByVal theAnnoFeatureClass As IFeatureClass, _
                                   ByVal theAnnoClassName As String, ByVal theMapNumber As String)

        Dim theLabelEngineLayerProperties As ILabelEngineLayerProperties2
        Dim theAnnoLayerProperties As IAnnotateLayerProperties = Nothing
        theLabelEngineLayerProperties = New LabelEngineLayerProperties
        theAnnoLayerProperties = DirectCast(theLabelEngineLayerProperties, IAnnotateLayerProperties)

        Dim theLineLabelPosition As ILineLabelPosition
        theLineLabelPosition = New LineLabelPosition
        Dim theOffset As Integer = 1

        Dim isTop As Boolean = False

        Try
            'Set up the label components and specs
            If String.Compare(theAnnoClassName, "Direction", True, CultureInfo.CurrentCulture) = 0 Then
                '------------------------------------------
                ' Set up labels for "Direction"
                '------------------------------------------
                theAnnoLayerProperties.Class = "Direction"
                theLabelEngineLayerProperties.IsExpressionSimple = False
                '------------------------------------------
                ' Set up the label expression
                '------------------------------------------
                ' TODO:  (RG) This will break when ArcGIS stops using VBA for 
                ' labeling... maybe rewrite as Python?
                '       => Should wait to see what ArcPy.Mapping does... maybe it can be used
                theLabelEngineLayerProperties.Expression = "Function FindLabel ([Direction]) " & vbCrLf & _
                "strTemp = [Direction]" & vbCrLf & _
                "strTemp = Replace(strTemp, ""-"", ""°"", 1, 1)" & vbCrLf & _
                "strTemp = Replace(strTemp, ""-"", ""'"", 1, 1)" & vbCrLf & _
                "strTemp = Left(strTemp, Len(strTemp) - 1) & ""''"" & Right(strTemp, 1)" & vbCrLf & _
                "strTemp = Replace(strTemp, "" "", """")" & vbCrLf & _
                "strDegree = Left(strTemp, InStr(1, strTemp, ""°""))" & vbCrLf & _
                "strMinute = Left(Right(strTemp, Len(strTemp) - Len(strDegree)), InStr(1, Right(strTemp, Len(strTemp) - Len(strDegree)), ""'""))" & vbCrLf & _
                "strSecond = Right(strTemp, Len(strTemp) - Len(strDegree) - Len(strMinute))" & vbCrLf & _
                "strMapNumber = ""%%%" & theMapNumber & """" & vbCrLf & _
                "If Len(strDegree) < 4 Then" & vbCrLf & _
                "strDegree = Left(strDegree, 1) & ""0"" & Right(strDegree, 2)" & vbCrLf & _
                "End If" & vbCrLf & _
                "If Len(strMinute) < 3 Then" & vbCrLf & _
                "strMinute = ""0"" & strMinute" & vbCrLf & _
                "End If" & vbCrLf & _
                "If Len(strSecond) < 5 Then" & vbCrLf & _
                "strSecond = ""0"" & strSecond" & vbCrLf & _
                "End If" & vbCrLf & _
                "FindLabel = strDegree & strMinute & strSecond & strMapNumber" & vbCrLf & _
                "End Function"

                If UpperValue = TopPosition.Direction Then
                    isTop = True
                Else
                    isTop = False
                End If

            ElseIf String.Compare(theAnnoClassName, "Distance", True, CultureInfo.CurrentCulture) = 0 Then
                '------------------------------------------
                ' Set up labels for "Distance"
                '------------------------------------------
                theAnnoLayerProperties.Class = "Distance"
                theLabelEngineLayerProperties.IsExpressionSimple = True
                theLabelEngineLayerProperties.Expression = "FormatNumber([Distance], 2)"

                If UpperValue = TopPosition.Distance Then
                    isTop = True
                Else
                    isTop = False
                End If
            End If
            setAnnoPlacement(isTop)

            With theLineLabelPosition
                .Above = IsAbove
                .Below = IsBelow
                .Offset = theOffset
                .InLine = False
                .AtEnd = False
                .AtStart = False
                .Left = False
                .OnTop = False
                .Right = False
                .Parallel = IsParallel
                .Horizontal = IsHorizontal
                .Perpendicular = IsPerpendicular
                .ProduceCurvedLabels = IsCurved
            End With

            'TODO:  (RG) Look into the use ISymbolCollectionElement here... ESRI warns about not redundantly storing the same symbol
            '       with each feature in the feature class since annotation feature classes are created with a symbol collection 
            '       and the TextElements of annotation features can reference symbols in this collection. Can't really say how 
            '       this works with the converter, however, since it is the converter that is populating the anno feature class.

            '------------------------------------------
            ' Assign overposter properties to label engine
            '------------------------------------------
            Dim theSymbolCollection As ISymbolCollection2 = New SymbolCollectionClass()
            Dim theAnnoClass As IAnnoClass
            theAnnoClass = DirectCast(theAnnoFeatureClass.Extension, IAnnoClass)
            theSymbolCollection = DirectCast(theAnnoClass.SymbolCollection, ISymbolCollection2)

            theLabelEngineLayerProperties.Symbol = DirectCast(theSymbolCollection.Symbol(GetSymbolId(theAnnoFeatureClass, AnnoClassName)), ITextSymbol)

            Dim theBasicOverposterLayerProps As IBasicOverposterLayerProperties
            theBasicOverposterLayerProps = New BasicOverposterLayerProperties
            theBasicOverposterLayerProps.NumLabelsOption = esriBasicNumLabelsOption.esriOneLabelPerShape
            theBasicOverposterLayerProps.LineLabelPosition = theLineLabelPosition
            theBasicOverposterLayerProps.LineOffset = getLineOffset(theLabelEngineLayerProperties.Symbol.Size, isTop)

            If IsWideLine Then
                theBasicOverposterLayerProps.LineOffset = theBasicOverposterLayerProps.LineOffset + AnnotationUtilities.WideLine
            End If

            theLabelEngineLayerProperties.BasicOverposterLayerProperties = theBasicOverposterLayerProps

            Dim theOverposterLayerProperties As IOverposterLayerProperties2
            theOverposterLayerProperties = DirectCast(theLabelEngineLayerProperties.OverposterLayerProperties, IOverposterLayerProperties2)
            theOverposterLayerProperties.TagUnplaced = False
            theLabelEngineLayerProperties.OverposterLayerProperties = DirectCast(theOverposterLayerProperties, IOverposterLayerProperties)

            theFeatureLayerPropsCollection.Add(DirectCast(theLabelEngineLayerProperties, IAnnotateLayerProperties))
        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Uses form property settings to determine how annotation should be placed relative to the source line
    ''' from which it is created.
    ''' </summary>
    ''' <param name="thisSize">The font size of the associated annotation feature class.</param>
    ''' <param name="isTop">The form property defining which element should be place "on top" (distance or direction).</param>
    ''' <returns>The distance used to offset the annotation pairs from the source line feature.</returns>
    ''' <remarks>The fontsize is roughly calculated by MapScale / 240.</remarks>
    Private Function getLineOffset(ByVal thisSize As Double, ByVal isTop As Boolean) As Double
        '------------------------------------------
        ' Calculate the line offset distances
        '------------------------------------------
        ' These formulas were based on desired annotation placement results for Polk County. 
        ' thisSize is the font size of the selected annotation's annotation feature class, and 
        ' is assumed to equal MapScale / 240
        Dim theLineOffset As Double
        If isTop Then
            If IsStandardAbove Then
                If IsBothSides Then
                    theLineOffset = thisSize / 2
                ElseIf IsBothAbove Then
                    theLineOffset = 4 * thisSize
                ElseIf IsBothBelow Then
                    theLineOffset = thisSize
                End If
            ElseIf IsDoubleAbove Then
                If IsBothSides Then
                    theLineOffset = 2 * thisSize
                ElseIf IsBothAbove Then
                    theLineOffset = 2 * (3.5 * thisSize) + thisSize
                ElseIf IsBothBelow Then
                    theLineOffset = 3 * thisSize
                End If
            End If
        ElseIf Not isTop Then
            If IsStandardBelow Then
                If IsBothSides Then
                    theLineOffset = (thisSize / 2) + thisSize
                ElseIf IsBothAbove Then
                    theLineOffset = thisSize / 2
                ElseIf IsBothBelow Then
                    theLineOffset = 5 * thisSize
                End If
            ElseIf IsDoubleBelow Then
                If IsBothSides Then
                    theLineOffset = 3 * thisSize
                ElseIf IsBothAbove Then
                    theLineOffset = 4 * thisSize
                ElseIf IsBothBelow Then
                    theLineOffset = 7.5 * thisSize
                End If
            End If
        End If
        Return theLineOffset
    End Function

    ''' <summary>
    ''' Determines annotation placement based on form settings.
    ''' </summary>
    ''' <param name="isTop">Form property defining which element is "on top".</param>
    ''' <remarks>Sets whether Direction or Distance is place "above" the line.</remarks>
    Private Sub setAnnoPlacement(ByVal isTop As Boolean)
        '------------------------------------------
        ' Set anno placement based on user form settings
        '------------------------------------------
        ' Set class properties according to how user wants to place annotation
        Select Case isTop
            Case True
                If IsBothSides Or IsBothAbove Then
                    IsAbove = True
                    IsBelow = False
                ElseIf IsBothBelow Then
                    IsAbove = False
                    IsBelow = True
                End If
            Case False
                If IsBothSides Or IsBothBelow Then
                    IsAbove = False
                    IsBelow = True
                ElseIf IsBothAbove Then
                    IsAbove = True
                    IsBelow = False
                End If
        End Select
    End Sub

    ' ''' <summary>
    ' ''' Obsolete. Cleans "Direction" and "Distance" subtypes, annotation classes, and symbols from the associated
    ' ''' annotation feature classes. These are placed by the label-to-annotation engine.
    ' ''' </summary>
    ' ''' <param name="theAnnoFC">The name of the annotation feature class</param>
    ' ''' <remarks>Obsolete. Current version uses scratch workspace which is discarded</remarks>
    'Private Sub cleanNewAnnotation(ByVal theAnnoFC As IFeatureClass)
    '    '------------------------------------------
    '    ' Cleans up the mess left by the converter
    '    '------------------------------------------
    '    ' The converter chucks all kinds of "Direction" or "Distance" subtypes, annotation classes, 
    '    ' and symbols into the annotation feature class(es). They must all be removed AFTER 
    '    ' processNewAnnotation has reassigned them to the correct annotation class and symbol. 

    '    Dim theSubtypes As ISubtypes = DirectCast(theAnnoFC, ISubtypes)
    '    Dim theEnumSubtype As IEnumSubtype = theSubtypes.Subtypes
    '    theEnumSubtype.Reset()
    '    Dim subtypeName As String
    '    Dim subtypeCode As Integer

    '    subtypeName = theEnumSubtype.Next(subtypeCode)

    '    Dim theAnnoClassExtenstion As IAnnotationClassExtension = DirectCast(theAnnoFC.Extension, IAnnotationClassExtension)
    '    Dim theSymbolCollection As ISymbolCollection2 = DirectCast(theAnnoClassExtenstion.SymbolCollection, ISymbolCollection2)
    '    Dim theAnnoLayerPropCollection As IAnnotateLayerPropertiesCollection2 = DirectCast(theAnnoClassExtenstion.AnnoProperties, IAnnotateLayerPropertiesCollection2)
    '    Dim theAnnoClassAdmin As IAnnoClassAdmin = DirectCast(theAnnoFC.Extension, IAnnoClassAdmin)

    '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '    ' --DESIGN COMMENT--
    '    ' If things go wrong with the CreateAnnotation process, a large number of "Direction" and "Distance" annotation classes 
    '    ' can be left behind in the annotation feature class. The following code was structured to ensure these remanants are
    '    ' removed from the annotation feature class. This is especially true for AnnotationClassIDs and symbols, which may
    '    ' be left behind even when the associated subtypes are removed. The last step of the cleanup process was designed
    '    ' to remove any remnant "Direction" or "Distance" annotation classes and symbols left from earlier aborted anno
    '    ' creation efforts. 
    '    ' NOTE=>    A clean annotation feature class is assumed prior to the CreateAnnotation process being implemented. 
    '    '           The following cleanup processes have been verified against aborted CreateAnnotation efforts in file,
    '    '           personal, and SDE databases, so they should handle the types of remant data involved with a crashing
    '    '           CreateAnnotation effort. There is no way to know, however, if they can handle cleaning up annotation
    '    '           feature classes that have other types of non-CreateAnnotion related issues.
    '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '    '------------------------------------------
    '    ' Now clean up the mess left by the converter
    '    '------------------------------------------
    '    ' Delete "Direction" and "Distance" subtypes
    '    ' Delete "Direction" and "Distance" annotation classes
    '    ' Delete "Direction" and "Distance" symbols
    '    While Not subtypeName Is Nothing
    '        'Remove Distance and Direction subtypes
    '        If String.Compare(subtypeName, 0, "Direction", 0, 9, CultureInfo.CurrentCulture, CompareOptions.IgnoreCase) = 0 Or _
    '           String.Compare(subtypeName, 0, "Distance", 0, 8, CultureInfo.CurrentCulture, CompareOptions.IgnoreCase) = 0 Then
    '            Try
    '                'Delete the subtype, annotation class, and symbol for "Direction" or "Distance" 
    '                theSubtypes.DeleteSubtype(subtypeCode)
    '                theAnnoLayerPropCollection.Remove(subtypeCode)
    '                theSymbolCollection.Remove(GetSymbolId(theAnnoFC, subtypeName))
    '                'Update the anno FCs property collection to reflect all the subtype removals (need an updated count below)... 
    '                theAnnoClassAdmin.UpdateProperties()
    '            Catch ex As Exception
    '                EditorExtension.ProcessUnhandledException(ex)
    '            End Try
    '        End If
    '        subtypeName = theEnumSubtype.Next(subtypeCode)
    '    End While


    'End Sub

    ' ''' <summary>
    ' ''' Obsolete.
    ' ''' </summary>
    'Private Shared Sub getMapIndexAndScale(ByVal obj As IObject, ByVal theFeature As IFeature, ByVal theSearchGeometry As IGeometry)

    '    Dim theMapScale As String
    '    Dim theMapNumber As String

    '    ' Get the Map Index map number field index.
    '    Dim theMapNumFieldIndex As Integer
    '    theMapNumFieldIndex = theFeature.Fields.FindField(EditorExtension.MapIndexSettings.MapNumberField)

    '    If theSearchGeometry Is Nothing Then Exit Sub
    '    If theSearchGeometry.IsEmpty Then Exit Sub

    '    If theMapNumFieldIndex > NotFoundIndex AndAlso Not IsMapIndex(obj) Then
    '        ' Get the Map Index map number for the location of the new feature.
    '        theMapNumber = GetValue(theSearchGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapNumberField, EditorExtension.MapIndexSettings.MapNumberField)
    '        ' Set the feature map number.
    '        If Len(theMapNumber) > 0 Then
    '            theFeature.Value(theMapNumFieldIndex) = theMapNumber
    '        Else
    '            theFeature.Value(theMapNumFieldIndex) = System.DBNull.Value
    '        End If
    '    End If

    '    ' Set the map scale (if the field exists) to the Map Index map scale for the feature location:

    '    ' Get the Map Index map scale field index.
    '    Dim theMapScaleFieldIndex As Integer
    '    theMapScaleFieldIndex = theFeature.Fields.FindField(EditorExtension.MapIndexSettings.MapScaleField)
    '    If theMapScaleFieldIndex > NotFoundIndex AndAlso Not IsMapIndex(obj) Then
    '        ' Get the Map Index map scale for the location of the new feature.
    '        theMapScale = GetValue(theSearchGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapScaleField, EditorExtension.MapIndexSettings.MapNumberField)
    '        ' Set the feature map scale.
    '        If Len(theMapScale) > 0 Then
    '            theFeature.Value(theMapScaleFieldIndex) = theMapScale
    '        Else
    '            theFeature.Value(theMapScaleFieldIndex) = System.DBNull.Value
    '        End If
    '    End If
    'End Sub

#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"

    ''' <summary>
    ''' Called by ArcMap once per second to check if the command is enabled.
    ''' </summary>
    ''' <returns>Boolean collective determining if tool can be enabled</returns>
    ''' <remarks>WARNING: Do not put computation-intensive code here.</remarks>
    Public ReadOnly Property Enabled() As Boolean
        Get
            Dim canEnable As Boolean
            canEnable = EditorExtension.CanEnableExtendedEditing
            canEnable = canEnable AndAlso EditorExtension.Editor.EditState = esriEditState.esriStateEditing
            canEnable = canEnable AndAlso EditorExtension.IsValidWorkspace
            'Return the opposite of canEnable since label to anno conversion engine only works OUTSIDE an edit session
            Return Not canEnable
        End Get
    End Property

#End Region

#End Region

#Region "Implemented Interface Properties"

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

End Class



