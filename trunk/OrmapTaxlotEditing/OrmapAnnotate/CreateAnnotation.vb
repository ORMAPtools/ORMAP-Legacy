#Region "Copyright 2008 ORMAP Tech Group"

' File:  TaxlotAssignment.vb
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
'Imports com.esri.arcgis.support.ms.stdole
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Environment
Imports System.Globalization
Imports System.Drawing.Text
Imports stdole
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
Imports OrmapTaxlotEditing.DataMonitor
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.EditorExtension
Imports OrmapTaxlotEditing.Utilities
Imports OrmapTaxlotEditing.AnnotationUtilities

#End Region

''' <summary>
''' Provides an ArcMap Tool with functionality to 
''' allow users to create Distance and Direction
''' annotation from Line Feature Classes that contain
''' appropriately formatted Distance and Direction
''' attributes.
''' NOTE=> Annotation Feature Classes must have defi-
''' nitions for an Annotation Class named "34" and a
''' Symbol named "34" (with appropriately defined
''' font size and type information). The tool will 
''' use these definitions for new annotation.
''' </summary>
''' <remarks><seealso cref="CreateAnnotation"/></remarks>
''' 

<ComClass(CreateAnnotation.ClassId, CreateAnnotation.InterfaceId, CreateAnnotation.EventsId), _
 ProgId("OrmapTaxlotEditing.CreateAnnotation")> _
Public NotInheritable Class CreateAnnotation
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

        MyBase.m_category = "OrmapAnnotate"  'localizable text 
        MyBase.m_caption = "CreateAnnotation"   'localizable text 
        MyBase.m_message = "Creates Distance and Direction annotation with user-selected placement preferences"   'localizable text 
        MyBase.m_toolTip = "Create Distance && Direction annotation" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_CreateAnnotation"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try


    End Sub
#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication
    Private _bitmapResourceName As String

#End Region

#Region "Properties"
    Private WithEvents _partnerCreateAnnotationForm As CreateAnnotationForm

    Friend ReadOnly Property PartnerCreateAnnotationForm() As CreateAnnotationForm
        Get
            If _partnerCreateAnnotationForm Is Nothing OrElse _partnerCreateAnnotationForm.IsDisposed Then
                setPartnerAnnotationOptionsForm(New CreateAnnotationForm())
            End If
            Return _partnerCreateAnnotationForm
        End Get
    End Property



    Private _annoClassName As String
    Public ReadOnly Property AnnoClassName() As String
        Get
            _annoClassName = "34"
            Return _annoClassName
        End Get
    End Property

    Private _isCurved As Boolean
    Public ReadOnly Property IsCurved() As Boolean
        Get
            _isCurved = PartnerCreateAnnotationForm.uxCurved.Checked
            Return _isCurved
        End Get
    End Property

    Private _isParallel As Boolean
    Public ReadOnly Property IsParallel() As Boolean
        Get
            _isParallel = PartnerCreateAnnotationForm.uxParallel.Checked
            Return _isParallel
        End Get
    End Property

    Private _isHorizontal As Boolean
    Public ReadOnly Property IsHorizontal() As Boolean
        Get
            _isHorizontal = PartnerCreateAnnotationForm.uxHorizontal.Checked
            Return _isHorizontal
        End Get
    End Property

    Private _isPerpendicular As Boolean
    Public ReadOnly Property IsPerpendicular() As Boolean
        Get
            _isPerpendicular = PartnerCreateAnnotationForm.uxPerpendicular.Checked
            Return _isPerpendicular
        End Get
    End Property

    Private _isAbove As Boolean
    Public Property IsAbove() As Boolean
        Get
            Return _isAbove
        End Get
        Set(ByVal value As Boolean)
            _isAbove = value
        End Set
    End Property

    Private _isBelow As Boolean
    Public Property IsBelow() As Boolean
        Get
            Return _isBelow
        End Get
        Set(ByVal value As Boolean)
            _isBelow = value
        End Set
    End Property

    Private _isBothSides As Boolean
    Public ReadOnly Property IsBothSides() As Boolean
        Get
            _isBothSides = PartnerCreateAnnotationForm.uxBothSides.Checked
            Return _isBothSides
        End Get
    End Property

    Private _isBothAbove As Boolean
    Public ReadOnly Property IsBothAbove() As Boolean
        Get
            _isBothAbove = PartnerCreateAnnotationForm.uxBothAbove.Checked
            Return _isBothAbove
        End Get
    End Property

    Private _isBothBelow As Boolean
    Public ReadOnly Property IsBothBelow() As Boolean
        Get
            _isBothBelow = PartnerCreateAnnotationForm.uxBothBelow.Checked
            Return _isBothBelow
        End Get
    End Property

    Private _isStandardAbove As Boolean
    Public ReadOnly Property IsStandardAbove() As Boolean
        Get
            _isStandardAbove = PartnerCreateAnnotationForm.uxStandardAbove.Checked
            Return _isStandardAbove
        End Get
    End Property

    Private _isDoubleAbove As Boolean
    Public ReadOnly Property IsDoubleAbove() As Boolean
        Get
            _isDoubleAbove = PartnerCreateAnnotationForm.uxDoubleAbove.Checked
            Return _isDoubleAbove
        End Get
    End Property

    Private _isStandardBelow As Boolean
    Public ReadOnly Property IsStandardBelow() As Boolean
        Get
            _isStandardBelow = PartnerCreateAnnotationForm.uxStandardBelow.Checked
            Return _isStandardBelow
        End Get
    End Property

    Private _isDoubleBelow As Boolean
    Public ReadOnly Property IsDoubleBelow() As Boolean
        Get
            _isDoubleBelow = PartnerCreateAnnotationForm.uxDoubleBelow.Checked
            Return _isDoubleBelow
        End Get
    End Property

    Private _isStandardLine As Boolean
    Public ReadOnly Property IsStandardLine() As Boolean
        Get
            _isStandardLine = PartnerCreateAnnotationForm.uxStandardLine.Checked
            Return _isStandardLine
        End Get
    End Property

    Private _bigLine As Boolean
    Public ReadOnly Property IsBigLine() As Boolean
        Get
            _bigLine = PartnerCreateAnnotationForm.uxBigLine.Checked
            Return _bigLine
        End Get
    End Property

    Private _isDirection As Boolean
    Public ReadOnly Property IsDirection() As Boolean
        Get
            _isDirection = PartnerCreateAnnotationForm.uxDirection.Checked
            Return _isDirection
        End Get
    End Property

    Private _isDistance As Boolean
    Public ReadOnly Property IsDistance() As Boolean
        Get
            _isDistance = PartnerCreateAnnotationForm.uxDistance.Checked
            Return _isDistance
        End Get
    End Property

    Private _referenceScale As Integer
    Public ReadOnly Property ReferenceScale() As Integer
        Get
            _referenceScale = CInt(PartnerCreateAnnotationForm.uxReferenceScale.Text)
            Return _referenceScale
        End Get
    End Property

#End Region

#Region "Event Handlers"

    Private Sub uxCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerCreateAnnotationForm.Cancel.Click
        PartnerCreateAnnotationForm.Close()
    End Sub


    Private Sub setPartnerAnnotationOptionsForm(ByVal value As CreateAnnotationForm)
        If value IsNot Nothing Then
            _partnerCreateAnnotationForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerCreateAnnotationForm.uxCreateAnno.Click, AddressOf uxCreateAnno_Click
            AddHandler _partnerCreateAnnotationForm.uxOptionsCancel.Click, AddressOf uxCancel_Click
        Else
            ' Unsubscribe to partner form events.
            RemoveHandler _partnerCreateAnnotationForm.uxCreateAnno.Click, AddressOf uxCreateAnno_Click
            RemoveHandler _partnerCreateAnnotationForm.uxOptionsCancel.Click, AddressOf uxCancel_Click
        End If
    End Sub
    Friend Sub DoButtonOperation()

        Try
            DataMonitor.CheckValidMapIndexDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & NewLine & _
                                "Please load this dataset into your map.", _
                                "Create Annotation", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            End If

            PartnerCreateAnnotationForm.ShowDialog()

        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)

        End Try

    End Sub
    Private Sub uxCreateAnno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim theLayer As IFeatureLayer
        Dim theMxDoc As IMxDocument
        Dim theMap As IMap
        Dim theSelectedFeaturesCursor As IFeatureCursor
        Dim theSelectedFeature As IFeature
        Dim theGeometry As IGeometry
        Dim theAnnoFCName As String

        'TODO:  Put a wait cursor in here... 
        'PartnerCreateAnnotationForm.UseWaitCursor = False

        'TODO:  This layer enumeration needs to be rewritten as a function... it gets used three times here:
        '       Once for the Geo Feature Layers and twice for the Anno Feature Layers
        '       Should be something like getLayerByUid(ByVal theUid as IUID) as IFeatureLayer
        theMxDoc = DirectCast(EditorExtension.Application.Document, IMxDocument)
        theMap = theMxDoc.FocusMap

        'Set the reference (Line offsets are based on 1:1200 scale default)
        theMap.MapScale = ReferenceScale

        Dim theActiveView As IActiveView = CType(theMap, IActiveView)
        Dim theMapExtent As IEnvelope = theActiveView.Extent

        'Enumerate all the feature layers in the map
        'NOTE=> IEnumLayer will NOT return annotation feature classes... they are handled as graphics layers (IFDOGraphicsLayer)
        Dim thisGeoLayersUID As IUID = New UID
        thisGeoLayersUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
        Dim theEnumLayer As IEnumLayer = GetLayerEnum(theMap, thisGeoLayersUID)

        'Loop through the layers
        theLayer = CType(theEnumLayer.Next, IFeatureLayer)
        Dim theAnnoFcCollection As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)

        Do Until (theLayer Is Nothing)
            'Eliminate layers that are turned off or not feature layers 
            If theLayer.Visible AndAlso TypeOf theLayer Is IFeatureLayer Then

                'Now get selected features in this layer. If not, skip to next Layer
                theSelectedFeaturesCursor = SpatialUtilities.GetSelectedFeatures(theLayer)

                If Not theSelectedFeaturesCursor Is Nothing Then
                    Dim theGeoFeatureLayer As IGeoFeatureLayer = CType(theLayer, IGeoFeatureLayer)
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
                    theSelectedFeature = theSelectedFeaturesCursor.NextFeature
                    Dim theFeatureSelection As IFeatureSelection = CType(theLayer, IFeatureSelection)
                    Do Until theSelectedFeature Is Nothing
                        'Now reselect just this feature for the conversion engine
                        theFeatureSelection.Clear()
                        theFeatureSelection.Add(theSelectedFeature)
                        theGeometry = theSelectedFeature.Shape

                        Dim theMapScale As String
                        'TODO:  This returns a different MapScale than is returned in the anno feature class when the EditorExtension.OnCreate 
                        'method is fired... Why????
                        theMapScale = GetValue(theGeometry, MapIndexFeatureLayer.FeatureClass, EditorExtension.MapIndexSettings.MapScaleField, EditorExtension.MapIndexSettings.MapNumberField)

                        'Get annoFC based on MapScale
                        theAnnoFCName = GetAnnoFCName(theMapScale)

                        'Annotation layers are Feature Dataset Objects and the enum needs the UID for IFDOGraphicsLayer
                        'NOTE=> See ArcObjects IMaps.Layers for list of map layer UID

                        Dim theAnnoFeatureClass As IFeatureClass = GetAnnoFeatureClass(theMap, theAnnoFCName)
                        Dim theFeatureLayerPropsCollection As IAnnotateLayerPropertiesCollection2
                        theFeatureLayerPropsCollection = CType(theGeoFeatureLayer.AnnotationProperties, IAnnotateLayerPropertiesCollection2)
                        theFeatureLayerPropsCollection.Clear()

                        setLabelProperties(theFeatureLayerPropsCollection, theAnnoFeatureClass, "Direction")
                        setLabelProperties(theFeatureLayerPropsCollection, theAnnoFeatureClass, "Distance")
                        convertLabelsToGDBAnnotation(theMap, theGeoFeatureLayer, theAnnoFeatureClass, False, theLayer.FeatureClass)
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
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If theAnnoFcCollection.Count = 0 Or Not theAnnoFcCollection.ContainsKey(theAnnoFCName) Then
                            theAnnoFcCollection.Add(theAnnoFCName, GetMaxOidByAnnoFC(theAnnoFeatureClass) - 1)
                        End If

                        theSelectedFeature = theSelectedFeaturesCursor.NextFeature
                    Loop

                End If
            End If
            theLayer = CType(theEnumLayer.Next, IFeatureLayer)
        Loop

        'TODO:  Make this a separate method (pass in theAnnoFcCollection)
        'Now process all the new annotation (since the converter creates new SymbolIDs, adds FeatureIDs, and does not place any of the 
        'ORMAP-required pieces such as MapNumber, user, date, etc).
        Dim thisFDOLayersUID2 As IUID = New UID
        thisFDOLayersUID2.Value = "{5CEAE408-4C0A-437F-9DB3-054D83919850}"
        Dim theAnnoEnumLayer2 As IEnumLayer = theMap.Layers(CType(thisFDOLayersUID2, UID), True)
        Dim theAnnoFeatureClass2 As IFeatureClass = Nothing
        Dim theAnnoLayer2 As IFeatureLayer
        theAnnoEnumLayer2.Reset()
        theAnnoLayer2 = CType(theAnnoEnumLayer2.Next, IFeatureLayer)

        Dim thisPair As KeyValuePair(Of String, Integer)
        For Each thisPair In theAnnoFcCollection
            'Get the Annotation Feature Class
            Do While Not (theAnnoLayer2 Is Nothing)
                If String.Compare(theAnnoLayer2.Name, thisPair.Key, True, CultureInfo.CurrentCulture) = 0 Then
                    theAnnoFeatureClass2 = theAnnoLayer2.FeatureClass
                    Exit Do
                End If
                theAnnoLayer2 = CType(theAnnoEnumLayer2.Next, IFeatureLayer)
            Loop
            processNewAnnotation(theAnnoFeatureClass2, CInt(thisPair.Value))
            theAnnoEnumLayer2.Reset()
            theAnnoLayer2 = CType(theAnnoEnumLayer2.Next, IFeatureLayer)
        Next thisPair

        theAnnoFcCollection = Nothing
        theActiveView.Refresh()
    End Sub
#End Region

#Region "Methods"
    Private Sub convertLabelsToGDBAnnotation(ByVal theMap As IMap, ByVal theGeoFeatureLayer As IGeoFeatureLayer, ByVal theAnnoFeatureClass As IFeatureClass, ByVal isFeatureLinked As Boolean, ByVal theFeatureClass As IFeatureClass)
        'TODO:  (RG) Remove theFeatureClass after testing... looking at alterations to AnnoClass and Symbols... 
        '       => Doesn't matter if converter is set to annoFC and annoFC workspace or not... still adds
        '          the AnnoClass and new Symbols to the annoFC and NOT the line FC used as source... WHAT???
        Dim theConvertLabelsToAnnotation As IConvertLabelsToAnnotation = New ConvertLabelsToAnnotationClass()
        Dim theTrackCancel As ITrackCancel = New CancelTrackerClass()

        theConvertLabelsToAnnotation.Initialize(theMap, esriAnnotationStorageType.esriDatabaseAnnotation, _
                                              esriLabelWhichFeatures.esriSelectedFeatures, True, theTrackCancel, Nothing)

        If Not theGeoFeatureLayer Is Nothing Then
            Dim theAnnoDataset As IDataset = DirectCast(theAnnoFeatureClass, IDataset)
            Dim theAnnoWorkspace As IFeatureWorkspace = DirectCast(theAnnoDataset.Workspace, IFeatureWorkspace)

            'Add the layer information to the converter object. Specify the parameters of the output annotation feature class here as well
            'TODO: (RG) Need to test if this works for feature linked anno
            theConvertLabelsToAnnotation.AddFeatureLayer(theGeoFeatureLayer, _
                                                       theAnnoFeatureClass.AliasName, _
                                                       theAnnoWorkspace, DirectCast(theAnnoFeatureClass.FeatureDataset, IFeatureDataset), _
                                                       isFeatureLinked, True, False, False, False, "")

            theGeoFeatureLayer.DisplayAnnotation = True
            theConvertLabelsToAnnotation.ConvertLabels()
            theGeoFeatureLayer.DisplayAnnotation = False
        End If
    End Sub

    Private Sub setLabelProperties(ByVal theFeatureLayerPropsCollection As IAnnotateLayerPropertiesCollection2, ByVal theAnnoFeatureClass As IFeatureClass, ByVal theAnnoClassName As String)

        Dim theLabelEngineLayerProperties As ILabelEngineLayerProperties2
        Dim theAnnoLayerProperties As IAnnotateLayerProperties = Nothing
        Dim isTop As Boolean = False

        theLabelEngineLayerProperties = New LabelEngineLayerProperties

        theAnnoLayerProperties = DirectCast(theLabelEngineLayerProperties, IAnnotateLayerProperties)

        Dim theLineLabelPosition As ILineLabelPosition
        theLineLabelPosition = New LineLabelPosition

        'Set up the label components and specs
        If String.Compare(theAnnoClassName, "Direction", True, CultureInfo.CurrentCulture) = 0 Then
            theAnnoLayerProperties.Class = "Direction"
            theLabelEngineLayerProperties.IsExpressionSimple = False
            'TODO:  (RG) This will break if ArcGIS stops using VBA for labeling... maybe rewrite as Python?
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
            "If Len(strDegree) < 4 Then" & vbCrLf & _
            "strDegree = Left(strDegree, 1) & ""0"" & Right(strDegree, 2)" & vbCrLf & _
            "End If" & vbCrLf & _
            "If Len(strMinute) < 3 Then" & vbCrLf & _
            "strMinute = ""0"" & strMinute" & vbCrLf & _
            "End If" & vbCrLf & _
            "If Len(strSecond) < 5 Then" & vbCrLf & _
            "strSecond = ""0"" & strSecond" & vbCrLf & _
            "End If" & vbCrLf & _
            "FindLabel = strDegree & strMinute & strSecond" & vbCrLf & _
            "End Function"

            'TODO:  (RG) Redundant use of if blocks for Direction and Distance... rewrite
            If IsDirection Then
                isTop = True
                If IsBothSides Or IsBothAbove Then
                    IsAbove = True
                    IsBelow = False
                ElseIf IsBothBelow Then
                    IsAbove = False
                    IsBelow = True
                End If
            ElseIf Not IsDirection Then
                isTop = False
                If IsBothSides Or IsBothBelow Then
                    IsAbove = False
                    IsBelow = True
                ElseIf IsBothAbove Then
                    IsAbove = True
                    IsBelow = False
                End If
            End If

        ElseIf String.Compare(theAnnoClassName, "Distance", True, CultureInfo.CurrentCulture) = 0 Then
            theAnnoLayerProperties.Class = "Distance"
            theLabelEngineLayerProperties.IsExpressionSimple = True
            theLabelEngineLayerProperties.Expression = "FormatNumber([Distance], 2)"

            If IsDistance Then
                isTop = True
                If IsBothSides Or IsBothAbove Then
                    IsAbove = True
                    IsBelow = False
                ElseIf IsBothBelow Then
                    IsAbove = False
                    IsBelow = True
                End If
            ElseIf Not IsDistance Then
                isTop = False
                If IsBothSides Or IsBothBelow Then
                    IsAbove = False
                    IsBelow = True
                ElseIf IsBothAbove Then
                    IsAbove = True
                    IsBelow = False
                End If
            End If
        End If

        With theLineLabelPosition
            .Above = IsAbove
            .Below = IsBelow
            .Offset = 1
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

        Dim theSymbolCollection As ISymbolCollection2 = New SymbolCollectionClass()
        Dim theAnnoClass As IAnnoClass
        theAnnoClass = DirectCast(theAnnoFeatureClass.Extension, IAnnoClass)
        theSymbolCollection = DirectCast(theAnnoClass.SymbolCollection, ISymbolCollection2)

        'TODO: (RG) Assign "34" to variable (or constant)
        theLabelEngineLayerProperties.Symbol = DirectCast(theSymbolCollection.Symbol(GetSymbolId(theAnnoFeatureClass, "34")), ITextSymbol)

        Dim theBasicOverposterLayerProps As IBasicOverposterLayerProperties
        theBasicOverposterLayerProps = New BasicOverposterLayerProperties
        theBasicOverposterLayerProps.NumLabelsOption = esriBasicNumLabelsOption.esriOneLabelPerShape
        theBasicOverposterLayerProps.LineLabelPosition = theLineLabelPosition
        theBasicOverposterLayerProps.LineOffset = getLineOffset(theLabelEngineLayerProperties.Symbol.Size, isTop)

        If IsBigLine Then
            'TODO:  (RG) There should be different size offsets for different scales...
            theBasicOverposterLayerProps.LineOffset = theBasicOverposterLayerProps.LineOffset + 60
        End If

        theLabelEngineLayerProperties.BasicOverposterLayerProperties = theBasicOverposterLayerProps

        Dim theOverposterLayerProperties As IOverposterLayerProperties2
        theOverposterLayerProperties = DirectCast(theLabelEngineLayerProperties.OverposterLayerProperties, IOverposterLayerProperties2)
        theOverposterLayerProperties.TagUnplaced = False
        theLabelEngineLayerProperties.OverposterLayerProperties = DirectCast(theOverposterLayerProperties, IOverposterLayerProperties)

        theFeatureLayerPropsCollection.Add(DirectCast(theLabelEngineLayerProperties, IAnnotateLayerProperties))
    End Sub

    Private Sub processNewAnnotation(ByVal theAnnoFeatureClass As IFeatureClass, ByVal theMinOid As Integer)
        Dim thisOid As Integer
        Dim theMaxOid As Integer
        Dim theAnnoDataset As IDataset = DirectCast(theAnnoFeatureClass, IDataset)
        Dim theAnnoWorkspace As IFeatureWorkspace = DirectCast(theAnnoDataset.Workspace, IFeatureWorkspace)
        Dim theAnnoWorkspaceEditControl As IWorkspaceEditControl = DirectCast(theAnnoWorkspace, IWorkspaceEditControl)

        EditorExtension.Editor.StartEditing(theAnnoDataset.Workspace)

        'Get the max ID from this anno layer (it will be the anno just created by the label to anno converter)
        Dim theQueryDef As IQueryDef
        Dim theRow As IRow
        Dim theIdCursor As ICursor
        theQueryDef = theAnnoWorkspace.CreateQueryDef
        theQueryDef.Tables = theAnnoDataset.Name
        theQueryDef.SubFields = "MAX(OBJECTID)"
        theIdCursor = theQueryDef.Evaluate
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '-- DESIGN COMMENT --
        '   Need to do some bookkeeping on the anno that was just created by the label to anno converter:
        '   1- Delete FeatureID (converter puts it in even when not feature-linking)
        '   2- Set dummpy MapNumber. If left <Null>, EditEvents_OnCreateFeature will not set correct value
        '   3- Set AutoMethod to Calculated
        '   4- Set AnnotationClassID to index for "34" (Distance and Bearing subtype
        '   5- Need to set correct Symbol
        '   6- Need to insert the new row (RowBuffer [fires OnCreate event])
        '   7- Need to delete the old row (Row [fires the OnDelete event])
        '   8- Need to delete the 'Distance' and 'Direction' symbols addded to the anno feature class
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        theRow = theIdCursor.NextRow
        Dim theFeatureIdIndex As Integer = theAnnoFeatureClass.FindField("FeatureID")
        Dim theMapNumberIndex As Integer = theAnnoFeatureClass.FindField("MapNumber")
        Dim theAutoMethodIndex As Integer = theAnnoFeatureClass.FindField("AutoMethod")
        Dim theAnnoClassIdIndex As Integer = theAnnoFeatureClass.FindField("AnnotationClassID")
        Dim theSymbolIdIndex As Integer = theAnnoFeatureClass.FindField("SymbolID")
        Dim theInsertCursor As ICursor
        Dim theSymbolName As String = AnnoClassName
        Dim theAnnoClassName As String = AnnoClassName
        Dim theSymbolId As Integer = GetSymbolId(theAnnoFeatureClass, theSymbolName)

        'Force simple edits to trigger the EditorExtension.EditEvents_OnCreate event handler 
        theAnnoWorkspaceEditControl.SetStoreEventsRequired()

        theMaxOid = CInt(theRow.Value(0))
        Dim theTable As ITable = DirectCast(theAnnoFeatureClass, ITable)
        For thisOid = theMinOid To theMaxOid
            Dim theOldRow As IRow = theTable.GetRow(thisOid)
            theOldRow.Value(theSymbolIdIndex) = theSymbolId
            'TODO:  (RG) System throws exception on InsertRow if line source direction or bearing is <null>... need to handle
            '       => Will acutally crash in SpatialUtilities.SetAnnoSize called from EditorExtension.EdidEvents_OnCreateFeature
            '               theGeometry = theFeature.Shape
            '               If theGeometry.IsEmpty Then
            '       throws exception because theFeature.Shape is null (neither empty nor present)

            'TODO:  (RG) Verify how this works when Autoupdate is on and when it is off

            'This is cleanest way (i.e., the way that actually WORKS) to ensure correct AnnoClassID, SymbolID, and MapNumber are 
            'placed (including by the the OnCreate event in EditorExtension
            Try
                theOldRow.Store()
                Dim theRowBuffer As IRowBuffer = theOldRow
                theRowBuffer.Value(theFeatureIdIndex) = System.DBNull.Value
                'TODO:  (RG) Should insert a DBNull for MapNumber if AutoUpdate is turned off, otherwise insert a dummy value 
                '       (so that EditorExtension_OnCreate will populate it with the correct MapNumber)
                theRowBuffer.Value(theMapNumberIndex) = "1.1.1"
                theRowBuffer.Value(theAutoMethodIndex) = "CON"
                theRowBuffer.Value(theAnnoClassIdIndex) = GetSubtypeCode(theAnnoFeatureClass, theAnnoClassName)
                theInsertCursor = theTable.Insert(True)
                theInsertCursor.InsertRow(theRowBuffer)
                theOldRow.Delete()
            Catch ex As Exception
                MsgBox(ex)
            End Try
        Next thisOid

        EditorExtension.Editor.StopEditing(True)
    End Sub

    Private Function getLineOffset(ByVal thisSize As Double, ByVal isTop As Boolean) As Double
        'These formulas were based on desired annotation placement results for Polk County. 
        'thisSize is the font size of the selected annotation's annotation feature class
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
            canEnable = canEnable AndAlso EditorExtension.Editor.EditState = esriEditState.esriStateEditing
            canEnable = canEnable AndAlso EditorExtension.IsValidWorkspace
            'Return the opposite of canEnable since label to anno conversion engine only works OUTSIDE an edit session
            Return Not canEnable
        End Get
    End Property

#End Region

#Region "Methods"

    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            _application = CType(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If

        ' TODO:  Add other initialization code
    End Sub

    Public Overrides Sub OnClick()
        DoButtonOperation()
    End Sub
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
    Public Const ClassId As String = "409183dd-d706-4258-89f4-4224d9f7077f"
    Public Const InterfaceId As String = "31b8958f-4290-4e34-9544-8108349339b9"
    Public Const EventsId As String = "b2c63309-c472-4079-bd32-9ab091d73216"
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
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

#End Region

End Class



