
#Region "Copyright 2011 ORMAP Tech Group"

' File:  SpiralUtilities.vb
'
' Original Author:  Jonathan McDowell, Clackamas County Technology Services 
'
' Date Created:  September 29, 2011
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

#Region "Imported Namespaces"


Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports System.Xml
Imports System.IO
Imports ESRI.ArcGIS.SystemUI
Imports stdole

#End Region

''' <summary>
''' This module provides helper routines for the construction of spirals.  The routines in this module are used by SCS_Botton.vb adn SpiralConstruction_Button.vb.
''' </summary>
''' <remarks></remarks>
''' 
Module SpiralUtilities
    Dim _editor As IEditor3 = CType(My.ArcMap.Editor, IEditor3)

    ''' <summary>
    ''' Checks the editing State
    ''' </summary>
    ''' <returns>True or False</returns>
    ''' <remarks></remarks>
    Friend Function IsEnable() As Boolean
        Dim IsEditing As Boolean
        If My.ArcMap.Editor.EditState = esriEditState.esriStateNotEditing Then
            IsEditing = False
        Else
            IsEditing = True
        End If
        Return IsEditing
    End Function

    ''' <summary>
    ''' Transforms the display coordinate to a map coordinate.
    ''' </summary>
    ''' <param name="X">Input X value</param>
    ''' <param name="Y">Input Y Value</param>
    ''' <returns>The point as withe map coordinates</returns>
    ''' <remarks></remarks>
    Friend Function getDataFrameCoords(ByVal X As Integer, ByVal Y As Integer) As IPoint
        'Dim displayTransformation As ESRI.ArcGIS.Display.IDisplayTransformation
        'displayTransformation = _app.Display.DisplayTransformation
        Dim theDisplayTransformation As IDisplayTransformation = My.ThisApplication.Display.DisplayTransformation

        Return theDisplayTransformation.ToMapPoint(X, Y)
    End Function
    ''' <summary>
    ''' Gets the closest snapping environment point
    ''' </summary>
    ''' <param name="point"></param>
    ''' <returns>a point</returns>
    ''' <remarks></remarks>
    Function getSnapPoint(ByVal point As IPoint) As IPoint
        Dim snapEnv As ISnapEnvironment = CType(_editor, ISnapEnvironment)
        snapEnv.SnapPoint(point)
        Return point
    End Function
    ''' <summary>
    ''' Constructs to spiral curve sprial transition
    ''' </summary>
    ''' <param name="theFromPoint">As an IPoint.  The beginning point of the spiral-curve-spiral construction</param>
    ''' <param name="theTangentPoint">As an IPoint.  The tangent point, or Point of Intersect of the tangents for the spiral-curve-spiral construction</param>
    ''' <param name="theToPoint">As an IPoint.  The end point of the spiral-curve-spiral transition</param>
    ''' <param name="theSpiralLengths"></param>
    ''' <param name="theRadius">As double, the radius of the central spiral</param>
    ''' <param name="isCCW">as boolena, is curve counter clockwise</param>
    ''' <remarks></remarks>
    Public Sub ConstructSCSbyLength(ByVal theFromPoint As IPoint, ByVal theTangentPoint As IPoint, ByVal theToPoint As IPoint, ByVal theSpiralLengths As Double, ByVal theRadius As Double, ByVal isCCW As Boolean)
        If My.ArcMap.Editor.EditState = esriEditState.esriStateNotEditing Then
            Exit Sub
        End If
        Try
            Dim toCurvature As Double = 1 / theRadius
            Dim DensifyParameter As Double = 0.5

            'Constructs the spiral curves
            Dim theFirstSpiralPolyLine As IPolyline6 = Construct_Spiral_by_length(theFromPoint, theTangentPoint, 0, toCurvature, isCCW, theSpiralLengths)

            If isCCW Then
                isCCW = False
            Else
                isCCW = True
            End If

            Dim theSecondSpiralPolyLine As IPolyline6 = Construct_Spiral_by_length(theToPoint, theTangentPoint, 0, toCurvature, isCCW, theSpiralLengths)

            'Constructs the Central Curve
            Dim TheCentralCurveConstruction As IConstructCircularArc2 = New CircularArc
            TheCentralCurveConstruction.ConstructEndPointsRadius(theFirstSpiralPolyLine.ToPoint, theSecondSpiralPolyLine.ToPoint, isCCW, theRadius, True)
            Dim theCentralCurve As ICurve3 = TryCast(TheCentralCurveConstruction, ICurve3)
            Dim TheCurvePolyline As ISegmentCollection = New PolylineClass()
            TheCurvePolyline.AddSegment(TryCast(TheCentralCurveConstruction, ISegment))

            Dim theFeatureclass As IFeatureClass = CType(My.ArcMap.Editor.Map.Layer(0), IFeatureLayer2).FeatureClass
            Dim theFirstSpiralFeature As IFeature = theFeatureclass.CreateFeature
            Dim theSecondSpiralFeature As IFeature = theFeatureclass.CreateFeature
            Dim theCenterCircularFeature As IFeature = theFeatureclass.CreateFeature

            'Add the new features to the feature Class
            My.ArcMap.Editor.StartOperation()
            theFirstSpiralFeature.Shape = CType(theFirstSpiralPolyLine, IGeometry)
            theFirstSpiralFeature.Store()
            theCenterCircularFeature.Shape = CType(TheCurvePolyline, IGeometry)
            theCenterCircularFeature.Store()
            theSecondSpiralFeature.Shape = CType(theSecondSpiralPolyLine, IGeometry)
            theSecondSpiralFeature.Store()
            My.ArcMap.Editor.StopOperation("Finished Construction")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub
    ''' <summary>
    ''' Creates the circle graphic showing where the cursor is snapping to in regards to getting the point inputs.
    ''' </summary>
    ''' <returns>as graphic marker element</returns>
    ''' <remarks></remarks>
    Public Function Create_Snap_Marker() As IMarkerElement
        Dim TheMarkerElement As IMarkerElement = New MarkerElement
        Dim theMarkerSymbol As ICharacterMarkerSymbol = New CharacterMarkerSymbol
        Dim theSnapFont As stdole.IFontDisp = CType(New stdole.StdFont, stdole.IFontDisp)

        With theSnapFont
            .Name = "ESRI Default Marker"
            .Size = My.ArcMap.Document.SearchTolerancePixels
        End With

        With theMarkerSymbol
            .Font = theSnapFont
            .CharacterIndex = 40
        End With

        TheMarkerElement.Symbol = theMarkerSymbol

        Return TheMarkerElement
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="theFromPoint"></param>
    ''' <param name="theTangentpoint"></param>
    ''' <param name="theFromCurvature"></param>
    ''' <param name="theToCurvature"></param>
    ''' <param name="isCCW"></param>
    ''' <param name="theSpiralLength"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Construct_Spiral_by_length(ByVal theFromPoint As IPoint, ByVal theTangentpoint As IPoint, ByRef theFromCurvature As Double, ByRef theToCurvature As Double, ByVal isCCW As Boolean, ByVal theSpiralLength As Double) As IPolyline6
        Dim thePolyLine As IPolyline6 = CType(New Polyline, IPolyline6)

        Try
            Dim theGeometryEnvironment As IGeometryEnvironment4 = New GeometryEnvironment
            Dim TheSpiralConstruction As IConstructClothoid = CType(theGeometryEnvironment, IConstructClothoid)
            thePolyLine = CType(TheSpiralConstruction.ConstructClothoidByLength(theFromPoint, theTangentpoint, isCCW, theFromCurvature, theToCurvature, theSpiralLength, esriCurveDensifyMethod.esriCurveDensifyByLength, densifyParameter:=0.5), IPolyline6)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Return thePolyLine
    End Function
End Module

