﻿Imports System.Runtime.InteropServices
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

Module SpiralUtilities

    Friend Function getDataFrameCoords(ByVal X As Integer, ByVal Y As Integer) As IPoint
        'Dim displayTransformation As ESRI.ArcGIS.Display.IDisplayTransformation
        'displayTransformation = _app.Display.DisplayTransformation
        Dim theDisplayTransformation As IDisplayTransformation = My.ThisApplication.Display.DisplayTransformation

        Return theDisplayTransformation.ToMapPoint(X, Y)
    End Function
End Module
