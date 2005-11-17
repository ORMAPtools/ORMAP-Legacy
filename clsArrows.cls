VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdArrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CLASS MODULE FOR THE HOOKS/ARROWS COMMAND
'OPENS frmArrows

Implements ICommand
Implements ITool
Private m_pApp As esriFramework.IApplication
Private m_pDoc As esriArcMapUI.IMxDocument
Private m_pMap As IMap
Private m_pExt As IExtensionConfig
Private m_dHookAngle As Double
Private m_bDoOnce As Boolean  'testing
Private m_bInUse As Boolean
Private m_pLineSymbol As ILineSymbol
Private m_pLinePolyline As IPolyline
Private m_pTextSymbol As ITextSymbol
Private m_pStartPoint As IPoint 'hooks
Private m_pTextPoint As IPoint
Private m_pFromBreakPoint As IPoint 'hooks
Private m_pToBreakPoint As IPoint 'hooks
Private m_pArrowPt1 As IPoint
Private m_pArrowPt2 As IPoint
Private m_pArrowPt3 As IPoint
Private m_pArrowPt4 As IPoint
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "C:\active\ModelingWorkshop_01-05-05\CustomCode\ormap\clsArrows.cls"


 
Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler

    
33:     ICommand_Enabled = True


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Checked() As Boolean
  On Error GoTo ErrorHandler

    
45:     ICommand_Checked = False


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler

    
57:     ICommand_Name = "AddArrowsAndHooks"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler


69:     ICommand_Caption = "Add Arrows/Hooks"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler


81:     ICommand_Tooltip = "Open a dialog to add arrows and hooks"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
  On Error GoTo ErrorHandler

    
93:     ICommand_Message = "Initialize the Add Arrows and Hooks form"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpFile() As String
  On Error GoTo ErrorHandler


105:     ICommand_HelpFile = ""


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpContextID() As Long
  On Error GoTo ErrorHandler


117:     ICommand_HelpContextID = ""


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpContextID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here
    ' ICommand_Bitmap =


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Category() As String
  On Error GoTo ErrorHandler


142:     ICommand_Category = "ORMAP"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
  On Error GoTo ErrorHandler

    
154:   Set m_pApp = hook
155:   If g_pApp Is Nothing Then
156:     Set g_pApp = hook
157:   End If
158:   Set m_pDoc = m_pApp.Document
159:   Set m_pMap = m_pDoc.FocusMap

  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

    'Get the MapIndex feature layer and fclass
    Dim pCLFLayer As IFeatureLayer
    Dim pTaxlotFlayer As IFeatureLayer
    Dim pMIFlayer As IFeatureLayer
173:     Set pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
174:     If pTaxlotFlayer Is Nothing Then
175:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        Exit Sub
178:     End If
    
180:     Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
181:     If pMIFlayer Is Nothing Then
182:         MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
        Exit Sub
185:     End If
    
187:     Set pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
188:     If pTaxlotFlayer Is Nothing Then
189:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        Exit Sub
192:     End If
    
194:     Set pCLFLayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
195:     If pCLFLayer Is Nothing Then
196:         MsgBox "Unable to locate Cartographic Lines layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCCartoLines
        Exit Sub
199:     End If
    'Open the form
201:     frmArrows.Show vbModal


  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here
    ' ITool_Cursor =


  Exit Property
ErrorHandler:
  HandleError True, "ITool_Cursor " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 

 
 
 
Private Sub ITool_OnDblClick()
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here


  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnDblClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnKeyDown(ByVal keyCode As Long, ByVal Shift As Long)
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here


  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnKeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnKeyUp(ByVal keyCode As Long, ByVal Shift As Long)
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here


  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnKeyUp " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here


  Exit Function
ErrorHandler:
  HandleError True, "ITool_OnContextMenu " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
 

 
Private Function ITool_Deactivate() As Boolean
  On Error GoTo ErrorHandler

    
  ' stop doing operation
281:   Set m_pTextSymbol = Nothing
282:   Set m_pTextPoint = Nothing
283:   Set m_pLinePolyline = Nothing
284:   Set m_pLineSymbol = Nothing
285:   m_bInUse = False

287:   ITool_Deactivate = True


  Exit Function
ErrorHandler:
  HandleError True, "ITool_Deactivate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler


299:   m_bInUse = True
  Dim pMXDoc As IMxDocument
  Dim pActiveView As IActiveView
302:   Set pMXDoc = m_pApp.Document
303:   Set pActiveView = pMXDoc.FocusMap
  
    'Get reference to MapIndex FC
    Dim sScale As String
    Dim pMIFC As IFeatureClass
    Dim pMIFL As IFeatureLayer
309:     Set pMIFL = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
310:     Set pMIFC = pMIFL.FeatureClass
  
  'If drawing hooks
313:   If frmArrows.lblCurrentTool.Caption = "hook" Then
    'Get point to measure distance from
315:     Set m_pStartPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
316:     m_bDoOnce = False 'testing
317:     Set m_pFromBreakPoint = m_pStartPoint
    'Get the scale of the current mapindex
319:     sScale = GetValueViaOverlay(m_pStartPoint, pMIFC, g_pFldnames.MIMapScaleFN)
320:     sScale = modUtils.ConvertToDescription(pMIFC.Fields, g_pFldnames.MIMapScaleFN, sScale)
  

'If drawing annotation arrows
324:   ElseIf frmArrows.lblCurrentTool.Caption = "arrow" Then
    
326:     If m_pArrowPt1 Is Nothing Then
327:         Set m_pArrowPt1 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
328:     ElseIf m_pArrowPt2 Is Nothing Then
329:         Set m_pArrowPt2 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
330:         DrawArrows
331:     ElseIf m_pArrowPt3 Is Nothing Then
332:         Set m_pArrowPt3 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
333:         DrawArrows
334:     ElseIf m_pArrowPt4 Is Nothing Then
335:         Set m_pArrowPt4 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
336:         DrawArrows
        Dim pArrowLine As IPolyline4
        Dim pArrowPoints As IPointCollection4
339:         Set pArrowPoints = New Polyline
340:         pArrowPoints.AddPoint m_pArrowPt1
341:         pArrowPoints.AddPoint m_pArrowPt2
342:         pArrowPoints.AddPoint m_pArrowPt3
343:         pArrowPoints.AddPoint m_pArrowPt4
344:         Set pArrowLine = pArrowPoints
345:         pArrowLine.Smooth pArrowLine.Length / 10

        Dim pFeature As IFeature
        Dim pWSEdit As IWorkspaceEdit
        Dim pDSet As IDataset
        Dim pArrowLayer As IFeatureLayer
        Dim pArrowFC As IFeatureClass
352:         Set pArrowLayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
353:         If pArrowLayer Is Nothing Then
354:           MsgBox "The layer, " & g_pFldnames.FCCartoLines & ", is not in the map.", vbExclamation, "Layer not found"
355:           m_pEditor.AbortOperation
          Exit Sub
357:         End If
358:         Set pArrowFC = pArrowLayer.FeatureClass
359:         Set pDSet = pArrowFC
360:         Set pWSEdit = pDSet.Workspace
        
        'create the arrow feature
363:         pWSEdit.StartEditOperation
364:         Set pFeature = pArrowFC.CreateFeature
365:         Set pFeature.Shape = pArrowLine
366:         lLineTypeFld = modUtils.LocateFields(pArrowFC, g_pFldnames.CLinesLineTypeFN)
        If lLineTypeFld = -1 Then Exit Sub
368:         pFeature.Value(lLineTypeFld) = 100
        
        'Get the current MapNumber
        Dim sCurMapNum As String
372:         sCurMapNum = GetValueViaOverlay(pFeature.Shape, pMIFC, g_pFldnames.MIMapNumberFN)
        Dim lCLMNfld As Long
374:         lCLMNfld = modUtils.LocateFields(pArrowFC, g_pFldnames.MIMapNumberFN)
375:         pFeature.Value(lCLMNfld) = sCurMapNum
        
377:         pFeature.Store
378:         pWSEdit.StopEditOperation
        
        'Refresh the display
381:         pActiveView.PartialRefresh esriViewBackground, Nothing, pFeature.Extent.Envelope
        
383:         Set pArrowLine = Nothing
384:         Set m_pArrowPt1 = Nothing
385:         Set m_pArrowPt2 = Nothing
386:         Set m_pArrowPt3 = Nothing
387:         Set m_pArrowPt4 = Nothing
        
389:         Set m_pTextSymbol = Nothing
390:         Set m_pTextPoint = Nothing
391:         Set m_pLinePolyline = Nothing
392:         Set m_pLineSymbol = Nothing
        'Deactivate the tool
394:         Set m_pApp.CurrentTool = Nothing
  
  
397:     Else 'Reset everything
398:         Set m_pArrowPt1 = Nothing
399:         Set m_pArrowPt2 = Nothing
400:         Set m_pArrowPt3 = Nothing
401:         Set m_pArrowPt4 = Nothing
        'Deactivate the tool
403:         Set m_pApp.CurrentTool = Nothing
404:     End If
405:   End If
  

  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler

'Draw the temporary line the user sees while moving the mouse
417: If frmArrows.lblCurrentTool.Caption = "hook" Then
  If (Not m_bInUse) Then Exit Sub
  
  Dim pMXDoc As IMxDocument
  Dim pActiveView As IActiveView
422:   Set pMXDoc = m_pApp.Document
423:   Set pActiveView = pMXDoc.FocusMap
  
  Dim bfirstTime As Boolean
426:   If (m_pLineSymbol Is Nothing) Then bfirstTime = True

  'Get current point
  Dim pPoint As IPoint
430:   Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
431:   Set m_pToBreakPoint = pPoint 'the unextended to point used to break the hooks
  
  
   'Draw a virtual line that represents the extended line
   Dim pPline As IPolyline
436:    Set pPline = New Polyline
437:    pPline.FromPoint = m_pStartPoint
438:    pPline.ToPoint = pPoint
   Dim pCv As ICurve
440:    Set pCv = pPline
   
   Dim pCPoint As IConstructPoint
443:    Set pCPoint = New Point
444:    pCPoint.ConstructAlong pCv, esriExtendAtTo, pCv.Length + CDbl(lLineLength), False
445:    Set pPoint = pCPoint
   
447:    If Not m_bDoOnce Then
        Dim pFCPoint As IConstructPoint
449:         Set pFCPoint = New Point
450:         pFCPoint.ConstructAlong pCv, esriExtendAtFrom, -(CDbl(lLineLength)), False
451:         Set m_pStartPoint = pFCPoint
452:         m_bDoOnce = True
453:    End If

  'Draw the line
456:   pActiveView.ScreenDisplay.StartDrawing pActiveView.ScreenDisplay.hdc, -1
    
458:   If bfirstTime Then
    Dim pRGBColor As IRgbColor
    Dim pSymbol As ISymbol
    Dim pFont As IFontDisp
    
    'Line Symbol
464:     Set m_pLineSymbol = New SimpleLineSymbol
465:     m_pLineSymbol.Width = 2
466:     Set pRGBColor = New RgbColor
467:     With pRGBColor
468:       .Red = 223
469:       .Green = 223
470:       .Blue = 223
471:     End With
472:     m_pLineSymbol.Color = pRGBColor
473:     Set pSymbol = m_pLineSymbol
474:     pSymbol.ROP2 = esriROPXOrPen
    
    'Text Symbol
477:     Set m_pTextSymbol = New TextSymbol
478:     m_pTextSymbol.HorizontalAlignment = esriTHACenter
479:     m_pTextSymbol.VerticalAlignment = esriTVACenter
480:     m_pTextSymbol.Size = 16
481:     Set pSymbol = m_pTextSymbol
482:     Set pFont = m_pTextSymbol.Font
483:     pFont.Name = "Arial"
484:     pSymbol.ROP2 = esriROPXOrPen
    
    'Create point to draw text in
487:     Set m_pTextPoint = New Point
    
489:   Else
    'Use existing symbols and draw existing text and polyline
491:     pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
492:     pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
493:     pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
494:     If (m_pLinePolyline.Length > 0) Then _
      pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
496:   End If

  'Get line between from and to points, and angle for text
  Dim pLine As ILine
500:   Set pLine = New esrigeometry.Line
501:   pLine.PutCoords m_pStartPoint, pPoint
  Dim angle As Double
503:   angle = pLine.angle
504:   angle = angle * (180# / 3.14159)
505:   If ((angle > 90#) And (angle < 180#)) Then
506:     angle = angle + 180#
507:   ElseIf ((angle < 0#) And (angle < -90#)) Then
508:     angle = angle - 180#
509:   ElseIf ((angle < -90#) And (angle > -180)) Then
510:     angle = angle - 180#
511:   ElseIf (angle > 180) Then
512:     angle = angle - 180#
513:   End If


  'For drawing text, get text(distance), angle, and point
  Dim deltaX As Double
  Dim deltaY As Double
  Dim distance As Double
520:   deltaX = pPoint.X - m_pStartPoint.X
521:   deltaY = pPoint.Y - m_pStartPoint.Y
522:   m_pTextPoint.X = m_pStartPoint.X + deltaX / 2#
523:   m_pTextPoint.Y = m_pStartPoint.Y + deltaY / 2#
524:   m_pTextSymbol.angle = angle
525:   distance = Round(Sqr((deltaX * deltaX) + (deltaY * deltaY)), 3)
526:   m_pTextSymbol.Text = "" '"[" & distance & "]"
  
  'Draw text
529:   pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
530:   pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
  
  
  'Get polyline with blank space for text
  Dim pPolyline As IPolyline
535:   Set pPolyline = New Polyline
  Dim pSegColl As ISegmentCollection
537:   Set pSegColl = pPolyline
538:   pSegColl.AddSegment pLine
539:   Set m_pLinePolyline = GetSmashedLine(pActiveView.ScreenDisplay, m_pTextSymbol, m_pTextPoint, pPolyline)
  
  'Draw polyline
542:   pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
543:   If (m_pLinePolyline.Length > 0) Then _
    pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
  
546:   pActiveView.ScreenDisplay.FinishDrawing

548: End If

  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseMove " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler

558:   If frmArrows.lblCurrentTool.Caption = "hook" Then
    If (Not m_bInUse) Then Exit Sub
560:     m_bInUse = False
    
    If (m_pLineSymbol Is Nothing) Then Exit Sub
    
    Dim pMXDoc As IMxDocument
    Dim pActiveView As IActiveView
566:     Set pMXDoc = m_pApp.Document
567:     Set pActiveView = pMXDoc.FocusMap
    
    'Draw temp line
570:     pActiveView.ScreenDisplay.StartDrawing pActiveView.ScreenDisplay.hdc, -1
571:     pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
572:     pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
573:     pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
574:     If (m_pLinePolyline.Length > 0) Then pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
575:     pActiveView.ScreenDisplay.FinishDrawing
576:     If frmArrows.lblCurrentTool.Caption = "hook" Then
    'Generate hooks based on the graphic polyline
578:       GenerateHooks m_pLinePolyline
579:     Else
      'Nothing yet
581:     End If
    
583:     Set m_pTextSymbol = Nothing
584:     Set m_pTextPoint = Nothing
585:     Set m_pLinePolyline = Nothing
586:     Set m_pLineSymbol = Nothing
587:   End If


  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseUp " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ITool_Refresh(ByVal hdc As esriSystem.OLE_HANDLE)
  On Error GoTo ErrorHandler



  Exit Sub
ErrorHandler:
  HandleError True, "ITool_Refresh " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function GetSmashedLine(pDisplay As IScreenDisplay, pTextSymbol As ISymbol, pPoint As IPoint, pPolyline As IPolyline) As IPolyline
  On Error GoTo ErrorHandler


  'Returns a graphic Polyline
  Dim pSmashed As IPolyline
  Dim pBoundary As IPolygon
612:   Set pBoundary = New Polygon
613:   pTextSymbol.QueryBoundary pDisplay.hdc, pDisplay.DisplayTransformation, pPoint, pBoundary
  Dim pTopo As ITopologicalOperator
615:   Set pTopo = pBoundary
  
  Dim pIntersect As IPolyline
618:   Set pIntersect = pTopo.Intersect(pPolyline, esriGeometry1Dimension)
619:   Set pTopo = pPolyline
620:   Set GetSmashedLine = pTopo.Difference(pIntersect)


  Exit Function
ErrorHandler:
  HandleError False, "GetSmashedLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub GenerateHooks(pSketch As esrigeometry.IGeometry)
  On Error GoTo ErrorHandler


632: m_dHookAngle = 20
'Make sure the edit sketch is a polyline
If Not TypeOf pSketch Is IPolyline Then Exit Sub

Dim pCurve As ICurve
637: Set pCurve = pSketch

Dim dHookLength As Double
Dim dOverlapLength As Double

'Get the map scale so hook length can be determined
Dim pMIFlayer As IFeatureLayer
Dim pMIFclass As IFeatureClass
645: Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
646: If pMIFlayer Is Nothing Then
647:     MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
    "This process requires a feature class called " & g_pFldnames.FCMapIndex
    Exit Sub
650: End If
651: Set pMIFclass = pMIFlayer.FeatureClass
Dim vMapScale1 As Variant
Dim vMapScale2 As Variant
654: vMapScale1 = modUtils.GetValueViaOverlay(pCurve.FromPoint, pMIFclass, g_pFldnames.MIMapScaleFN)
655: vMapScale2 = modUtils.GetValueViaOverlay(pCurve.ToPoint, pMIFclass, g_pFldnames.MIMapScaleFN)
656: If IsNull(vMapScale1) Or IsNull(vMapScale2) Then
657:     MsgBox "No mapscale for current MapIndex.  Unable to create hooks", vbOKOnly
    Exit Sub
659: End If
660: If vMapScale1 <> vMapScale2 Then
661:     MsgBox "Hook can not span Mapindex polygons with differenc scale", vbCritical
    Exit Sub
663: End If
'ISSUE:  NOT ALL SCALES DEFINED
Dim lLineLength As Long
666: If vMapScale1 = 600 Then
667:     lLineLength = 20
668: ElseIf vMapScale1 = 1200 Then
669:     lLineLength = 20
670: ElseIf vMapScale1 = 2400 Then
671:     lLineLength = 40
672: ElseIf vMapScale1 = 4800 Then
673:     lLineLength = 80
674: ElseIf vMapScale1 = 24000 Then
675:     lLineLength = 400
676: Else
677:     MsgBox "Not a valid mapscale.  Unable to create hooks", vbOKOnly
    Exit Sub
679: End If
680: dHookLength = lLineLength * 0.1
    
'dHookLength = (m_dHookLength * pCurve.Length)

Dim bOK As Boolean
'Make sure the polyline only has two vertices (start and end point only)
686: bOK = IsSketcha2PointLine(pSketch)

688: If bOK = False Then
689:   m_pEditor.AbortOperation
  Exit Sub
691: End If

Dim pApp As esriFramework.IApplication
694: Set pApp = m_pApp 'm_pEditor.Parent
Dim pDoc As esriArcMapUI.IMxDocument
696: Set pDoc = pApp.Document


'Get the hook layer
Dim lLineTypeFld As Long
Dim pWSEdit As IWorkspaceEdit
Dim pDSet As IDataset
Dim pHookLayer As IFeatureLayer
Dim pHookFC As IFeatureClass
705: Set pHookLayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
706: If pHookLayer Is Nothing Then
707:   MsgBox "The layer, " & m_sHookDSName & ", is not in the map.", vbExclamation, "Layer not found"
708:   m_pEditor.AbortOperation
  Exit Sub
710: End If
711: Set pHookFC = pHookLayer.FeatureClass
712: Set pDSet = pHookFC
713: Set pWSEdit = pDSet.Workspace
714: lLineTypeFld = modUtils.LocateFields(pHookFC, g_pFldnames.CLinesLineTypeFN)
If lLineTypeFld = -1 Then Exit Sub
'Create the new hook feature
Dim pFeature As IFeature
Dim pPointColl As esrigeometry.IPointCollection
Dim bCreateDefaultHook As Boolean

721: bCreateDefaultHook = True

723: Set pPointColl = pSketch

Dim dSideA As Double
Dim dSideB As Double
Dim dSideC As Double
Dim pNormal As ILine

'Create a long line that is essentially what the user drew
'Create a hook at each end of the line
'Create hooks based on specified angle and hook length (based on scale)
733: If bCreateDefaultHook = True Then 'Always will be true here
  Dim pNewPointColl As IPointCollection
735:   Set pNewPointColl = New Polyline
    
737:   dSideA = (dHookLength * Sin((360 - m_dHookAngle) * (3.14 / 180)))
738:   dSideC = dHookLength
  
740:   dSideB = (dSideC ^ 2) - (dSideA ^ 2)
741:   dSideB = Sqr(dSideB)
  
743:   Set pNormal = New esrigeometry.Line
744:   pCurve.QueryNormal esriNoExtension, dSideB, False, dSideA, pNormal
  
746:   pNewPointColl.AddPoint pNormal.ToPoint
747:   pNewPointColl.AddPoint pPointColl.Point(0)
748:   pNewPointColl.AddPoint pPointColl.Point(1)
      
750:   dSideA = (dHookLength * Sin(m_dHookAngle * (3.14 / 180)))
751:   dSideC = dHookLength
  
753:   dSideB = (dSideC ^ 2) - (dSideA ^ 2)
754:   dSideB = Sqr(dSideB)
  
756:   pCurve.QueryNormal esriNoExtension, (pCurve.Length - dSideB), False, dSideA, pNormal

758:   pNewPointColl.AddPoint pNormal.ToPoint
  
  'Now get rid of the line between the start and end points (where user clicked)
  Dim pGeomColl As IGeometryCollection
  Dim pPath1 As ISegmentCollection
  Dim pPath2 As ISegmentCollection
  Dim pPath3 As ISegmentCollection
  Dim pSegCollection As ISegmentCollection
  Dim pWholeLine As IPolyline4
  Dim lNewPartIndex As Long
  Dim lNewSegIndex As Long
  Dim bSplitHappened As Boolean
  Dim bBool As Boolean
  Dim pEnumSeg As IEnumSegment
  Dim pSeg As ISegment
  Dim pOutseg As ISegment
  Dim lPartIndex As Long
  Dim lSegIndex As Long
776:   Set pWholeLine = pNewPointColl
777:   pWholeLine.SplitAtPoint m_pFromBreakPoint, True, bBool, bSplitHappened, lNewPartIndex, lNewSegIndex
778:   pWholeLine.SplitAtPoint m_pToBreakPoint, True, bBool, bSplitHappened, lNewPartIndex, lNewSegIndex
  
780:   Set pPath1 = New Path
781:   Set pPath2 = New Path
782:   Set pPath3 = New Path
783:   Set pSegCollection = pWholeLine

785:   Set pEnumSeg = pSegCollection.EnumSegments
786:   pEnumSeg.Next pSeg, lPartIndex, lSegIndex
787:   While Not pSeg Is Nothing
788:     If lSegIndex < 1 Then
789:         pPath1.AddSegment pSeg
790:     ElseIf lSegIndex = 1 Then
791:         pPath2.AddSegment pSeg
792:     ElseIf lSegIndex = 2 Then
793:         pPath3.AddSegment pSeg
794:     End If
795:     pEnumSeg.Next pSeg, lPartIndex, lSegIndex
796:   Wend
797:   Set pGeomColl = New Polyline
798:   pGeomColl.AddGeometry pPath1
799:   pGeomColl.AddGeometry pPath2
800:   pGeomColl.AddGeometry pPath3
801:   pGeomColl.GeometriesChanged
  

  'Store the new hook feature
805:   pWSEdit.StartEditOperation
806:   Set pFeature = pHookFC.CreateFeature
807:   Set pFeature.Shape = pGeomColl 'pWholeLine 'pFinalPolyLine 'pNewPointColl
808:   pFeature.Value(lLineTypeFld) = 101
809:   pFeature.Store
810:   pWSEdit.StopEditOperation
  
  'Refresh the display
  Dim pMXDoc As IMxDocument
  Dim pActiveView As IActiveView
815:   Set pMXDoc = m_pApp.Document
816:   Set pActiveView = pMXDoc.FocusMap
817:   pActiveView.PartialRefresh esriViewBackground, Nothing, pFeature.Extent.Envelope
  
819: End If


  Exit Sub
ErrorHandler:
  HandleError False, "GenerateHooks " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function IsSketcha2PointLine(pGeom As IGeometry) As Boolean
  On Error GoTo ErrorHandler


Dim pPointColl As IPointCollection
  
If Not TypeOf pGeom Is IPointCollection Then Exit Function

835: Set pPointColl = pGeom
836: If pPointColl.PointCount <> 2 Then
837:   MsgBox "When creating the parcel hook only digitize two points", vbExclamation, "Parcel Hook Error"
  Exit Function
839: End If

841: IsSketcha2PointLine = True


  Exit Function
ErrorHandler:
  HandleError True, "IsSketcha2PointLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function ReturnExtended(pExt As esriSegmentExtension, pPolyline As IPolyline, lLength As Long) As IPolyline
  On Error GoTo ErrorHandler


    Dim pCurve As ICurve
    Dim pLine As ILine
    Dim pPline As IPolyline
    Dim pTopoOpt As ITopologicalOperator
    Dim pGeoCol As IGeometryCollection
    ' Need to extend the end (creates an ILine object)
859:     Set pCurve = pPolyline
860:     Set pLine = New esrigeometry.Line
861:     pCurve.QueryTangent pExt, 1, False, _
                        lLength, pLine
    ' Convert ILine to an IPolyline
864:     Set pPline = New Polyline
865:     pPline.FromPoint = pLine.FromPoint
866:     pPline.ToPoint = pLine.ToPoint
867:     Set ReturnExtended = pPline


  Exit Function
ErrorHandler:
  HandleError False, "ReturnExtended " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub DrawArrows()
  On Error GoTo ErrorHandler

    Dim pMXDoc As IMxDocument
    Dim pActiveView As IActiveView
880:     Set pMXDoc = m_pApp.Document
881:     Set pActiveView = pMXDoc.FocusMap
    'Set up line symbol to display temporary line
    Dim pRGBColor As IRgbColor
    Dim pSymbol As ISymbol
    Dim pFont As IFontDisp
    'Line Symbol
887:     Set m_pLineSymbol = New SimpleLineSymbol
888:     m_pLineSymbol.Width = 2
889:     Set pRGBColor = New RgbColor
890:     With pRGBColor
891:       .Red = 223
892:       .Green = 223
893:       .Blue = 223
894:     End With
895:     m_pLineSymbol.Color = pRGBColor
896:     Set pSymbol = m_pLineSymbol
897:     pSymbol.ROP2 = esriROPXOrPen
    
    Dim pArrowLine As IPointCollection4
    Dim pArrowLine2 As IPolyline4
901:     Set pArrowLine = New Polyline
902:     If Not m_pArrowPt1 Is Nothing Then pArrowLine.AddPoint m_pArrowPt1
903:     If Not m_pArrowPt2 Is Nothing Then pArrowLine.AddPoint m_pArrowPt2
904:     If Not m_pArrowPt3 Is Nothing Then pArrowLine.AddPoint m_pArrowPt3
905:     If Not m_pArrowPt4 Is Nothing Then pArrowLine.AddPoint m_pArrowPt4
906:     Set pArrowLine2 = pArrowLine
    'pArrowLine.Smooth pArrowLine.Length / 10
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''Make lines, then add to segment collection
    '''''''''see mouse move event
    ''''''''''''''''''''''''''''''''''''
    
    
    
916:     pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
917:     If (pArrowLine2.Length > 0) Then _
        pActiveView.ScreenDisplay.DrawPolyline pArrowLine2
  
920:     pActiveView.ScreenDisplay.FinishDrawing

    'Set m_pLinePolyline = GetSmashedLine(pActiveView.ScreenDisplay, m_pTextSymbol, m_pTextPoint, pPolyline)

  Exit Sub
ErrorHandler:
  HandleError False, "DrawArrows " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
