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
'
' File name:            cmdArrows
'
' Initial Author:       Type your name here
'
' Date Created:
'
' Description: CLASS MODULE FOR THE HOOKS/ARROWS COMMAND
'
' Entry points:
'       List the public variables and their purposes.
'       List the properties and routines that the module exposes to the rest of the program.
'
' Dependencies: OPENS frmArrows
'       How does this file depend or relate to other files?
'
' Issues:
'       What are unsolved bugs, bottlenecks,
'       possible future enhancements, and
'       descriptions of other issues.
'
' Method:
'       Describe any complex details that make sense on the file level.  This includes explanations
'       of complex algorithms, how different routines within the module interact, and a description
'       of a data structure used in the module.
'
' Updates:
'               None

Option Explicit
'******************************
' Global/Public Definitions
'------------------------------
' Public API Declarations
'------------------------------

'------------------------------
' Public Enums and Constants
'------------------------------

'------------------------------
' Public variables
'------------------------------

'------------------------------
' Public Types
'------------------------------

'------------------------------
' Public loop variables
'------------------------------

'******************************
' Private Definitions
'------------------------------
' Private API declarations
'------------------------------

'------------------------------
' Private Variables
'------------------------------
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
'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "clsArrows.cls"

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------

'interfaces
Implements ICommand
Implements ITool
 
Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler

    
101:     ICommand_Enabled = True


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Checked() As Boolean
  On Error GoTo ErrorHandler

    
113:     ICommand_Checked = False


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler

    
125:     ICommand_Name = "AddArrowsAndHooks"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler


137:     ICommand_Caption = "Add Arrows/Hooks"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler


149:     ICommand_Tooltip = "Open a dialog to add arrows and hooks"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
  On Error GoTo ErrorHandler

    
161:     ICommand_Message = "Initialize the Add Arrows and Hooks form"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpFile() As String
  On Error GoTo ErrorHandler


173:     ICommand_HelpFile = ""


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpContextID() As Long
  On Error GoTo ErrorHandler


185:     ICommand_HelpContextID = ""


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


210:     ICommand_Category = "ORMAP"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
  On Error GoTo ErrorHandler

    
222:   Set m_pApp = hook
223:   If g_pApp Is Nothing Then
224:     Set g_pApp = hook
225:   End If
226:   Set m_pDoc = m_pApp.Document
227:   Set m_pMap = m_pDoc.FocusMap

  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
'***************************************************************************
'Name:  ICommand_OnClick
'Initial Author:        James Moore
'Subsequent Author:     Type your name here.
'Created:       10/11/2006
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

    'Get the MapIndex feature layer and fclass
    Dim pCLFLayer As IFeatureLayer
    Dim pTaxlotFlayer As IFeatureLayer
    Dim pMIFlayer As IFeatureLayer
    Dim sMsg1 As String
    
264:     Set pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
265:     If pTaxlotFlayer Is Nothing Then
266:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
268:         GoTo Process_Exit
269:     End If
    
271:     Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
272:     If pMIFlayer Is Nothing Then
273:         MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
275:         GoTo Process_Exit
276:     End If
    
278:     Set pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
279:     If pTaxlotFlayer Is Nothing Then
280:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
282:         GoTo Process_Exit
283:     End If
    
285:     Set pCLFLayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
286:     If pCLFLayer Is Nothing Then
287:         MsgBox "Unable to locate Cartographic Lines layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCCartoLines
289:         GoTo Process_Exit
290:     End If
    'Open the form
292:     frmArrows.Show vbModal

Process_Exit:
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
372:   Set m_pTextSymbol = Nothing
373:   Set m_pTextPoint = Nothing
374:   Set m_pLinePolyline = Nothing
375:   Set m_pLineSymbol = Nothing
376:   m_bInUse = False

378:   ITool_Deactivate = True


  Exit Function
ErrorHandler:
  HandleError True, "ITool_Deactivate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  ITool_OnMouseDown
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point
'***************************************************************************
Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler

410:   m_bInUse = True
  Dim pMXDoc As IMxDocument
  Dim pActiveView As IActiveView
413:   Set pMXDoc = m_pApp.Document
414:   Set pActiveView = pMXDoc.FocusMap
  
    'Get reference to MapIndex FC
    Dim sScale As String
    Dim pMIFC As IFeatureClass
    Dim pMIFL As IFeatureLayer
420:     Set pMIFL = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
421:     Set pMIFC = pMIFL.FeatureClass
  
  'If drawing hooks
424:   If frmArrows.lblCurrentTool.Caption = "hook" Then
    'Get point to measure distance from
426:     Set m_pStartPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
427:     m_bDoOnce = False 'testing
428:     Set m_pFromBreakPoint = m_pStartPoint
    'Get the scale of the current mapindex
430:     sScale = GetValueViaOverlay(m_pStartPoint, pMIFC, g_pFldnames.MIMapScaleFN)
431:     sScale = modUtils.ConvertToDescription(pMIFC.Fields, g_pFldnames.MIMapScaleFN, sScale)
  

'If drawing annotation arrows
435:   ElseIf frmArrows.lblCurrentTool.Caption = "arrow" Then
    
437:     If m_pArrowPt1 Is Nothing Then
438:         Set m_pArrowPt1 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
439:     ElseIf m_pArrowPt2 Is Nothing Then
440:         Set m_pArrowPt2 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
441:         DrawArrows
442:     ElseIf m_pArrowPt3 Is Nothing Then
443:         Set m_pArrowPt3 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
444:         DrawArrows
445:     ElseIf m_pArrowPt4 Is Nothing Then
446:         Set m_pArrowPt4 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
447:         DrawArrows
        Dim pArrowLine As IPolyline4
        Dim pArrowPoints As IPointCollection4
450:         Set pArrowPoints = New Polyline
451:         pArrowPoints.AddPoint m_pArrowPt1
452:         pArrowPoints.AddPoint m_pArrowPt2
453:         pArrowPoints.AddPoint m_pArrowPt3
454:         pArrowPoints.AddPoint m_pArrowPt4
455:         Set pArrowLine = pArrowPoints
456:         pArrowLine.Smooth pArrowLine.Length / 10

        Dim pFeature As IFeature
        Dim pWSEdit As IWorkspaceEdit
        Dim pDSet As IDataset
        Dim pArrowLayer As IFeatureLayer
        Dim pArrowFC As IFeatureClass
463:         Set pArrowLayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
464:         If pArrowLayer Is Nothing Then
465:           MsgBox "The layer, " & g_pFldnames.FCCartoLines & ", is not in the map.", vbExclamation, "Layer not found"
466:           m_pEditor.AbortOperation
467:           GoTo Process_Exit
468:         End If
469:         Set pArrowFC = pArrowLayer.FeatureClass
470:         Set pDSet = pArrowFC
471:         Set pWSEdit = pDSet.Workspace
        
        'create the arrow feature
474:         pWSEdit.StartEditOperation
475:         Set pFeature = pArrowFC.CreateFeature
476:         Set pFeature.Shape = pArrowLine
477:         lLineTypeFld = modUtils.LocateFields(pArrowFC, g_pFldnames.CLinesLineTypeFN)
478:         If lLineTypeFld = -1 Then GoTo Process_Exit
479:         pFeature.Value(lLineTypeFld) = 100
        
        'Get the current MapNumber
        Dim sCurMapNum As String
483:         sCurMapNum = GetValueViaOverlay(pFeature.Shape, pMIFC, g_pFldnames.MIMapNumberFN)
        Dim lCLMNfld As Long
485:         lCLMNfld = modUtils.LocateFields(pArrowFC, g_pFldnames.MIMapNumberFN)
486:         pFeature.Value(lCLMNfld) = sCurMapNum
        
488:         pFeature.Store
489:         pWSEdit.StopEditOperation
        
        'Refresh the display
492:         pActiveView.PartialRefresh esriViewBackground, Nothing, pFeature.Extent.Envelope
        
494:         Set pArrowLine = Nothing
495:         Set m_pArrowPt1 = Nothing
496:         Set m_pArrowPt2 = Nothing
497:         Set m_pArrowPt3 = Nothing
498:         Set m_pArrowPt4 = Nothing
        
500:         Set m_pTextSymbol = Nothing
501:         Set m_pTextPoint = Nothing
502:         Set m_pLinePolyline = Nothing
503:         Set m_pLineSymbol = Nothing
        'Deactivate the tool
505:         Set m_pApp.CurrentTool = Nothing
  
  
508:     Else 'Reset everything
509:         Set m_pArrowPt1 = Nothing
510:         Set m_pArrowPt2 = Nothing
511:         Set m_pArrowPt3 = Nothing
512:         Set m_pArrowPt4 = Nothing
        'Deactivate the tool
514:         Set m_pApp.CurrentTool = Nothing
515:     End If
516:   End If
  
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  ITool_OnMouseMove
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point created
'***************************************************************************
Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler

'Draw the temporary line the user sees while moving the mouse
549: If frmArrows.lblCurrentTool.Caption = "hook" Then
550:   If (Not m_bInUse) Then GoTo Process_Exit
  
  Dim pMXDoc As IMxDocument
  Dim pActiveView As IActiveView
554:   Set pMXDoc = m_pApp.Document
555:   Set pActiveView = pMXDoc.FocusMap
  
  Dim bfirstTime As Boolean
558:   If (m_pLineSymbol Is Nothing) Then bfirstTime = True

  'Get current point
  Dim pPoint As IPoint
562:   Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
563:   Set m_pToBreakPoint = pPoint 'the unextended to point used to break the hooks
  
  
   'Draw a virtual line that represents the extended line
   Dim pPline As IPolyline
568:    Set pPline = New Polyline
569:    pPline.FromPoint = m_pStartPoint
570:    pPline.ToPoint = pPoint
   Dim pCv As ICurve
572:    Set pCv = pPline
   
   Dim pCPoint As IConstructPoint
575:    Set pCPoint = New Point
576:    pCPoint.ConstructAlong pCv, esriExtendAtTo, pCv.Length + CDbl(lLineLength), False
577:    Set pPoint = pCPoint
   
579:    If Not m_bDoOnce Then
        Dim pFCPoint As IConstructPoint
581:         Set pFCPoint = New Point
582:         pFCPoint.ConstructAlong pCv, esriExtendAtFrom, -(CDbl(lLineLength)), False
583:         Set m_pStartPoint = pFCPoint
584:         m_bDoOnce = True
585:    End If

  'Draw the line
588:   pActiveView.ScreenDisplay.StartDrawing pActiveView.ScreenDisplay.hdc, -1
    
590:   If bfirstTime Then
    Dim pRGBColor As IRgbColor
    Dim pSymbol As ISymbol
    Dim pFont As IFontDisp
    
    'Line Symbol
596:     Set m_pLineSymbol = New SimpleLineSymbol
597:     m_pLineSymbol.Width = 2
598:     Set pRGBColor = New RgbColor
599:     With pRGBColor
600:       .Red = 223
601:       .Green = 223
602:       .Blue = 223
603:     End With
604:     m_pLineSymbol.Color = pRGBColor
605:     Set pSymbol = m_pLineSymbol
606:     pSymbol.ROP2 = esriROPXOrPen
    
    'Text Symbol
609:     Set m_pTextSymbol = New TextSymbol
610:     m_pTextSymbol.HorizontalAlignment = esriTHACenter
611:     m_pTextSymbol.VerticalAlignment = esriTVACenter
612:     m_pTextSymbol.Size = 16
613:     Set pSymbol = m_pTextSymbol
614:     Set pFont = m_pTextSymbol.Font
615:     pFont.Name = "Arial"
616:     pSymbol.ROP2 = esriROPXOrPen
    
    'Create point to draw text in
619:     Set m_pTextPoint = New Point
    
621:   Else
    'Use existing symbols and draw existing text and polyline
623:     pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
624:     pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
625:     pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
626:     If (m_pLinePolyline.Length > 0) Then _
      pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
628:   End If

  'Get line between from and to points, and angle for text
  Dim pLine As ILine
632:   Set pLine = New esrigeometry.Line
633:   pLine.PutCoords m_pStartPoint, pPoint
  Dim angle As Double
635:   angle = pLine.angle
636:   angle = angle * (180# / 3.14159)
637:   If ((angle > 90#) And (angle < 180#)) Then
638:     angle = angle + 180#
639:   ElseIf ((angle < 0#) And (angle < -90#)) Then
640:     angle = angle - 180#
641:   ElseIf ((angle < -90#) And (angle > -180)) Then
642:     angle = angle - 180#
643:   ElseIf (angle > 180) Then
644:     angle = angle - 180#
645:   End If


  'For drawing text, get text(distance), angle, and point
  Dim deltaX As Double
  Dim deltaY As Double
  Dim distance As Double
652:   deltaX = pPoint.X - m_pStartPoint.X
653:   deltaY = pPoint.Y - m_pStartPoint.Y
654:   m_pTextPoint.X = m_pStartPoint.X + deltaX / 2#
655:   m_pTextPoint.Y = m_pStartPoint.Y + deltaY / 2#
656:   m_pTextSymbol.angle = angle
657:   distance = Round(Sqr((deltaX * deltaX) + (deltaY * deltaY)), 3)
658:   m_pTextSymbol.Text = "" '"[" & distance & "]"
  
  'Draw text
661:   pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
662:   pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
  
  
  'Get polyline with blank space for text
  Dim pPolyline As IPolyline
667:   Set pPolyline = New Polyline
  Dim pSegColl As ISegmentCollection
669:   Set pSegColl = pPolyline
670:   pSegColl.AddSegment pLine
671:   Set m_pLinePolyline = GetSmashedLine(pActiveView.ScreenDisplay, m_pTextSymbol, m_pTextPoint, pPolyline)
  
  'Draw polyline
674:   pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
675:   If (m_pLinePolyline.Length > 0) Then _
    pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
  
678:   pActiveView.ScreenDisplay.FinishDrawing

680: End If
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseMove " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  ITool_OnMouseUp
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler

711:   If frmArrows.lblCurrentTool.Caption = "hook" Then
712:     If (Not m_bInUse) Then GoTo Process_Exit
713:     m_bInUse = False
    
715:     If (m_pLineSymbol Is Nothing) Then GoTo Process_Exit
    
    Dim pMXDoc As IMxDocument
    Dim pActiveView As IActiveView
719:     Set pMXDoc = m_pApp.Document
720:     Set pActiveView = pMXDoc.FocusMap
    
    'Draw temp line
723:     pActiveView.ScreenDisplay.StartDrawing pActiveView.ScreenDisplay.hdc, -1
724:     pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
725:     pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
726:     pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
727:     If (m_pLinePolyline.Length > 0) Then pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
728:     pActiveView.ScreenDisplay.FinishDrawing
729:     If frmArrows.lblCurrentTool.Caption = "hook" Then
    'Generate hooks based on the graphic polyline
731:       GenerateHooks m_pLinePolyline
732:     Else
      'Nothing yet
734:     End If
    
736:     Set m_pTextSymbol = Nothing
737:     Set m_pTextPoint = Nothing
738:     Set m_pLinePolyline = Nothing
739:     Set m_pLineSymbol = Nothing
740:   End If

Process_Exit:
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

'***************************************************************************
'Name:  GetSmashedLine
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Private Function GetSmashedLine(pDisplay As IScreenDisplay, pTextSymbol As ISymbol, pPoint As IPoint, pPolyline As IPolyline) As IPolyline
  On Error GoTo ErrorHandler


  'Returns a graphic Polyline
  Dim pSmashed As IPolyline
  Dim pBoundary As IPolygon
786:   Set pBoundary = New Polygon
787:   pTextSymbol.QueryBoundary pDisplay.hdc, pDisplay.DisplayTransformation, pPoint, pBoundary
  Dim pTopo As ITopologicalOperator
789:   Set pTopo = pBoundary
  
  Dim pIntersect As IPolyline
792:   Set pIntersect = pTopo.Intersect(pPolyline, esriGeometry1Dimension)
793:   Set pTopo = pPolyline
794:   Set GetSmashedLine = pTopo.Difference(pIntersect)


  Exit Function
ErrorHandler:
  HandleError False, "GetSmashedLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  GenerateHooks
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  created single exit point
'***************************************************************************
Private Sub GenerateHooks(pSketch As esrigeometry.IGeometry)
  On Error GoTo ErrorHandler

826: m_dHookAngle = 20
'Make sure the edit sketch is a polyline
828: If Not TypeOf pSketch Is IPolyline Then GoTo Process_Exit

Dim pCurve As ICurve
831: Set pCurve = pSketch

Dim dHookLength As Double
Dim dOverlapLength As Double

'Get the map scale so hook length can be determined
Dim pMIFlayer As IFeatureLayer
Dim pMIFclass As IFeatureClass
839: Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
840: If pMIFlayer Is Nothing Then
841:     MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
    "This process requires a feature class called " & g_pFldnames.FCMapIndex
843:     GoTo Process_Exit
844: End If
845: Set pMIFclass = pMIFlayer.FeatureClass
Dim vMapScale1 As Variant
Dim vMapScale2 As Variant
848: vMapScale1 = modUtils.GetValueViaOverlay(pCurve.FromPoint, pMIFclass, g_pFldnames.MIMapScaleFN)
849: vMapScale2 = modUtils.GetValueViaOverlay(pCurve.ToPoint, pMIFclass, g_pFldnames.MIMapScaleFN)
850: If IsNull(vMapScale1) Or IsNull(vMapScale2) Then
851:     MsgBox "No mapscale for current MapIndex.  Unable to create hooks", vbOKOnly
852:     GoTo Process_Exit
853: End If
854: If vMapScale1 <> vMapScale2 Then
855:     MsgBox "Hook can not span Mapindex polygons with differenc scale", vbCritical
856:     GoTo Process_Exit
857: End If
'ISSUE:  NOT ALL SCALES DEFINED
Dim lLineLength As Long
860: If vMapScale1 = 600 Then
861:     lLineLength = 20
862: ElseIf vMapScale1 = 1200 Then
863:     lLineLength = 20
864: ElseIf vMapScale1 = 2400 Then
865:     lLineLength = 40
866: ElseIf vMapScale1 = 4800 Then
867:     lLineLength = 80
868: ElseIf vMapScale1 = 24000 Then
869:     lLineLength = 400
870: Else
871:     MsgBox "Not a valid mapscale.  Unable to create hooks", vbOKOnly
872:     GoTo Process_Exit
873: End If
874: dHookLength = lLineLength * 0.1
    
'dHookLength = (m_dHookLength * pCurve.Length)

Dim bOK As Boolean
'Make sure the polyline only has two vertices (start and end point only)
880: bOK = IsSketcha2PointLine(pSketch)

882: If bOK = False Then
883:   m_pEditor.AbortOperation
884:   GoTo Process_Exit
885: End If

Dim pApp As esriFramework.IApplication
888: Set pApp = m_pApp 'm_pEditor.Parent
Dim pDoc As esriArcMapUI.IMxDocument
890: Set pDoc = pApp.Document


'Get the hook layer
Dim lLineTypeFld As Long
Dim pWSEdit As IWorkspaceEdit
Dim pDSet As IDataset
Dim pHookLayer As IFeatureLayer
Dim pHookFC As IFeatureClass
899: Set pHookLayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
900: If pHookLayer Is Nothing Then
901:   MsgBox "The layer, " & m_sHookDSName & ", is not in the map.", vbExclamation, "Layer not found"
902:   m_pEditor.AbortOperation
903:   GoTo Process_Exit
904: End If
905: Set pHookFC = pHookLayer.FeatureClass
906: Set pDSet = pHookFC
907: Set pWSEdit = pDSet.Workspace
908: lLineTypeFld = modUtils.LocateFields(pHookFC, g_pFldnames.CLinesLineTypeFN)
909: If lLineTypeFld = -1 Then GoTo Process_Exit
'Create the new hook feature
Dim pFeature As IFeature
Dim pPointColl As esrigeometry.IPointCollection
Dim bCreateDefaultHook As Boolean

915: bCreateDefaultHook = True

917: Set pPointColl = pSketch

Dim dSideA As Double
Dim dSideB As Double
Dim dSideC As Double
Dim pNormal As ILine

'Create a long line that is essentially what the user drew
'Create a hook at each end of the line
'Create hooks based on specified angle and hook length (based on scale)
927: If bCreateDefaultHook = True Then 'Always will be true here
  Dim pNewPointColl As IPointCollection
929:   Set pNewPointColl = New Polyline
    
931:   dSideA = (dHookLength * Sin((360 - m_dHookAngle) * (3.14 / 180)))
932:   dSideC = dHookLength
  
934:   dSideB = (dSideC ^ 2) - (dSideA ^ 2)
935:   dSideB = Sqr(dSideB)
  
937:   Set pNormal = New esrigeometry.Line
938:   pCurve.QueryNormal esriNoExtension, dSideB, False, dSideA, pNormal
  
940:   pNewPointColl.AddPoint pNormal.ToPoint
941:   pNewPointColl.AddPoint pPointColl.Point(0)
942:   pNewPointColl.AddPoint pPointColl.Point(1)
      
944:   dSideA = (dHookLength * Sin(m_dHookAngle * (3.14 / 180)))
945:   dSideC = dHookLength
  
947:   dSideB = (dSideC ^ 2) - (dSideA ^ 2)
948:   dSideB = Sqr(dSideB)
  
950:   pCurve.QueryNormal esriNoExtension, (pCurve.Length - dSideB), False, dSideA, pNormal

952:   pNewPointColl.AddPoint pNormal.ToPoint
  
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
970:   Set pWholeLine = pNewPointColl
971:   pWholeLine.SplitAtPoint m_pFromBreakPoint, True, bBool, bSplitHappened, lNewPartIndex, lNewSegIndex
972:   pWholeLine.SplitAtPoint m_pToBreakPoint, True, bBool, bSplitHappened, lNewPartIndex, lNewSegIndex
  
974:   Set pPath1 = New Path
975:   Set pPath2 = New Path
976:   Set pPath3 = New Path
977:   Set pSegCollection = pWholeLine

979:   Set pEnumSeg = pSegCollection.EnumSegments
980:   pEnumSeg.Next pSeg, lPartIndex, lSegIndex
981:   While Not pSeg Is Nothing
982:     If lSegIndex < 1 Then
983:         pPath1.AddSegment pSeg
984:     ElseIf lSegIndex = 1 Then
985:         pPath2.AddSegment pSeg
986:     ElseIf lSegIndex = 2 Then
987:         pPath3.AddSegment pSeg
988:     End If
989:     pEnumSeg.Next pSeg, lPartIndex, lSegIndex
990:   Wend
991:   Set pGeomColl = New Polyline
992:   pGeomColl.AddGeometry pPath1
993:   pGeomColl.AddGeometry pPath2
994:   pGeomColl.AddGeometry pPath3
995:   pGeomColl.GeometriesChanged
  

  'Store the new hook feature
999:   pWSEdit.StartEditOperation
1000:   Set pFeature = pHookFC.CreateFeature
1001:   Set pFeature.Shape = pGeomColl 'pWholeLine 'pFinalPolyLine 'pNewPointColl
1002:   pFeature.Value(lLineTypeFld) = 101
1003:   pFeature.Store
1004:   pWSEdit.StopEditOperation
  
  'Refresh the display
  Dim pMXDoc As IMxDocument
  Dim pActiveView As IActiveView
1009:   Set pMXDoc = m_pApp.Document
1010:   Set pActiveView = pMXDoc.FocusMap
1011:   pActiveView.PartialRefresh esriViewBackground, Nothing, pFeature.Extent.Envelope
  
1013: End If

Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError False, "GenerateHooks " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  IsSketcha2PointLine
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Public Function IsSketcha2PointLine(pGeom As IGeometry) As Boolean
  On Error GoTo ErrorHandler


Dim pPointColl As IPointCollection
  
1048: If Not TypeOf pGeom Is IPointCollection Then GoTo Process_Exit

1050: Set pPointColl = pGeom
1051: If pPointColl.PointCount <> 2 Then
1052:   MsgBox "When creating the parcel hook only digitize two points", vbExclamation, "Parcel Hook Error"
1053:   GoTo Process_Exit
1054: End If

1056: IsSketcha2PointLine = True

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "IsSketcha2PointLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  ReturnExtended
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Private Function ReturnExtended(pExt As esriSegmentExtension, pPolyline As IPolyline, lLength As Long) As IPolyline
  On Error GoTo ErrorHandler

    Dim pCurve As ICurve
    Dim pLine As ILine
    Dim pPline As IPolyline
    Dim pTopoOpt As ITopologicalOperator
    Dim pGeoCol As IGeometryCollection
    ' Need to extend the end (creates an ILine object)
1094:     Set pCurve = pPolyline
1095:     Set pLine = New esrigeometry.Line
1096:     pCurve.QueryTangent pExt, 1, False, _
                        lLength, pLine
    ' Convert ILine to an IPolyline
1099:     Set pPline = New Polyline
1100:     pPline.FromPoint = pLine.FromPoint
1101:     pPline.ToPoint = pLine.ToPoint
1102:     Set ReturnExtended = pPline


  Exit Function
ErrorHandler:
  HandleError False, "ReturnExtended " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  DrawArrows
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Private Sub DrawArrows()
  On Error GoTo ErrorHandler

    Dim pMXDoc As IMxDocument
    Dim pActiveView As IActiveView
1136:     Set pMXDoc = m_pApp.Document
1137:     Set pActiveView = pMXDoc.FocusMap
    'Set up line symbol to display temporary line
    Dim pRGBColor As IRgbColor
    Dim pSymbol As ISymbol
    Dim pFont As IFontDisp
    'Line Symbol
1143:     Set m_pLineSymbol = New SimpleLineSymbol
1144:     m_pLineSymbol.Width = 2
1145:     Set pRGBColor = New RgbColor
1146:     With pRGBColor
1147:       .Red = 223
1148:       .Green = 223
1149:       .Blue = 223
1150:     End With
1151:     m_pLineSymbol.Color = pRGBColor
1152:     Set pSymbol = m_pLineSymbol
1153:     pSymbol.ROP2 = esriROPXOrPen
    
    Dim pArrowLine As IPointCollection4
    Dim pArrowLine2 As IPolyline4
1157:     Set pArrowLine = New Polyline
1158:     If Not m_pArrowPt1 Is Nothing Then pArrowLine.AddPoint m_pArrowPt1
1159:     If Not m_pArrowPt2 Is Nothing Then pArrowLine.AddPoint m_pArrowPt2
1160:     If Not m_pArrowPt3 Is Nothing Then pArrowLine.AddPoint m_pArrowPt3
1161:     If Not m_pArrowPt4 Is Nothing Then pArrowLine.AddPoint m_pArrowPt4
1162:     Set pArrowLine2 = pArrowLine
    'pArrowLine.Smooth pArrowLine.Length / 10
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''Make lines, then add to segment collection
    '''''''''see mouse move event
    ''''''''''''''''''''''''''''''''''''
        
1170:     pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
1171:     If (pArrowLine2.Length > 0) Then _
        pActiveView.ScreenDisplay.DrawPolyline pArrowLine2
  
1174:     pActiveView.ScreenDisplay.FinishDrawing

    'Set m_pLinePolyline = GetSmashedLine(pActiveView.ScreenDisplay, m_pTextSymbol, m_pTextPoint, pPolyline)

  Exit Sub
ErrorHandler:
  HandleError False, "DrawArrows " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
