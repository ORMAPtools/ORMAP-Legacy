VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdTaxlotAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CLASS MODULE FOR TAXLOTASSIGNMENT COMMAND
'OPENS frmTaxlotAssignment
'INCLUDES A TOOL THAT IS ACTIVATED WHEN THE FORM IS FILLED OUT
Implements ICommand
Implements ITool

Private m_bEnabled As Boolean
Private m_pBitmap As IPictureDisp
Private m_pCursor As IPictureDisp
Private m_pApp As esriFramework.IApplication
Private m_pDoc As esriArcMapUI.IMxDocument
Private m_pMap As IMap
Private m_pTaxlotFlayer As IFeatureLayer2
Private m_pTaxlotFClass As IFeatureClass
Private m_pMIFlayer As IFeatureLayer2
Private m_pMIFclass As IFeatureClass
Private m_lOMTLNumFld As Long
Private m_lOMNumFld As Long
Private m_lTLTaxlotFld As Long
Private m_lTLMNFld As Long
Private m_lCNTaxlotFld As Long
Private m_lCNMNumFld As Long
Private m_lTLAnomalyFld As Long
Private m_pExt As IExtensionConfig
Private WithEvents m_pEditorEvents As Editor
Attribute m_pEditorEvents.VB_VarHelpID = -1
Private m_pEditor As IEditor2
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "C:\active\ModelingWorkshop_01-05-05\CustomCode\ormap\cmdTaxlotAssignment.cls"

 
Private Sub Class_Initialize()
  On Error GoTo ErrorHandler


35:   Set m_pBitmap = LoadResPicture(101, 0)
36:   Set m_pCursor = LoadResPicture(101, 2)


  Exit Sub
ErrorHandler:
  HandleError True, "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
  On Error GoTo ErrorHandler

  
48:   Unload frmTaxlotAssignment
49:   Set m_pBitmap = Nothing
50:   Set m_pCursor = Nothing
51:   Set m_pBitmap = Nothing
52:   Set m_pCursor = Nothing
53:   Set m_pApp = Nothing
54:   Set m_pDoc = Nothing
55:   Set m_pMap = Nothing
56:   Set m_pTaxlotFlayer = Nothing
57:   Set m_pTaxlotFClass = Nothing
58:   Set m_pMIFlayer = Nothing
59:   Set m_pMIFclass = Nothing
60:   Set m_pExt = Nothing
61:   Set m_pEditorEvents = Nothing
62:   Set m_pEditor = Nothing


  Exit Sub
ErrorHandler:
  HandleError True, "Class_Terminate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get ICommand_Checked() As Boolean
  On Error GoTo ErrorHandler



  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler


84:   If m_bEnabled = True Then
85:     ICommand_Enabled = True
86:   End If


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler


98:      ICommand_Name = "TaxlotAssignment"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler


110:      ICommand_Caption = "Taxlot Assignment"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler


122:      ICommand_Tooltip = "Taxlot Assignment Tool"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
  On Error GoTo ErrorHandler

    
134:   ICommand_Message = "Set starting value and increment value to populate values in Taxlots"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpFile() As String
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here
    ' ICommand_HelpFile =


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpContextID() As Long
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here
    ' ICommand_HelpContextID =


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpContextID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
  On Error GoTo ErrorHandler

    
172:    ICommand_Bitmap = m_pBitmap


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Category() As String
  On Error GoTo ErrorHandler


184:      ICommand_Category = "ORMAP"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
  On Error GoTo ErrorHandler

    
196:   Set m_pApp = hook
197:   If g_pApp Is Nothing Then
198:     Set g_pApp = hook
199:   End If
200:   Set m_pDoc = m_pApp.Document
201:   Set m_pMap = m_pDoc.FocusMap
  
  'Get a reference to the editor extension
  Dim pID As New UID
205:   pID = "esriCore.Editor"
206:   Set m_pApp = hook
207:   Set m_pEditor = m_pApp.FindExtensionByCLSID(pID)
208:   Set m_pEditorEvents = m_pEditor


  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
    
 
Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler


    'Set references to feature classes that will be used throughout
    'Find Taxlot
223:     Set m_pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
224:     If m_pTaxlotFlayer Is Nothing Then
225:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        Exit Sub
228:     End If
229:     Set m_pTaxlotFClass = m_pTaxlotFlayer.FeatureClass
    'Find MapIndex
231:     Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
232:     If m_pMIFlayer Is Nothing Then
233:         response = MsgBox("Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex & ".  " & _
        "Load " & g_pFldnames.FCMapIndex & " automatically?", vbYesNo)
        If response <> vbYes Then Exit Sub
237:         modUtils.LoadFCIntoMap g_pFldnames.FCMapIndex, m_pTaxlotFClass
238:         Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        If m_pMIFlayer Is Nothing Then Exit Sub
240:     End If
241:     Set m_pMIFclass = m_pMIFlayer.FeatureClass
242:     m_lOMTLNumFld = m_pTaxlotFClass.FindField(g_pFldnames.TLOrmapTaxlotFN)
243:     If m_lOMTLNumFld = -1 Then
244:         MsgBox "Unable to locate ORMAPTaxlot field.  Check ini file"
        Exit Sub
246:     End If
247:     m_lOMNumFld = m_pTaxlotFClass.FindField(g_pFldnames.TLOrmapMapNumberFN)
248:     If m_lOMNumFld = -1 Then
249:         MsgBox "Unable to locate ORMAPMapnumber field.  Check ini file"
        Exit Sub
251:     End If
252:     m_lTLTaxlotFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLTaxlotFN)
253:     m_lTLAnomalyFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLAnomalyFN)
    'Initialize the TaxlotAssignment form
255:     frmTaxlotAssignment.Show vbModal


  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
  On Error GoTo ErrorHandler

    
267:     ITool_Cursor = m_pCursor
    

  Exit Property
ErrorHandler:
  HandleError True, "ITool_Cursor " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler
  'This code is fired when the user points to a taxlot with the TaxlotAssignment tool
  'Search for underlying MapIndex.  Populate attributes based on MapIndex
  'Assign Taxlot based on input taxlot and optional incremented value settings
    
281:   If Button = 1 Then
    'Save values from Form to variables and increment value in form
    Dim lvalue As Long
    Dim lincrement As Long
    Dim lNextValue As Long
    Dim sFieldName As String
 
    'If "NUMBER" selected, then make sure value is numeric
289:     If frmTaxlotAssignment.cmbTaxlotNum.Text = "NUMBER" Then
290:         lvalue = frmTaxlotAssignment.txtTaxlotNum.Text
291:         If frmTaxlotAssignment.tglBy1.Value = True Then
292:             lincrement = 1
293:         ElseIf frmTaxlotAssignment.tglBy10.Value = True Then
294:             lincrement = 10
295:         Else
296:             lincrement = 100
297:         End If
298:     Else 'If one of the non-numeric categories is selected
        
    
301:     End If

    'Select the target layer feature where user clicked
    Dim pMap As IMap
    Dim pPoint As IPoint
    Dim pGeometry As IGeometry
    Dim pEnv As IEnvelope
    Dim pSpatialFilter As ISpatialFilter
    Dim strShapeFieldName As String
    Dim pFCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim pActiveView As IActiveView

314:     Set pMap = m_pEditor.Map
315:     Set pPoint = m_pEditor.Display.DisplayTransformation.ToMapPoint(X, Y)
316:     Set pGeometry = m_pEditor.CreateSearchShape(pPoint)
317:     If frmTaxlotAssignment.cmbTaxlotNum.Text = "NUMBER" Then
        'Verify that within this map index, this taxlot number is unique
        'If not unique, prompt user to enter a new value
320:         If Not modUtils.ValidateTaxlotNum(frmTaxlotAssignment.txtTaxlotNum.Text, pGeometry) Then
321:             MsgBox "The current Taxlot value (" & frmTaxlotAssignment.txtTaxlotNum.Text & _
            ") is not unique withing this MapIndex.  Please enter a new number"
323:             frmTaxlotAssignment.Show vbModal
            Exit Sub
325:         End If
326:     End If
327:     Set pEnv = pGeometry 'QI
328:     Set pSpatialFilter = New SpatialFilter
329:     Set pSpatialFilter.Geometry = pGeometry
330:     strShapeFieldName = m_pTaxlotFClass.ShapeFieldName
331:     Set pSpatialFilter.OutputSpatialReference(strShapeFieldName) = pMap.SpatialReference
332:     pSpatialFilter.GeometryField = m_pTaxlotFClass.ShapeFieldName
333:     pSpatialFilter.SpatialRel = esriSpatialRelIntersects
334:     Set pFCursor = m_pTaxlotFClass.Search(pSpatialFilter, False)
    If pFCursor Is Nothing Then Exit Sub 'If no taxlots selected
336:     Set pFeature = pFCursor.NextFeature
337:     If Not pFeature Is Nothing Then
      'Update the feature
339:         m_pEditor.StartOperation
        Dim sOMTLNval As String
        Dim sExistOMNum As String
342:         sExistOMNum = IIf(IsNull(pFeature.Value(m_lOMNumFld)), "", pFeature.Value(m_lOMNumFld))
343:         If sExistOMNum = "" Then 'Obtain it from MapIndex polygon
344:             modUtils.CalcTaxlotValues pFeature, m_pMIFlayer
345:             sExistOMNum = IIf(IsNull(pFeature.Value(m_lOMNumFld)), "", pFeature.Value(m_lOMNumFld))
346:             If sExistOMNum = "" Then
347:                 MsgBox "ORMAPMapNumber not present in this taxlot or MapIndex.  Use the MapIndex tool to create one"
                Exit Sub
349:             End If
350:         End If
        'Assign Taxlot value
        Dim sExistTLNumVal As String
353:         sExistTLNumVal = IIf(IsNull(pFeature.Value(m_lTLTaxlotFld)), "", pFeature.Value(m_lTLTaxlotFld))
354:         If Not sExistTLNumVal = "" Then
            Dim response As Integer
356:             response = MsgBox("Taxlot currently has a Taxlot value (" & sExistTLNumVal & ").  Update it?", vbYesNo)
            If response = vbNo Then Exit Sub
358:         End If
        'Taxlot can be less than 5-digits
        'The Taxlot value in OrMapMapNum must be 5 digits.
        'Two versions of the taxlot number will be used for these purposes.
        Dim sNewTLNum As String
        Dim sNewTLNum_5digit As String
364:         If frmTaxlotAssignment.cmbTaxlotNum.Text = "NUMBER" Then
365:             sNewTLNum = frmTaxlotAssignment.txtTaxlotNum.Text 'User entered number
366:             sNewTLNum_5digit = sNewTLNum
            'Make sure number is 5 characters
368:             If Len(sNewTLNum_5digit) < 5 Then
369:              Do Until Len(sNewTLNum_5digit) = 5
370:                 sNewTLNum_5digit = "0" & sNewTLNum_5digit
371:              Loop
372:             End If
373:         Else
            'Remove leading Zeros for taxlot number if any exist
375:             sNewTLNum_5digit = frmTaxlotAssignment.cmbTaxlotNum.Text 'Predefined selection
            'sNewTLNum_5digit = sNewTLNum
377:             sNewTLNum = Replace(sNewTLNum_5digit, "0", "")
378:         End If
        'Determine if Special Interests field is something other than default
        'If so, include it in ORMAPtaxlot
        Dim sTLSpecVAl As String
382:         sTLSpecVAl = GetSpecialInterests(pFeature)
383:         If IsNull(sTLSpecVAl) Then sTLSpecVAl = "00000"
384:         pFeature.Value(m_lTLTaxlotFld) = sNewTLNum
        Dim sShortOMNum As String 'Remove suffixTYpe and suffixNum
386:         sShortOMNum = ShortenOMMapNum(sExistOMNum)
387:         sOMTLNval = sShortOMNum & sNewTLNum_5digit & sTLSpecVAl '"00000"
        'Copy Anomaly from MapIndex
        Dim sAnom As String
390:         sAnom = ParseOMMapNum(sExistOMNum, "anomaly")
391:         pFeature.Value(m_lTLAnomalyFld) = sAnom
        'Assign OrmapTaxlot value
393:         pFeature.Value(m_lOMTLNumFld) = sOMTLNval
394:         pFeature.Store
        'AutoIncrement if necessary
396:         If frmTaxlotAssignment.tglAutoYes.Value = True And frmTaxlotAssignment.cmbTaxlotNum.Text = "NUMBER" Then
397:             lNextValue = lvalue + lincrement
            'Add leading zeros if necessary
            Dim sNextValue As String
400:             sNextValue = CStr(lNextValue)
            'If Len(sNextValue) < 5 Then
             'Do Until Len(sNextValue) = 5
            ''    sNextValue = "0" & sNextValue
             'Loop
           ' End If
406:             frmTaxlotAssignment.txtTaxlotNum.Text = lNextValue
407:         End If
408:       m_pEditor.StopOperation "AutoIncrement Attribute"
409:     If frmTaxlotAssignment.tglAutoYes.Value = False Then
        'Deactivate the tool
411:         Set m_pApp.CurrentTool = Nothing
412:     End If
413:     Else
        'Deactivate the tool
415:         Set m_pApp.CurrentTool = Nothing
        Exit Sub 'Don't do anything if no feature
417:     End If
418:     Set pActiveView = pMap
419:     pActiveView.PartialRefresh esriViewGeography, m_pTaxlotFlayer, pFeature.Extent.Envelope
420:   End If
    

  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here
    

  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseMove " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here
    

  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseUp " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnDblClick()
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here
    

  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnDblClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnKeyDown(ByVal keyCode As Long, ByVal Shift As Long)
  On Error GoTo ErrorHandler

    
468:   If keyCode = vbKeyO Then
469:     frmTaxlotAssignment.Show vbModal
470:   End If

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

    
506:     ITool_Deactivate = True


  Exit Function
ErrorHandler:
  HandleError True, "ITool_Deactivate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub ITool_Refresh(ByVal hdc As esriSystem.OLE_HANDLE)
  On Error GoTo ErrorHandler



  Exit Sub
ErrorHandler:
  HandleError True, "ITool_Refresh " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub m_pEditorEvents_OnChangeFeature(ByVal obj As esriGeoDatabase.IObject)
  On Error GoTo ErrorHandler

    'Code triggered by editing a feature
    'These are the autocalculated fields
    If Not modUtils.IsOrMapFeature(obj) Then Exit Sub
    Dim u As New UID
    Dim pFeat As IFeature
    
533:     UpdateAutoFields obj
534:     If modUtils.IsTaxlot(obj) Then
        'Obtain ORMAPMapnum via overlay and calculate other field values
536:         modUtils.CalcTaxlotValues obj, m_pMIFlayer
537:     ElseIf modUtils.IsAnno(obj) Then
        Dim pAnnotationFeature As IAnnotationFeature
539:         Set pAnnotationFeature = obj
        
        'Get the parent feature so mapnumber can be obtained
        Dim lParentID As Long
543:         lParentID = pAnnotationFeature.LinkedFeatureID
544:         If lParentID > -1 Then 'Feature linked
545:             Set pFeat = GetRelatedObjects(obj)
            If pFeat Is Nothing Then Exit Sub
547:         Else
            'Not feature linked anno, so we can use the feature as is
549:             Set pFeat = obj
550:         End If
    
        'Set anno size
553:         SetAnnoSize obj, pFeat

555:     End If

  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnChangeFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
Private Sub m_pEditorEvents_OnCreateFeature(ByVal obj As esriGeoDatabase.IObject)
  On Error GoTo ErrorHandler

    'Code triggered by creating a new feature
    'These are the autocalculated fields
    If Not IsOrMapFeature(obj) Then Exit Sub
    Dim sMapNum As String
    Dim pMIFlayer As IFeatureLayer
    Dim pMIFclass As IFeatureClass
    Dim pFeat As IFeature
    Dim lAnnoMapNumFld As Long
    Dim pAnnotationFeature As IAnnotationFeature
    Dim lParentID As Long
    Dim pGeometry As IGeometry
    Dim pEnv As IEnvelope
    Dim pCenter As IPoint
    
578:     UpdateAutoFields obj
579:     If modUtils.IsTaxlot(obj) Then
        'Obtain ORMAPMapnum via overlay and calculate other field values
581:         modUtils.CalcTaxlotValues obj, m_pMIFlayer
        
583:     ElseIf modUtils.IsAnno(obj) Then
        'Capture MapNumber for each anno feature created
585:         lAnnoMapNumFld = LocateFields(obj.Class, g_pFldnames.MIMapNumberFN)
        If lAnnoMapNumFld = -1 Then Exit Sub
587:         Set pAnnotationFeature = obj
        
        'Get the parent feature so mapnumber can be obtained
590:         lParentID = pAnnotationFeature.LinkedFeatureID
591:         If lParentID > -1 Then 'Feature linked
592:             Set pFeat = GetRelatedObjects(obj)
            If pFeat Is Nothing Then Exit Sub
594:         Else
            'Not feature linked anno, so we can use the feature as is
596:             Set pFeat = obj
597:         End If

599:         Set pGeometry = pFeat.Shape
        If pGeometry.IsEmpty Then Exit Sub
601:         Set pEnv = pGeometry.Envelope
602:         Set pCenter = CT_GetCenterOfEnvelope(pEnv)
603:         Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        If pMIFlayer Is Nothing Then Exit Sub
605:         Set pMIFclass = pMIFlayer.FeatureClass
606:         sMapNum = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
607:         obj.Value(lAnnoMapNumFld) = sMapNum
        
        'Set size based on mapscale
610:         SetAnnoSize obj, pFeat
611:     End If

  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnCreateFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub m_pEditorEvents_OnDeleteFeature(ByVal obj As esriGeoDatabase.IObject)
  On Error GoTo ErrorHandler
    'This code listens for deletes
    'If deleting taxlots, capture the mapnumber and taxlot and record in CancelledNumbers
    If Not IsOrMapFeature(obj) Then Exit Sub
624:     If modUtils.IsTaxlot(obj) Then
            Dim pDSet As IDataset
            Dim pCNTable As ITable
            Dim pWSpace As IWorkspace
            Dim pFWSpace As IFeatureWorkspace
            Dim pFeat As IFeature
            Dim pRow As IRow
631:             Set pFeat = obj
632:             Set m_pTaxlotFClass = pFeat.Class
633:             Set pDSet = m_pTaxlotFClass
634:             Set pWSpace = pDSet.Workspace
635:             Set pFWSpace = pWSpace
636:             Set pCNTable = pFWSpace.OpenTable(g_pFldnames.TabCancelled)
637:             If pCNTable Is Nothing Then
                'Could privide warning that table is missing, but could get annoying
                Exit Sub
640:             End If
641:             m_lTLTaxlotFld = m_pTaxlotFClass.FindField(g_pFldnames.TLTaxlotFN)
642:             m_lTLMNFld = m_pTaxlotFClass.FindField(g_pFldnames.TLMapNumberFN)
643:             m_lCNTaxlotFld = pCNTable.FindField(g_pFldnames.TLTaxlotFN)
644:             m_lCNMNumFld = pCNTable.FindField(g_pFldnames.TLMapNumberFN)
            If m_lCNTaxlotFld = -1 Or m_lCNMNumFld = -1 Then Exit Sub
            'If no null values, copy them to Cancelled numbers
647:             If Not IsNull(obj.Value(m_lTLTaxlotFld)) And Not IsNull(obj.Value(m_lTLMNFld)) Then
648:                 Set pRow = pCNTable.CreateRow
                If pRow Is Nothing Then Exit Sub
650:                 pRow.Value(m_lCNTaxlotFld) = obj.Value(m_lTLTaxlotFld)
651:                 pRow.Value(m_lCNMNumFld) = obj.Value(m_lTLMNFld)
652:                 pRow.Store
653:                 Set pRow = Nothing
654:             End If
655:         End If

  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnDeleteFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub m_pEditorEvents_OnStartEditing()
  On Error GoTo ErrorHandler

    
666:   m_bEnabled = True


  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnStartEditing " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub m_pEditorEvents_OnStopEditing(ByVal save As Boolean)
  On Error GoTo ErrorHandler


678:   m_bEnabled = False


  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnStopEditing " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




