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
'
' File name:            cmdTaxlotAssignment
'
' Initial Author:       Type your name here
'
' Date Created:     10/11/2006
'
' Description: CLASS MODULE FOR TAXLOTASSIGNMENT COMMAND
'
'
'       Short description of the file's overall purpose.
'
' Entry points:
'       List the public variables and their purposes.
'       List the properties and routines that the module exposes to the rest of the program.
'
' Dependencies: OPENS frmTaxlotAssignment
'               INCLUDES A TOOL THAT IS ACTIVATED WHEN THE FORM IS FILLED OUT
'
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
'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "cmdTaxlotAssignment.cls" ' Variables used by the Error handler function - DO NOT REMOVE
'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------

'implements
Implements ICommand
Implements ITool


Private Sub Class_Initialize()
  On Error GoTo ErrorHandler


107:   Set m_pBitmap = LoadResPicture(101, 0)
108:   Set m_pCursor = LoadResPicture(101, 2)


  Exit Sub
ErrorHandler:
  HandleError True, "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
  On Error GoTo ErrorHandler

  
120:   Unload frmTaxlotAssignment
121:   Set m_pBitmap = Nothing
122:   Set m_pCursor = Nothing
123:   Set m_pBitmap = Nothing
124:   Set m_pCursor = Nothing
125:   Set m_pApp = Nothing
126:   Set m_pDoc = Nothing
127:   Set m_pMap = Nothing
128:   Set m_pTaxlotFlayer = Nothing
129:   Set m_pTaxlotFClass = Nothing
130:   Set m_pMIFlayer = Nothing
131:   Set m_pMIFclass = Nothing
132:   Set m_pExt = Nothing
133:   Set m_pEditorEvents = Nothing
134:   Set m_pEditor = Nothing


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


156:   If m_bEnabled = True Then
157:     ICommand_Enabled = True
158:   End If


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler


170:      ICommand_Name = "TaxlotAssignment"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler


182:      ICommand_Caption = "Taxlot Assignment"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler


194:      ICommand_Tooltip = "Taxlot Assignment Tool"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
  On Error GoTo ErrorHandler

    
206:   ICommand_Message = "Set starting value and increment value to populate values in Taxlots"


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

    
244:    ICommand_Bitmap = m_pBitmap


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Category() As String
  On Error GoTo ErrorHandler


256:      ICommand_Category = "ORMAP"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
  On Error GoTo ErrorHandler

    
268:   Set m_pApp = hook
269:   If g_pApp Is Nothing Then
270:     Set g_pApp = hook
271:   End If
272:   Set m_pDoc = m_pApp.Document
273:   Set m_pMap = m_pDoc.FocusMap
  
  'Get a reference to the editor extension
  Dim pID As New UID
277:   pID = "esriCore.Editor"
278:   Set m_pApp = hook
279:   Set m_pEditor = m_pApp.FindExtensionByCLSID(pID)
280:   Set m_pEditorEvents = m_pEditor


  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
    
 
'***************************************************************************
'Name:  ICommand_OnClick
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:       Set references to feature classes that will be used throughout
'Called From:
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
'James Moore    10/11/2006  defined single exit point
'***************************************************************************
Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

    'Find Taxlot
313:     Set m_pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
314:     If m_pTaxlotFlayer Is Nothing Then
315:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
317:         GoTo Process_Exit
318:     End If
319:     Set m_pTaxlotFClass = m_pTaxlotFlayer.FeatureClass
    'Find MapIndex
321:     Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
322:     If m_pMIFlayer Is Nothing Then
323:         response = MsgBox("Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex & ".  " & _
        "Load " & g_pFldnames.FCMapIndex & " automatically?", vbYesNo)
326:         If response <> vbYes Then GoTo Process_Exit
327:         modUtils.LoadFCIntoMap g_pFldnames.FCMapIndex, m_pTaxlotFClass
328:         Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
329:         If m_pMIFlayer Is Nothing Then GoTo Process_Exit
330:     End If
331:     Set m_pMIFclass = m_pMIFlayer.FeatureClass
332:     m_lOMTLNumFld = m_pTaxlotFClass.FindField(g_pFldnames.TLOrmapTaxlotFN)
333:     If m_lOMTLNumFld = -1 Then
334:         MsgBox "Unable to locate ORMAPTaxlot field.  Check ini file"
335:         GoTo Process_Exit
336:     End If
337:     m_lOMNumFld = m_pTaxlotFClass.FindField(g_pFldnames.TLOrmapMapNumberFN)
338:     If m_lOMNumFld = -1 Then
339:         MsgBox "Unable to locate ORMAPMapnumber field.  Check ini file"
340:         GoTo Process_Exit
341:     End If
342:     m_lTLTaxlotFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLTaxlotFN)
343:     m_lTLAnomalyFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLAnomalyFN)
    'Initialize the TaxlotAssignment form
345:     frmTaxlotAssignment.Show vbModal

Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
  On Error GoTo ErrorHandler

    
357:     ITool_Cursor = m_pCursor
    

  Exit Property
ErrorHandler:
  HandleError True, "ITool_Cursor " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
'***************************************************************************
'Name:  ITool_OnMouseDown
'Initial Author:        James Moore
'Subsequent Author:     Type your name here.
'Created:       10/11/2006
'Purpose:
'Called From:
'Description:  'This code is fired when the user points to a taxlot with the TaxlotAssignment tool
'Methods:   Search for underlying MapIndex.  Populate attributes based on MapIndex
'           Assign Taxlot based on input taxlot and optional incremented value settings
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
'James Moore    10/11/2006  using strcomp function. Changed to intrinsic constant, established single exit point
'***************************************************************************
Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler
    
390:   If Button = vbLeftButton Then
    'Save values from Form to variables and increment value in form
    Dim lvalue As Long
    Dim lincrement As Long
    Dim lNextValue As Long
    Dim sFieldName As String

    'If "NUMBER" selected, then make sure value is numeric
398:     If StrComp(frmTaxlotAssignment.cmbTaxlotNum.Text, "NUMBER", vbTextCompare) = 0 Then
399:         lvalue = frmTaxlotAssignment.txtTaxlotNum.Text
400:         If frmTaxlotAssignment.tglBy1.Value = True Then
401:             lincrement = 1
402:         ElseIf frmTaxlotAssignment.tglBy10.Value = True Then
403:             lincrement = 10
404:         Else
405:             lincrement = 100
406:         End If
407:     Else 'If one of the non-numeric categories is selected
        ' dont have a else if there is nothing to do

410:     End If

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

423:     Set pMap = m_pEditor.Map
424:     Set pPoint = m_pEditor.Display.DisplayTransformation.ToMapPoint(X, Y)
425:     Set pGeometry = m_pEditor.CreateSearchShape(pPoint)
426:     If frmTaxlotAssignment.cmbTaxlotNum.Text = "NUMBER" Then
        'Verify that within this map index, this taxlot number is unique
        'If not unique, prompt user to enter a new value
429:         If Not modUtils.ValidateTaxlotNum(frmTaxlotAssignment.txtTaxlotNum.Text, pGeometry) Then
430:             MsgBox "The current Taxlot value (" & frmTaxlotAssignment.txtTaxlotNum.Text & _
            ") is not unique withing this MapIndex.  Please enter a new number"
432:             frmTaxlotAssignment.Show vbModal
433:             GoTo Process_Exit
434:         End If
435:     End If
436:     Set pEnv = pGeometry 'QI
437:     Set pSpatialFilter = New SpatialFilter
438:     Set pSpatialFilter.Geometry = pGeometry
439:     strShapeFieldName = m_pTaxlotFClass.ShapeFieldName
440:     Set pSpatialFilter.OutputSpatialReference(strShapeFieldName) = pMap.SpatialReference
441:     pSpatialFilter.GeometryField = m_pTaxlotFClass.ShapeFieldName
442:     pSpatialFilter.SpatialRel = esriSpatialRelIntersects
443:     Set pFCursor = m_pTaxlotFClass.Search(pSpatialFilter, False)
444:     If pFCursor Is Nothing Then GoTo Process_Exit 'If no taxlots selected
445:     Set pFeature = pFCursor.NextFeature
446:     If Not pFeature Is Nothing Then
      'Update the feature
448:         m_pEditor.StartOperation
        Dim sOMTLNval As String
        Dim sExistOMNum As String
451:         sExistOMNum = IIf(IsNull(pFeature.Value(m_lOMNumFld)), "", pFeature.Value(m_lOMNumFld))
452:         If sExistOMNum = "" Then 'Obtain it from MapIndex polygon
453:             modUtils.CalcTaxlotValues pFeature, m_pMIFlayer
454:             sExistOMNum = IIf(IsNull(pFeature.Value(m_lOMNumFld)), "", pFeature.Value(m_lOMNumFld))
455:             If sExistOMNum = "" Then
456:                 MsgBox "ORMAPMapNumber not present in this taxlot or MapIndex.  Use the MapIndex tool to create one"
457:                 GoTo Process_Exit
458:             End If
459:         End If
        'Assign Taxlot value
        Dim sExistTLNumVal As String
462:         sExistTLNumVal = IIf(IsNull(pFeature.Value(m_lTLTaxlotFld)), "", pFeature.Value(m_lTLTaxlotFld))
463:         If Not sExistTLNumVal = "" Then
            Dim response As Integer
465:             response = MsgBox("Taxlot currently has a Taxlot value (" & sExistTLNumVal & ").  Update it?", vbYesNo)
466:             If response = vbNo Then GoTo Process_Exit
467:         End If
        'Taxlot can be less than 5-digits
        'The Taxlot value in OrMapMapNum must be 5 digits.
        'Two versions of the taxlot number will be used for these purposes.
        Dim sNewTLNum As String
        Dim sNewTLNum_5digit As String
473:         If StrComp(frmTaxlotAssignment.cmbTaxlotNum.Text, "NUMBER", vbTextCompare) = 0 Then
474:             sNewTLNum = frmTaxlotAssignment.txtTaxlotNum.Text 'User entered number
475:             sNewTLNum_5digit = sNewTLNum
            'Make sure number is 5 characters
477:             If Len(sNewTLNum_5digit) < 5 Then
478:              Do Until Len(sNewTLNum_5digit) = 5
479:                 sNewTLNum_5digit = "0" & sNewTLNum_5digit
480:              Loop
481:             End If
482:         Else
            'Remove leading Zeros for taxlot number if any exist
484:             sNewTLNum_5digit = frmTaxlotAssignment.cmbTaxlotNum.Text 'Predefined selection
            'sNewTLNum_5digit = sNewTLNum
486:             sNewTLNum = Replace(sNewTLNum_5digit, "0", "")
487:         End If
        '++ BEGIN, Laura Gordon, November 29, 2005
        'Determine if Special Interests field is something other than default
        'If so, include it in ORMAPtaxlot
        '+Dim sTLSpecVAl As String
        '+sTLSpecVAl = GetSpecialInterests(pFeature)
        '+If IsNull(sTLSpecVAl) Then sTLSpecVAl = "00000"
        Dim sTLMapSufTypeVAl As String
        Dim sTLMapSufNumVAl As String
496:         sTLMapSufTypeVAl = GetMapSufType(pFeature)
497:         sTLMapSufNumVAl = GetMapSufNum(pFeature)
498:         If IsNull(sTLMapSufTypeVAl) Then sTLMapSufTypeVAl = "0"
499:         If IsNull(sTLMapSufNumVAl) Then sTLMapSufNumVAl = "000"
        '++ END, Laura Gordon, November 29, 2005

502:         pFeature.Value(m_lTLTaxlotFld) = sNewTLNum
        Dim sShortOMNum As String 'Remove suffixTYpe and suffixNum
504:         sShortOMNum = ShortenOMMapNum(sExistOMNum)
        '++ BEGIN, Laura Gordon, November 29, 2005
        '+sOMTLNval = sShortOMNum & sNewTLNum_5digit & sTLSpecVAl '"00000"
507:         sOMTLNval = sShortOMNum & sTLMapSufTypeVAl & sTLMapSufNumVAl & sNewTLNum_5digit
        '++ END, Laura Gordon, November 29, 2005

        'Copy Anomaly from MapIndex
        Dim sAnom As String
512:         sAnom = ParseOMMapNum(sExistOMNum, "anomaly")
513:         pFeature.Value(m_lTLAnomalyFld) = sAnom
        'Assign OrmapTaxlot value
515:         pFeature.Value(m_lOMTLNumFld) = sOMTLNval
516:         pFeature.Store
        'AutoIncrement if necessary
518:         If frmTaxlotAssignment.tglAutoYes.Value = True And StrComp(frmTaxlotAssignment.cmbTaxlotNum.Text, "NUMBER", vbTextCompare) = 0 Then
519:             lNextValue = lvalue + lincrement
            'Add leading zeros if necessary
            Dim sNextValue As String
522:             sNextValue = CStr(lNextValue)
            'If Len(sNextValue) < 5 Then
             'Do Until Len(sNextValue) = 5
            ''    sNextValue = "0" & sNextValue
             'Loop
           ' End If
528:             frmTaxlotAssignment.txtTaxlotNum.Text = lNextValue
529:         End If
530:       m_pEditor.StopOperation "AutoIncrement Attribute"
531:     If frmTaxlotAssignment.tglAutoYes.Value = False Then
        'Deactivate the tool
533:         Set m_pApp.CurrentTool = Nothing
534:     End If
535:     Else
        'Deactivate the tool
537:         Set m_pApp.CurrentTool = Nothing
538:         GoTo Process_Exit 'Don't do anything if no feature
539:     End If
540:     Set pActiveView = pMap
541:     pActiveView.PartialRefresh esriViewGeography, m_pTaxlotFlayer, pFeature.Extent.Envelope
542:   End If
    
Process_Exit:
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

    
590:   If keyCode = vbKeyO Then
591:     frmTaxlotAssignment.Show vbModal
592:   End If

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

    
628:     ITool_Deactivate = True


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

'***************************************************************************
'Name:  m_pEditorEvents_OnChangeFeature
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Code triggered by editing a feature.
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
Private Sub m_pEditorEvents_OnChangeFeature(ByVal obj As esriGeoDatabase.IObject)
  On Error GoTo ErrorHandler

    '
    'These are the autocalculated fields
671:     UpdateAutoFields obj
        
    If Not modUtils.IsOrMapFeature(obj) Then Exit Sub
    Dim u As New UID
    Dim pFeat As IFeature
    
677:     If modUtils.IsTaxlot(obj) Then
        'Obtain ORMAPMapnum via overlay and calculate other field values
679:         modUtils.CalcTaxlotValues obj, m_pMIFlayer
680:     ElseIf modUtils.IsAnno(obj) Then
        Dim pAnnotationFeature As IAnnotationFeature
682:         Set pAnnotationFeature = obj
        
        'Get the parent feature so mapnumber can be obtained
        Dim lParentID As Long
686:         lParentID = pAnnotationFeature.LinkedFeatureID
687:         If lParentID > -1 Then 'Feature linked
688:             Set pFeat = GetRelatedObjects(obj)
            If pFeat Is Nothing Then Exit Sub
690:         Else
            'Not feature linked anno, so we can use the feature as is
692:             Set pFeat = obj
693:         End If
    
        'Set anno size
696:         SetAnnoSize obj, pFeat

698:     End If

  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnChangeFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
'***************************************************************************
'Name:  m_pEditorEvents_OnCreateFeature
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Code triggered by creating a new feature.
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
'James Moore    10/11/2006  defined single exit point
'***************************************************************************
Private Sub m_pEditorEvents_OnCreateFeature(ByVal obj As esriGeoDatabase.IObject)
  On Error GoTo ErrorHandler

    'These are the autocalculated fields
729:     UpdateAutoFields obj
        
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
    Dim lMapScaleFld As Long
    Dim lMapNumFld As Long
    Dim vMapScaleVal As Variant
    Dim vMapNumVal As Variant

746:     lMapNumFld = obj.Fields.FindField(g_pFldnames.MIMapNumberFN)
747:     lMapScaleFld = obj.Fields.FindField(g_pFldnames.MIMapScaleFN)
748:     If modUtils.IsTaxlot(obj) Then
        'Obtain ORMAPMapnum via overlay and calculate other field values
750:         modUtils.CalcTaxlotValues obj, m_pMIFlayer
        
752:     ElseIf modUtils.IsAnno(obj) Then
        'Capture MapNumber for each anno feature created
754:         lAnnoMapNumFld = LocateFields(obj.Class, g_pFldnames.MIMapNumberFN)
755:         If lAnnoMapNumFld = -1 Then GoTo Process_Exit
756:         Set pAnnotationFeature = obj
        
        'Get the parent feature so mapnumber can be obtained
759:         lParentID = pAnnotationFeature.LinkedFeatureID
760:         If lParentID > -1 Then 'Feature linked
761:             Set pFeat = GetRelatedObjects(obj)
762:             If pFeat Is Nothing Then GoTo Process_Exit
763:         Else
            'Not feature linked anno, so we can use the feature as is
765:             Set pFeat = obj
766:         End If

768:         Set pGeometry = pFeat.Shape
769:         If pGeometry.IsEmpty Then GoTo Process_Exit
770:         Set pEnv = pGeometry.Envelope
771:         Set pCenter = CT_GetCenterOfEnvelope(pEnv)
772:         Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
773:         If pMIFlayer Is Nothing Then GoTo Process_Exit
774:         Set pMIFclass = pMIFlayer.FeatureClass
775:         sMapNum = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
776:         obj.Value(lAnnoMapNumFld) = sMapNum
777:         If lMapScaleFld > -1 Then
778:             vMapScaleVal = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
779:             If Not IsNull(vMapScaleVal) Then
780:                 obj.Value(lMapScaleFld) = vMapScaleVal
781:             End If
782:         End If
        'Set size based on mapscale
784:         SetAnnoSize obj, pFeat
785:     Else
        'Update MapScale for all features with a MapScale field (except MapIndex)
787:         If lMapScaleFld > -1 And Not modUtils.IsMapIndex(obj) Then
788:             Set pFeat = obj
789:             Set pGeometry = pFeat.Shape
790:             If pGeometry.IsEmpty Then GoTo Process_Exit
791:             Set pEnv = pGeometry.Envelope
792:             Set pCenter = CT_GetCenterOfEnvelope(pEnv)
793:             Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
794:             If pMIFlayer Is Nothing Then GoTo Process_Exit
795:             Set pMIFclass = pMIFlayer.FeatureClass
796:             vMapScaleVal = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
797:             If Not IsNull(vMapScaleVal) Then
798:                 obj.Value(lMapScaleFld) = vMapScaleVal
799:             End If
            'If a dataset with MapNumber, populate it
801:             If lMapNumFld > -1 Then
802:                 vMapNumVal = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
803:                 If Not IsNull(vMapNumVal) Then
804:                     obj.Value(lMapNumFld) = vMapNumVal
805:                 End If
806:             End If
807:         End If
808:     End If
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnCreateFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


'***************************************************************************
'Name:  m_pEditorEvents_OnDeleteFeature
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
'James Moore    10/11/2006  defined single entry point
'***************************************************************************
Private Sub m_pEditorEvents_OnDeleteFeature(ByVal obj As esriGeoDatabase.IObject)
  On Error GoTo ErrorHandler
    'This code listens for deletes
    'If deleting taxlots, capture the mapnumber and taxlot and record in CancelledNumbers
841:     If Not IsOrMapFeature(obj) Then GoTo Process_Exit
842:     If modUtils.IsTaxlot(obj) Then
            Dim pDSet As IDataset
            Dim pCNTable As ITable
            Dim pWSpace As IWorkspace
            Dim pFWSpace As IFeatureWorkspace
            Dim pFeat As IFeature
            Dim pRow As IRow
849:             Set pFeat = obj
850:             Set m_pTaxlotFClass = pFeat.Class
851:             Set pDSet = m_pTaxlotFClass
852:             Set pWSpace = pDSet.Workspace
853:             Set pFWSpace = pWSpace
854:             Set pCNTable = pFWSpace.OpenTable(g_pFldnames.TabCancelled)
855:             If pCNTable Is Nothing Then
                'Could privide warning that table is missing, but could get annoying
857:                 GoTo Process_Exit
858:             End If
859:             m_lTLTaxlotFld = m_pTaxlotFClass.FindField(g_pFldnames.TLTaxlotFN)
860:             m_lTLMNFld = m_pTaxlotFClass.FindField(g_pFldnames.TLMapNumberFN)
861:             m_lCNTaxlotFld = pCNTable.FindField(g_pFldnames.TLTaxlotFN)
862:             m_lCNMNumFld = pCNTable.FindField(g_pFldnames.TLMapNumberFN)
863:             If m_lCNTaxlotFld = -1 Or m_lCNMNumFld = -1 Then GoTo Process_Exit
            'If no null values, copy them to Cancelled numbers
865:             If Not IsNull(obj.Value(m_lTLTaxlotFld)) And Not IsNull(obj.Value(m_lTLMNFld)) Then
866:                 Set pRow = pCNTable.CreateRow
867:                 If pRow Is Nothing Then GoTo Process_Exit
868:                 pRow.Value(m_lCNTaxlotFld) = obj.Value(m_lTLTaxlotFld)
869:                 pRow.Value(m_lCNMNumFld) = obj.Value(m_lTLMNFld)
870:                 pRow.Store
871:                 Set pRow = Nothing
872:             End If
873:         End If
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnDeleteFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub m_pEditorEvents_OnStartEditing()
  On Error GoTo ErrorHandler

    
884:   m_bEnabled = True


  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnStartEditing " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub m_pEditorEvents_OnStopEditing(ByVal save As Boolean)
  On Error GoTo ErrorHandler


896:   m_bEnabled = False


  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnStopEditing " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




