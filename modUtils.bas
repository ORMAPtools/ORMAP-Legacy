Attribute VB_Name = "modUtils"
'
' File name:            modUtils
'
' Initial Author:       Type your name here
'
' Date Created:
'
' Description:
'       GENERAL UTILITY MODULE
'MOST COMMONLY USED PROCEDURES ARE LOCATED HERE
'
' Entry points:
'       List the public variables and their purposes.
'       List the properties and routines that the module exposes to the rest of the program.
'
' Dependencies:
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
'            JWM 10/11/2006 added comment header to each function
'

Option Explicit
'******************************
' Global/Public Definitions
'------------------------------
' Public API Declarations
'------------------------------

'------------------------------
' Public Enums and Constants
'------------------------------
Const LB_FINDSTRING = &H18F
Const LB_FINDSTRINGEXACT = &H1A2
Const CB_FINDSTRING = &H14C
Const CB_FINDSTRINGEXACT = &H158
' Variables used by the Error handler function - DO NOT REMOVE
'++ JWM 10/11/2006 Reomved the path to this module as it will not always be in the same place
Const c_sModuleFileName As String = "modUtils.bas"
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
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long
'------------------------------
' Private Variables
'------------------------------
Private m_bContinue As Boolean
'------------------------------
'Private Constants and Enums
'------------------------------

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------
'
'***************************************************************************
'Name:  FindFeatureLayerByDS
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Return the Feature Layer based on its dataset name
'               This is an easy way to locate a feature layerr in the TOC.
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
Public Function FindFeatureLayerByDS(DatasetName As String) As IFeatureLayer
  On Error GoTo ErrorHandler
    
    Dim pMXDoc As IMxDocument
    Dim pMap As IMap
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim i As Integer
    
118:     Set pMXDoc = g_pApp.Document
119:     Set pMap = pMXDoc.FocusMap
  
121:     With pMap
122:         For i = 0 To .LayerCount - 1
123:             If TypeOf .Layer(i) Is IFeatureLayer Then
124:                 Set pFeatureLayer = .Layer(i)
125:                 Set pDataset = pFeatureLayer.FeatureClass
126:                 If Not pDataset Is Nothing Then
'++ JWM 10/11/2006 using strcomp function
128:                     If StrComp(pDataset.Name, DatasetName, vbTextCompare) = 0 Then
129:                         Set FindFeatureLayerByDS = pFeatureLayer
130:                         Exit For
131:                     End If
132:                 End If
133:             End If
134:         Next i
135:     End With
  
137:     If pFeatureLayer Is Nothing Then

139:     End If
Process_Exit:
    Exit Function
ErrorHandler:
  HandleError True, "FindFeatureLayerByDS " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Public Function GetFWorkspace(pObj As esriGeoDatabase.IObject) As IFeatureWorkspace
  On Error GoTo ErrorHandler


  Dim pFWS As IFeatureWorkspace
  Dim pObjClass As IObjectClass
  Dim pDataset As IDataset
154:   Set pObjClass = pObj.Class
155:   Set pDataset = pObjClass
156:   Set pFWS = pDataset.Workspace
157:   Set GetFWorkspace = pFWS

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "GetFWorkspace " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  ReadValue
'Initial Author: Chris Buhi
'Subsequent Author:
'Created:
'Purpose:   Reads a value from a row, given a field name
'
'Called From:   frmMapIndex.InitForm, modutils.CompareAndSaveValue

'Methods:       Describe any complex details.
'Parameters:    What variables are brought into this routine?
'Outputs:       What variables are changed in this routine?
'Returns:       If a domain field, the descriptive value is returned instead of the stored code
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Replaced all Exit Function with Goto Process_Exit to have a single exit point
'***************************************************************************
Public Function ReadValue(pRow As IRow, pFldName As String, Optional pDataType As String) As Variant
  On Error GoTo ErrorHandler

    Dim sVal As String
    Dim lFld As Long
191:     lFld = pRow.Fields.FindField(pFldName)
192:     If lFld > -1 Then
193:       If pDataType = "date" Then
        'If a date and value is null, return a default date value
        '??? How should this be treated?
        Dim pDate As Date
197:         sVal = IIf(IsNull(pRow.Value(lFld)), pDate, pRow.Value(lFld))
198:       Else
199:         sVal = IIf(IsNull(pRow.Value(lFld)), "", pRow.Value(lFld))
200:       End If
      'Determine if domain field
      Dim pField As IField
203:       Set pField = pRow.Fields.Field(lFld)
      Dim pDomain As IDomain
205:       Set pDomain = pField.Domain
206:       If pDomain Is Nothing Then
207:         ReadValue = sVal
208:         GoTo Process_Exit
209:       Else
        'Determine type of domain  -If Coded Value, get the description
211:         If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
213:           Set pCVDomain = pDomain
          Dim lCode As Long
          Dim vDomainVal As Variant
216:           vDomainVal = pRow.Value(lFld)
          Dim i As Integer
          'Search the domain for the code
219:           For i = 0 To pCVDomain.CodeCount - 1
220:              If pCVDomain.Value(i) = vDomainVal Then
              'return the description
222:               ReadValue = pCVDomain.Name(i)
223:               GoTo Process_Exit
224:             End If
225:           Next i
226:         Else ' If range domain, return the numeric value
227:           ReadValue = sVal
228:           GoTo Process_Exit:
229:         End If
230:       End If  'If pDomain is nothing/Else
231:       ReadValue = sVal
232:     Else
      'Field not found
234:       ReadValue = ""
235:     End If 'If lFld > -1/Else

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "ReadValue " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


'***************************************************************************
'Name:  AddCodesToCmb
'Initial Author:        Chris Buhi
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Add the descriptive values from each domain to the drop down comboboxes
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point using Goto and made optional parameter true
'***************************************************************************
Public Function AddCodesToCmb(pFldName As String, _
                              pFields As IFields, _
                              cboValues As ComboBox, _
                              curVal As Variant, _
                              Optional blnAllowSpace As Boolean = True) As Boolean
  On Error GoTo ErrorHandler

    
'jwm    If IsMissing(blnAllowSpace) Then blnAllowSpace = True
  
   'Get the Coded Value Domain from the field
      Dim lFld As Long
277:       lFld = pFields.FindField(pFldName)
278:       If lFld > -1 Then
        Dim pField As IField
280:         Set pField = pFields.Field(lFld)
        Dim pDomain As IDomain
282:         Set pDomain = pField.Domain
283:         If pDomain Is Nothing Then
284:           AddCodesToCmb = False
285:           GoTo Process_Exit
286:         Else
          'Determine type of domain  -If Coded Value, get the description
288:           If TypeOf pDomain Is ICodedValueDomain Then
            Dim pCVDomain As ICodedValueDomain
290:             Set pCVDomain = pDomain
            ' +++ Get a count of the coded values
            Dim lCodes As Long
            Dim i As Long
294:             lCodes = pCVDomain.CodeCount
            Dim sVal As Variant
            ' +++ Loop through the list of values and add them
            ' +++ and their names to the combo box
298:             If Not blnAllowSpace Then
299:               With cboValues
300:               If .ListCount > 0 Then
301:                 If (.List(0) = "") Or (.List(0) = "") Then
302:                   .RemoveItem (0)
303:                 End If
304:               End If
305:               If .ListCount > 0 Then
'++ JWM 10/11/2006 Is this if statement comparing against the same thing ?
307:                 If (.List(.ListCount - 1) = "") Or (.List(.ListCount - 1) = "") Then
308:                   .RemoveItem (.ListCount - 1)
309:                 End If
310:               End If
311:               End With
312:             End If
313:             For i = 0 To lCodes - 1
              'Commented line adds codes and description
              'cboValues.AddItem pCVDomain.Value(i) & ": " & pCVDomain.Name(i)
316:               cboValues.AddItem pCVDomain.Name(i)
317:             Next i
            'Successful completion of addition
            'If current value is null, add an empty string and make it active
320:             If curVal = "" Then
321:               If blnAllowSpace Then
322:                 cboValues.AddItem ""
323:                 cboValues.ListIndex = FindControlString(cboValues, "", 0, True)
                'cboValues.Text = ""
325:               Else
326:                 cboValues.ListIndex = 0
327:               End If
328:             Else 'Otherwise, select the existing value from the list
329:               cboValues.ListIndex = FindControlString(cboValues, curVal, 0, True)
330:             End If
            
332:             AddCodesToCmb = True
333:           Else
            'if Range Domain, do not add values
335:             AddCodesToCmb = False
336:           End If
337:         End If 'if a valid domain
338:       Else 'Field not found
339:         AddCodesToCmb = False
340:       End If

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "AddCodesToCmb " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  ConvertCode
'Initial Author:        Chris Buhi
'Subsequent Author:     Type your name here.
'Created:       10/11/2006
'Purpose:       Converts a domain descriptive value to the stored code
'Called From:   frmMapIndex.cmdAssign_Click, modutils.compareandsavevalue, modutils.calctaxlotvalue
'Description:   Domain values chosen from combo boxes must be converted to the code before being stored.
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
'James Moore    10/11/2006  Single exit point using goto
'***************************************************************************
Public Function ConvertCode(pRow As IRow, pFldName As String, pVal As Variant) As Variant
  On Error GoTo ErrorHandler

    Dim lFld As Long
373:     lFld = pRow.Fields.FindField(pFldName)
374:     If lFld > -1 Then
      'Determine if domain field
      Dim pField As IField
377:       Set pField = pRow.Fields.Field(lFld)
      Dim pDomain As IDomain
379:       Set pDomain = pField.Domain
380:       If pDomain Is Nothing Then
381:         ConvertCode = pVal
382:         GoTo Process_Exit
383:       Else
        'Determine type of domain  -If Coded Value, get the description
385:         If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
387:           Set pCVDomain = pDomain
          Dim lCode As Long
          Dim i As Integer
          'Given the description, search the domain for the code
391:           For i = 0 To pCVDomain.CodeCount - 1
392:             If pCVDomain.Name(i) = pVal Then
393:               ConvertCode = pCVDomain.Value(i) 'Return the code value
394:               GoTo Process_Exit
395:             End If
396:           Next i
397:         Else ' If range domain, return the numeric value
398:           ConvertCode = pVal
399:           GoTo Process_Exit
400:         End If
401:       End If  'If pDomain is nothing/Else
402:       ConvertCode = pVal
403:     Else
      'Field not found
405:       ConvertCode = ""
406:     End If 'If lFld > -1/Else

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "ConvertCode " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
 
'***************************************************************************
'Name:  ConvertToDescription
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:       Converts a domain descriptive value to the stored code
'Called From:   frmMapIndex.InitForm
'Description:   Domain values chosen from combo boxes must be converted to the code
'               before being stored
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
'James Moore    10/11/2006  Single exit point using goto
'***************************************************************************
Public Function ConvertToDescription(pFlds As IFields, pFldName As String, pVal As Variant) As Variant
  On Error GoTo ErrorHandler

    Dim lFld As Long
440:     lFld = pFlds.FindField(pFldName)
441:     If lFld > -1 Then
      'Determine if domain field
      Dim pField As IField
444:       Set pField = pFlds.Field(lFld)
      Dim pDomain As IDomain
446:       Set pDomain = pField.Domain
447:       If pDomain Is Nothing Then
448:         ConvertToDescription = pVal
449:         GoTo Process_Exit
450:       Else
        'Determine type of domain  -If Coded Value, get the description
452:         If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
454:           Set pCVDomain = pDomain
          Dim lCode As Long
          Dim i As Integer
          'Given the description, search the domain for the code
458:           For i = 0 To pCVDomain.CodeCount - 1
459:             If pCVDomain.Value(i) = pVal Then
460:               ConvertToDescription = pCVDomain.Name(i) 'Return the code value
461:               GoTo Process_Exit
462:             End If
463:           Next i
464:         Else ' If range domain, return the numeric value
465:           ConvertToDescription = pVal
466:           GoTo Process_Exit
467:         End If
468:       End If  'If pDomain is nothing/Else
469:       ConvertToDescription = pVal
470:     Else
      'Field not found
472:       ConvertToDescription = ""
473:     End If 'If lFld > -1/Else

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "ConvertToDescription " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
 
 

'***************************************************************************
'Name:  CompareAndSaveValue
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Description:   Compare the descriptive value in the GUI to the original descriptive value
'Purpose:       Return an object that indicates the status (changed/unchanged) of this row
'Called From:   modutils.ReadValue, modutils.ConvertCode
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
Public Sub CompareAndSaveValue(pRow As IRow, pFldName As String, vValNew As Variant, pRowChanged As clsRowChanged)
  On Error GoTo ErrorHandler

    Dim vValOrg As Variant
508:     vValOrg = modUtils.ReadValue(pRow, pFldName)
509:     If vValNew <> vValOrg Then
      'Get the Code value that is to be stored in the db
511:       vValNew = modUtils.ConvertCode(pRow, pFldName, vValNew)
      'If the value is changed, update the row
      Dim lFld As Long
514:       lFld = pRow.Fields.FindField(pFldName)
515:       If lFld > -1 Then
        Dim pFldType As esriFieldType
517:         pFldType = pRow.Fields.Field(lFld).Type
518:         If pFldType = esriFieldTypeDouble Then
          Dim dValNew As Double
520:           If IsNumeric(vValNew) Then dValNew = CDbl(vValNew)
521:           If dValNew <> vValOrg Then
522:             pRow.Value(lFld) = dValNew
523:             pRowChanged.RowChanged = True
524:           End If
525:         ElseIf pFldType = esriFieldTypeInteger Or pFldType = esriFieldTypeSmallInteger Then
          Dim iValNew As Long
527:           If IsNumeric(vValNew) Then iValNew = CLng(vValNew)
528:           If iValNew <> vValOrg Then
529:             pRow.Value(lFld) = iValNew
530:             pRowChanged.RowChanged = True
531:           End If
532:         ElseIf pFldType = esriFieldTypeSingle Then
          Dim sValNew As Single
534:           If IsNumeric(vValNew) Then sValNew = CSng(vValNew)
535:           If sValNew <> vValOrg Then
536:             pRow.Value(lFld) = sValNew
537:             pRowChanged.RowChanged = True
538:           End If
539:         ElseIf pFldType = esriFieldTypeDate Then
          Dim dtValNew As Date
541:           If IsDate(vValNew) Then dtValNew = CDate(vValNew)
542:           If dtValNew <> vValOrg Then
543:             pRow.Value(lFld) = dtValNew
544:             pRowChanged.RowChanged = True
545:           End If
546:         ElseIf pFldType = esriFieldTypeString Then
          Dim sgValNew As String
548:           sgValNew = vValNew
549:           If sgValNew <> vValOrg Then
550:             pRow.Value(lFld) = sgValNew
551:             pRowChanged.RowChanged = True
552:           End If
553:         Else
          'Unknown field type
555:         End If
556:      End If
557:   End If

Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "CompareAndSaveValue " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



'***************************************************************************
'Name:  GetValueViaOverlay
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Overlay the passed in feature with a feature class
'Called From: cmdTaxlotAssignment.m_pEditorEvents_OncreateFeature,cmdArrows.GenerateHooks,
'               modutils.ValidateTaxlotNum, modutils.SetAnnosize
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:       Returns the value from the specified field as a variant
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------

'***************************************************************************
Public Function GetValueViaOverlay(pGeom As IGeometry, pOverlayFC As IFeatureClass, sFldName As String) As Variant
  On Error GoTo ErrorHandler

591:   GetValueViaOverlay = ""
592:   If Not pGeom Is Nothing And Not pOverlayFC Is Nothing And Not sFldName = "" Then
    Dim pFeatCur As IFeatureCursor
594:     Set pFeatCur = SpatialQuery(pOverlayFC, pGeom, esriSpatialRelIntersects)
595:     If Not pFeatCur Is Nothing Then
      'Get the first feature.  if more than one, let the user decide
      Dim pFeat As IFeature
598:       Set pFeat = pFeatCur.NextFeature
599:       If Not pFeat Is Nothing Then
        Dim lFld As Long
601:         lFld = pFeat.Fields.FindField(sFldName)
602:         If lFld > -1 Then
          'Get the  value
604:           GetValueViaOverlay = IIf(IsNull(pFeat.Value(lFld)), "", pFeat.Value(lFld))
605:         End If
606:       End If
607:     End If
608:   End If

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "GetValueViaOverlay " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
'***************************************************************************
'Name:  FindControlString
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Purpose:       Find a string in the control.
'Description:   The third argument is the index *after* which to start the search (first item if omitted).
'               If the fourth argument is True it searches for an exact match.
'
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:       The index of the match, or -1 if not found.
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point
'***************************************************************************
Public Function FindControlString(ctrl As Control, ByVal strSearch As String, Optional lStartIdx As Long = -1, Optional ExactMatch As Boolean) As Long
  On Error GoTo ErrorHandler

  Dim uMsg As Long
642:   If TypeOf ctrl Is ListBox Then
643:     uMsg = IIf(ExactMatch, LB_FINDSTRINGEXACT, LB_FINDSTRING)
644:   ElseIf TypeOf ctrl Is ComboBox Then
645:     uMsg = IIf(ExactMatch, CB_FINDSTRINGEXACT, CB_FINDSTRING)
646:   Else
647:     GoTo Process_Exit
648:   End If
649:   FindControlString = SendMessageString(ctrl.hwnd, uMsg, lStartIdx, strSearch)

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "FindControlString " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  SpatialQuery
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Purpose:   Return a feature cursor based on the results of a spatial query

'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:       Returns a search cursor (faster than update)
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006      Initial creation
'***************************************************************************
Public Function SpatialQuery(pFeatureClassIN As esriGeoDatabase.IFeatureClass, _
                             searchGeometry As esrigeometry.IGeometry, _
                             spatialRelation As esriGeoDatabase.esriSpatialRelEnum, _
                             Optional whereClause As String = "" _
                             ) As esriGeoDatabase.IFeatureCursor
  On Error GoTo ErrorHandler

    ' create a spatial query filter
    Dim pSpatialFilter As esriGeoDatabase.ISpatialFilter
687:     Set pSpatialFilter = New esriGeoDatabase.SpatialFilter
    
    ' specify the geometry to query with
690:     Set pSpatialFilter.Geometry = searchGeometry
    
    ' specify what the geometry file is called on the Feature Class that we will be querying against
    Dim strShpFld As String
694:     strShpFld = pFeatureClassIN.ShapeFieldName
695:     pSpatialFilter.GeometryField = strShpFld
    
    'specify the type of spatial operation to use
698:     pSpatialFilter.SpatialRel = spatialRelation

    ' create the where statement
701:     pSpatialFilter.whereClause = whereClause
    
    ' create a cursor that will return the results
    Dim pFeatCursor As esriGeoDatabase.IFeatureCursor
    
    ' perform the query
    Dim pQueryFilter As esriGeoDatabase.IQueryFilter
708:     Set pQueryFilter = pSpatialFilter
709:     Set pFeatCursor = pFeatureClassIN.Search(pQueryFilter, False)
    
711:     Set SpatialQuery = pFeatCursor

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "SpatialQuery " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  SpatialQueryForEdit
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Same as SpatialQuery, but returns an update cursor
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
'
'***************************************************************************
Public Function SpatialQueryForEdit(pFeatureClassIN As esriGeoDatabase.IFeatureClass, _
                             searchGeometry As esrigeometry.IGeometry, _
                             spatialRelation As esriGeoDatabase.esriSpatialRelEnum, _
                             Optional whereClause As String = "" _
                             ) As esriGeoDatabase.IFeatureCursor
  On Error GoTo ErrorHandler

    
    ' create a spatial query filter
    Dim pSpatialFilter As esriGeoDatabase.ISpatialFilter
749:     Set pSpatialFilter = New esriGeoDatabase.SpatialFilter
    
    ' specify the geometry to query with
752:     Set pSpatialFilter.Geometry = searchGeometry
    
    ' specify what the geometry file is called on the Feature Class that we will be querying against
    Dim strShpFld As String
756:     strShpFld = pFeatureClassIN.ShapeFieldName
757:     pSpatialFilter.GeometryField = strShpFld
    
    'specify the type of spatial operation to use
760:     pSpatialFilter.SpatialRel = spatialRelation

    ' create the where statement
763:     pSpatialFilter.whereClause = whereClause
    
    ' create a cursor that will return the results
    Dim pFeatCursor As esriGeoDatabase.IFeatureCursor
    
    ' perform the query
    Dim pQueryFilter As esriGeoDatabase.IQueryFilter
770:     Set pQueryFilter = pSpatialFilter
    'Set pFeatCursor = pFeatureClassIN.Search(pQueryFilter, False)
772:     Set pFeatCursor = pFeatureClassIN.Update(pQueryFilter, False)
    
774:     Set SpatialQueryForEdit = pFeatCursor

  Exit Function
ErrorHandler:
  HandleError True, "SpatialQueryForEdit " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
'***************************************************************************
'Name:  AttributeQuery
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
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Public Function AttributeQuery(pTable As esriGeoDatabase.ITable, _
                               Optional whereClause As String = "" _
                               ) As esriGeoDatabase.ICursor
  On Error GoTo ErrorHandler

'Return a cursor based on an attribute query
' create a query filter
Dim pQueryFilter As esriGeoDatabase.IQueryFilter
809: Set pQueryFilter = New esriGeoDatabase.QueryFilter

' create the where statement
'whereClause = Replace(whereClause, "HYDRO1.", "")
813: pQueryFilter.whereClause = whereClause

' create a cursor that will return the results
Dim pCursor As esriGeoDatabase.ICursor

' query the table passed into the fuction
819: Set pCursor = pTable.Search(pQueryFilter, False)

'Count the number of selected records
Dim selCount As Long
823: selCount = pTable.RowCount(pQueryFilter)
824: If selCount = 0 Then
825:   Set AttributeQuery = Nothing
826: Else
827:   Set AttributeQuery = pCursor
828: End If

Exit Function
ErrorHandler:
  HandleError True, "AttributeQuery " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  GetDomainDefaultValue
'Initial Author:
'Subsequent Author:
'Created:
'Purpose:   Returns the default value if this is a domain field with a default
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
'James Moore    10/11/2006  single exit point using goto
'***************************************************************************
Public Function GetDomainDefaultValue(pTable As ITable, sFldName As String) As Variant
  On Error GoTo ErrorHandler

     Dim lFld As Long
     Dim pField As IField
860:      lFld = pTable.FindField(sFldName)
861:      If lFld > -1 Then
862:         Set pField = pTable.Fields.Field(lFld)
863:      Else
864:         GetDomainDefaultValue = ""
865:         GoTo Process_Exit
866:      End If
     Dim pDomain As IDomain
868:      Set pDomain = pField.Domain
869:       If pDomain Is Nothing Then
870:         GetDomainDefaultValue = ""
871:         GoTo Process_Exit
872:       Else
        'Determine type of domain  -If Coded Value, get the description
874:         If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
876:           Set pCVDomain = pDomain
          Dim lCode As Long
          Dim vDomainVal As Variant
879:           vDomainVal = pField.DefaultValue
          Dim i As Integer
          'Search the domain for the code
882:           For i = 0 To pCVDomain.CodeCount - 1
883:              If pCVDomain.Value(i) = vDomainVal Then
              'return the description
885:               GetDomainDefaultValue = pCVDomain.Name(i)
886:               GoTo Process_Exit
887:             End If
888:           Next i
889:         Else ' If range domain, return the numeric value
890:           GetDomainDefaultValue = pField.DefaultValue
891:           GoTo Process_Exit
892:         End If
893:       End If  'If pDomain is nothing/Else

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "GetDomainDefaultValue " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  GetSelectedFeatures
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:       Return an IFeatureCursor for the selected features
'Called From:
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point
'***************************************************************************
Public Function GetSelectedFeatures(pFLayer As IFeatureLayer) As IFeatureCursor
  On Error GoTo ErrorHandler


  '  exit if not applicable:
926:   If Not TypeOf pFLayer Is IFeatureLayer Then
927:     GoTo Process_Exit
928:   End If
  
  Dim pFSelection As IFeatureSelection
931:   Set pFSelection = pFLayer
  
933:   pFSelection.SelectionSet.Search Nothing, False, GetSelectedFeatures
  
Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "GetSelectedFeatures " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  HasSelectedFeatures
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:       Determines if the feature layer has a selection
'Called From:
'Methods:       Describe any complex details.
'Parameters:
'Returns:       True or False
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point
'***************************************************************************
Public Function HasSelectedFeatures(pFLayer As IFeatureLayer2) As Boolean
  On Error GoTo ErrorHandler
  '
962:   If pFLayer Is Nothing Then GoTo Process_Exit
  
  '  exit if not applicable:
965:   If Not TypeOf pFLayer Is IFeatureLayer Then
966:     GoTo Process_Exit
967:   End If
  
  Dim pFSelection As IFeatureSelection
  
971:   Set pFSelection = pFLayer
  Dim pFeatCur As IFeatureCursor
973:   pFSelection.SelectionSet.Search Nothing, False, pFeatCur
  Dim pFeat As IFeature
975:   If Not pFeatCur Is Nothing Then
976:     Set pFeat = pFeatCur.NextFeature
977:     If Not pFeat Is Nothing Then 'At least one feature selected
978:         Set pFeat = pFeatCur.NextFeature
979:         If Not pFeat Is Nothing Then 'More than one selected
980:             HasSelectedFeatures = False
981:             GoTo Process_Exit
982:         Else
983:             HasSelectedFeatures = True 'Just one selected
984:             GoTo Process_Exit
985:         End If
986:     Else 'nothing selected
987:         HasSelectedFeatures = False
988:         GoTo Process_Exit
989:     End If
990:   End If
  
Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "HasSelectedFeatures " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  ParseOMMapNum
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Return specific ORMAP values from this string as the whole number represents
'           multiple entities
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
'James Moore    10/11/2006  Single exit point
'***************************************************************************
Public Function ParseOMMapNum(sVal As String, sPartName As String) As String
  On Error GoTo ErrorHandler
    
1022:     If Not Len(sVal) = 24 Then
        'MsgBox "ORMAPMapNumber shoud be 24 characters and instead is " & Len(sVal)
1024:         ParseOMMapNum = ""
1025:         GoTo Process_Exit
1026:     End If
    Select Case LCase(sPartName)
        Case "county"
1029:             ParseOMMapNum = ExtractString(sVal, 1, 2)
            'If CLng(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "town"
1032:             ParseOMMapNum = ExtractString(sVal, 3, 4)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "000"
        Case "townpart"
            
1036:             ParseOMMapNum = ExtractString(sVal, 5, 7)
            'If CDbl(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "towndir"
1039:             ParseOMMapNum = ExtractString(sVal, 8, 8)
        Case "range"
1041:             ParseOMMapNum = ExtractString(sVal, 9, 10)
            'If CLng(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "rangepart"
1044:             ParseOMMapNum = ExtractString(sVal, 11, 13)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "000"
        Case "rangedir"
1047:             ParseOMMapNum = ExtractString(sVal, 14, 14)
        Case "section"
1049:             ParseOMMapNum = ExtractString(sVal, 15, 16)
            'If CLng(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "qtr"
1052:             ParseOMMapNum = ExtractString(sVal, 17, 17)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "0"
        Case "qtrqtr"
1055:             ParseOMMapNum = ExtractString(sVal, 18, 18)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "0"
        Case "anomaly"
1058:             ParseOMMapNum = ExtractString(sVal, 19, 20)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "00"
        Case "suffixtype"
1061:              ParseOMMapNum = ExtractString(sVal, 21, 21)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "0"
        Case "suffixnum"
1064:             ParseOMMapNum = ExtractString(sVal, 22, 24)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "000"
        Case Else
            'some handling?
1068:     End Select

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "ParseOMMapNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  FormatOMMapNum
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Return properly formatted part of OM MapNum string
'Called From:
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:   A formatted  OM MapNum string
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point using goto
'***************************************************************************
Public Function FormatOMMapNum(ByRef asVal As String, ByRef asPartName As String) As String
  On Error GoTo ErrorHandler

    'FormatOMMapNum = Replace(sVal, ".", "")
1100:     FormatOMMapNum = asVal
    Select Case LCase(asPartName)
        Case "county"
1103:             If Len(FormatOMMapNum) <> 2 Then
1104:                 FormatOMMapNum = AddLeadingZeros(FormatOMMapNum, 2)
1105:             End If
        Case "town"
1107:             If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "00"
        Case "townpart"
1109:             FormatOMMapNum = Replace(FormatOMMapNum, "0.", ".")
            'If Len(FormatOMMapNum) <> 3 Then FormatOMMapNum = "000"
        Case "towndir"
1112:             If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "N"
        Case "range"
1114:             If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "01"
        Case "rangepart"
1116:             FormatOMMapNum = Replace(FormatOMMapNum, "0.", ".")
            'If Len(FormatOMMapNum) <> 3 Then FormatOMMapNum = "000"
            'If Len(sVal) <> 3 Then FormatOMMapNum = "000"
        Case "rangedir"
1120:             If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "W"
        Case "section"
1122:             If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "00"
        Case "qtr"
1124:             If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "qtrqtr"
1126:             If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "suffixtype"
1128:             If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "suffixnum"
1130:             If Len(FormatOMMapNum) <> 0 And Len(FormatOMMapNum) > 3 Then
1131:                 FormatOMMapNum = "000"
1132:                 GoTo Process_Exit
1133:             ElseIf Len(FormatOMMapNum) = 1 Then
1134:                 FormatOMMapNum = "00" & FormatOMMapNum
1135:                 GoTo Process_Exit
1136:             ElseIf Len(FormatOMMapNum) = 2 Then
1137:                 FormatOMMapNum = "0" & FormatOMMapNum
1138:                 GoTo Process_Exit
1139:             End If
        Case "anomaly"
1141:             If Len(FormatOMMapNum) > 2 Or Len(FormatOMMapNum) = 0 Then
1142:                 FormatOMMapNum = "00"
1143:                 GoTo Process_Exit
1144:             ElseIf Len(FormatOMMapNum) = 2 Then
            
1146:             ElseIf Len(FormatOMMapNum) = 1 Then
1147:                 FormatOMMapNum = "0" & FormatOMMapNum
1148:                 GoTo Process_Exit
1149:             End If
        Case Else
            'Nothing to implement now
1152:     End Select

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "FormatOMMapNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


'***************************************************************************
'Name:  ExtractString
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Use the low and high values to extract the required string.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Private Function ExtractString(sFullString As String, llow As Long, lhigh As Long) As String
  On Error GoTo ErrorHandler

    Dim sVal1 As String
    Dim sVal2 As String
1186:     sVal1 = Right(sFullString, Len(sFullString) - (llow - 1))
1187:     sVal2 = Left(sVal1, (lhigh - llow) + 1)
1188:     ExtractString = sVal2

  Exit Function
ErrorHandler:
  HandleError False, "ExtractString " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  IsTaxlot
'Initial Author:        James Moore
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Determines if this feature is in the Taxlot feature class
'               Used by generic functions to determine what has to be done.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:       True or False
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Public Function IsTaxlot(obj As IObject) As Boolean
  On Error GoTo ErrorHandler
    
    Dim pOC As IObjectClass
    Dim pDS As IDataset
1221:     Set pOC = obj.Class
1222:     Set pDS = pOC
1223:     If LCase(pDS.Name) = LCase(g_pFldnames.FCTaxlot) Then
1224:         IsTaxlot = True
1225:     Else
1226:         IsTaxlot = False
1227:     End If

  Exit Function
ErrorHandler:
  HandleError True, "IsTaxlot " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  IsMapIndex
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Determines if this feature is in the Taxlot feature class
'               Used by generic functions to determine what has to be done.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:       True or False
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Public Function IsMapIndex(obj As IObject) As Boolean
  On Error GoTo ErrorHandler

    Dim pOC As IObjectClass
    Dim pDS As IDataset
1260:     Set pOC = obj.Class
1261:     Set pDS = pOC
1262:     If LCase(pDS.Name) = LCase(g_pFldnames.FCMapIndex) Then
1263:         IsMapIndex = True
1264:     Else
1265:         IsMapIndex = False
1266:     End If


  Exit Function
ErrorHandler:
  HandleError True, "IsMapIndex " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  IsAnno
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Determines if this feature is annotation feature class
'               Used by generic functions to determine what has to be done.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:   True or False
'Errors:    This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Public Function IsAnno(obj As IObject) As Boolean
  On Error GoTo ErrorHandler

1298:     IsAnno = False

    Dim pOC As IObjectClass
    Dim pDS As IDataset
1302:     Set pOC = obj.Class
1303:     Set pDS = pOC
1304:     If TypeOf obj Is IFeature Then
        Dim pFC As IFeatureClass
1306:         Set pFC = pOC
1307:         If pFC.FeatureType = esriFTAnnotation Then IsAnno = True
1308:     End If


  Exit Function
ErrorHandler:
  HandleError True, "IsAnno " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
'***************************************************************************
'Name:  ValidateTaxlotNum
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Ensure that the numeric taxlot number is unique within the current map index
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:    This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point
'***************************************************************************
Public Function ValidateTaxlotNum(sEnteredTLval As String, pGeometry As IGeometry) As Boolean
  On Error GoTo ErrorHandler

    Dim pTaxlotFlayer As IFeatureLayer2
    Dim pTaxlotFclass As IFeatureClass
    Dim pMIFlayer As IFeatureLayer2
    Dim pMIFclass As IFeatureClass
1342:     Set pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
1343:     If pTaxlotFlayer Is Nothing Then
1344:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
1346:         GoTo Process_Exit
1347:     End If
1348:     Set pTaxlotFclass = pTaxlotFlayer.FeatureClass
1349:     Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
1350:     If pMIFlayer Is Nothing Then
1351:         MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
1353:         GoTo Process_Exit
1354:     End If
1355:     Set pMIFclass = pMIFlayer.FeatureClass
    'Get fields needed to populate the form
    Dim lMIOMNum As Long
    Dim lTLOMNum As Long
    Dim lTLTaxlot As Long
    Dim sMIOMval As String
    Dim sTLOMval As String
1362:     lMIOMNum = modUtils.LocateFields(pMIFclass, g_pFldnames.MIORMAPMapNumberFN)
1363:     lTLOMNum = modUtils.LocateFields(pTaxlotFclass, g_pFldnames.TLOrmapMapNumberFN)
1364:     lTLTaxlot = modUtils.LocateFields(pTaxlotFclass, g_pFldnames.TLTaxlotFN)
1365:     sMIOMval = GetValueViaOverlay(pGeometry, pMIFclass, g_pFldnames.MIORMAPMapNumberFN)
    'if no Mapindex or ORMAP mapnum, then no need to continue
1367:     If sMIOMval = "" Then
1368:         ValidateTaxlotNum = True
1369:         GoTo Process_Exit
1370:     End If
    'Make sure this number is unique within taxlots with this OM number
    Dim pCursor As ICursor
    Dim sWhere As String
1374:     sWhere = g_pFldnames.TLOrmapMapNumberFN & " = '" & sMIOMval & _
            "' and " & g_pFldnames.TLTaxlotFN & " = '" & sEnteredTLval & "'"
1376:     Set pCursor = AttributeQuery(pTaxlotFclass, sWhere)
1377:     If Not pCursor Is Nothing Then
        Dim pRow As IRow
1379:         Set pRow = pCursor.NextRow
1380:         If Not pRow Is Nothing Then
1381:             ValidateTaxlotNum = False
1382:         Else
1383:             ValidateTaxlotNum = True
1384:         End If
1385:     Else
1386:         ValidateTaxlotNum = True
1387:     End If

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "ValidateTaxlotNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  CalcTaxlotValues
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Calculates Taxlot vaules from ORMAPMapnum
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
'James Moore    10/11/2006  single exit point using goto
'***************************************************************************
Public Sub CalcTaxlotValues(pFeat As IFeature, pMIFlayer As IFeatureLayer)
  On Error GoTo ErrorHandler

    Dim pTaxlotFclass As IFeatureClass
    Dim pMIFclass As IFeatureClass
    Dim lOMTLNumFld As Long
    Dim lOMNumFld As Long
    Dim lMNumFld As Long
    Dim lTaxlotFld As Long
    Dim lTLCntyFld As Long
    Dim lTLTownFld As Long
    Dim lTLTownPartFld As Long
    Dim lTLTownDirFld As Long
    Dim lTLRangeFld As Long
    Dim lTLRangePartFld As Long
    Dim lTLRangeDirFld As Long
    Dim lTLSectNumFld As Long
    Dim lTLQtrFld As Long
    Dim lTLQQFld As Long
    Dim lTLMapSufTypeFld As Long
    Dim lTLMapSufNumFld As Long
    Dim lTLSpecInterestFld As Long
    Dim lTLMapTaxlotFld As Long
    Dim lTLMapNumberFld As Long
    Dim lTLAnomalyFld As Long
    Dim lTaxlotMapAcres As Long
    Dim lTaxlotShapeArea As Long
    Dim response As Variant
    
1444:     Set pTaxlotFclass = pFeat.Class
    'Find MapIndex
1446:     Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
1447:     If pMIFlayer Is Nothing Then
1448:         response = MsgBox("Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex & ".  " & _
        "Load " & g_pFldnames.FCMapIndex & " automatically?", vbYesNo)
1451:         If response <> vbYes Then GoTo Process_Exit
1452:         modUtils.LoadFCIntoMap g_pFldnames.FCMapIndex, pTaxlotFclass
        'Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
1454:         If pMIFlayer Is Nothing Then GoTo Process_Exit
1455:     End If

    'Find all fields needed
1458:     m_bContinue = True
1459:     lOMTLNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLOrmapTaxlotFN)
1460:     lOMNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLOrmapMapNumberFN)
1461:     lMNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLMapNumberFN)
1462:     lTLCntyFld = LocateFields(pTaxlotFclass, g_pFldnames.TLCountyFN)
1463:     lTaxlotFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTaxlotFN)
1464:     lTLTownFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTownFN)
1465:     lTLTownPartFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTownPartFN)
1466:     lTLTownDirFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTownDirFN)
1467:     lTLRangeFld = LocateFields(pTaxlotFclass, g_pFldnames.TLRangeFN)
1468:     lTLRangePartFld = LocateFields(pTaxlotFclass, g_pFldnames.TLRangePartFN)
1469:     lTLRangeDirFld = LocateFields(pTaxlotFclass, g_pFldnames.TLRangeDirFN)
1470:     lTLSectNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSectNumberFN)
1471:     lTLQtrFld = LocateFields(pTaxlotFclass, g_pFldnames.TLQtrFN)
1472:     lTLQQFld = LocateFields(pTaxlotFclass, g_pFldnames.TLQtrQtrFN)
1473:     lTLMapSufTypeFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSufTypeFN)
1474:     lTLMapSufNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSufNumFN)
1475:     lTLSpecInterestFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSpecInterestFN)
1476:     lTLMapTaxlotFld = LocateFields(pTaxlotFclass, g_pFldnames.TLMapTaxlotFN)
1477:     lTLMapNumberFld = LocateFields(pTaxlotFclass, g_pFldnames.TLMapNumberFN)
1478:     lTaxlotMapAcres = LocateFields(pTaxlotFclass, g_pFldnames.TLMapAcresFN)
1479:     lTLAnomalyFld = LocateFields(pTaxlotFclass, g_pFldnames.TLAnomalyFN)
1480:     If Not m_bContinue Then GoTo Process_Exit 'If any fields not found

    'Obtain the map index poly via overlay
    Dim sExistVal As String
    Dim pArea As IArea
    Dim pCenter As IPoint
    Dim sExistOMMapNum As String
    Dim sExistMapNum As String
    
1489:     Set pArea = pFeat.Shape
1490:     Set pCenter = pArea.Centroid
    
    'Update Acreage
1493:     pFeat.Value(lTaxlotMapAcres) = pArea.Area / 43560  '(pFeat.Value(lTaxlotShapeArea) / 43560)
    'Return the OMMapNum and MapNum and insert values into Taxlot
1495:     sExistOMMapNum = GetValueViaOverlay(pCenter, pMIFlayer.FeatureClass, g_pFldnames.MIORMAPMapNumberFN)
1496:     If sExistOMMapNum = "" Then GoTo Process_Exit 'If no value for whatever reason, don't continue
1497:     sExistMapNum = GetValueViaOverlay(pCenter, pMIFlayer.FeatureClass, g_pFldnames.MIMapNumberFN)
1498:     If sExistMapNum = "" Then GoTo Process_Exit 'If no value for whatever reason, don't continue
    'Store individual components of map number in taxlot
1500:     pFeat.Value(lOMNumFld) = sExistOMMapNum
1501:     pFeat.Value(lMNumFld) = sExistMapNum
    
    'County
1504:     sExistVal = ParseOMMapNum(sExistOMMapNum, "county")
1505:     sExistVal = ConvertCode(pFeat, g_pFldnames.TLCountyFN, sExistVal)
1506:     pFeat.Value(lTLCntyFld) = CInt(sExistVal) 'Store county in county field
    
    'Town
1509:     sExistVal = ParseOMMapNum(sExistOMMapNum, "town")
1510:     pFeat.Value(lTLTownFld) = CInt(sExistVal)

    'TownPart
1513:     sExistVal = ParseOMMapNum(sExistOMMapNum, "townpart")
1514:     pFeat.Value(lTLTownPartFld) = CDbl(sExistVal)

    'TownDir
1517:     sExistVal = ParseOMMapNum(sExistOMMapNum, "towndir")
1518:     pFeat.Value(lTLTownDirFld) = sExistVal

    'Range
1521:     sExistVal = ParseOMMapNum(sExistOMMapNum, "range")
1522:     pFeat.Value(lTLRangeFld) = CInt(sExistVal)

    'RangePart
1525:     sExistVal = ParseOMMapNum(sExistOMMapNum, "rangepart")
1526:     pFeat.Value(lTLRangePartFld) = CDbl(sExistVal)

    'RangeDir
1529:     sExistVal = ParseOMMapNum(sExistOMMapNum, "rangedir")
1530:     pFeat.Value(lTLRangeDirFld) = sExistVal

    'Section
1533:     sExistVal = ParseOMMapNum(sExistOMMapNum, "section")
1534:     pFeat.Value(lTLSectNumFld) = CInt(sExistVal)
 
    'Qtr
1537:     sExistVal = ParseOMMapNum(sExistOMMapNum, "qtr")
1538:     pFeat.Value(lTLQtrFld) = sExistVal
    
    'QtrQtr
1541:     sExistVal = ParseOMMapNum(sExistOMMapNum, "qtrqtr")
1542:     pFeat.Value(lTLQQFld) = sExistVal

    'MapSuffixType
1545:     sExistVal = ParseOMMapNum(sExistOMMapNum, "suffixtype")
1546:     sExistVal = ConvertCode(pFeat, g_pFldnames.TLSufTypeFN, sExistVal)
1547:     pFeat.Value(lTLMapSufTypeFld) = sExistVal
    
    'MapSuffixNum
1550:     sExistVal = ParseOMMapNum(sExistOMMapNum, "suffixnum")
1551:     pFeat.Value(lTLMapSufNumFld) = sExistVal
    
    'Anomaly
1554:     sExistVal = ParseOMMapNum(sExistOMMapNum, "anomaly")
1555:     pFeat.Value(lTLAnomalyFld) = sExistVal
    
    'SpecialInterest
1558:     sExistVal = IIf(IsNull(pFeat.Value(lTLSpecInterestFld)), "00000", pFeat.Value(lTLSpecInterestFld))
1559:     If Len(sExistVal) < 5 Then
1560:      Do Until Len(sExistVal) = 5
1561:         sExistVal = "0" & sExistVal
1562:      Loop
1563:     End If
1564:     pFeat.Value(lTLSpecInterestFld) = sExistVal
    
    'Recalculate OMTaxlot
1567:     If IsNull(pFeat.Value(lTaxlotFld)) Then GoTo Process_Exit
    Dim sTaxlotVal As String
    'Taxlot has actual taxlot number.  ORMAPTaxlot requires a 5-digit number, so leading zeros have to be added
1570:     sTaxlotVal = pFeat.Value(lTaxlotFld)
1571:     sTaxlotVal = AddLeadingZeros(sTaxlotVal, 5)
    Dim sNewOMTLNum As String
    Dim sExistOMTLNum As String
1574:     If IsNull(pFeat.Value(lOMTLNumFld)) Then GoTo Process_Exit
1575:     sExistOMTLNum = pFeat.Value(lOMTLNumFld)
1576:     sNewOMTLNum = CalcOMTLNum(sExistOMTLNum, pFeat, sTaxlotVal)
    'If no changes, don't save value
1578:     If Not sExistOMTLNum = sNewOMTLNum Then
1579:         pFeat.Value(lOMTLNumFld) = sNewOMTLNum
1580:     End If
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "CalcTaxlotValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  AddLeadingZeros
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Add leading zeros if necessary
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
'James Moore    10/11/2006  asCurString was being passed by Reference now passed by value
'***************************************************************************
Public Function AddLeadingZeros(ByVal asCurString As String, ByRef lWidth As Long) As String
  On Error GoTo ErrorHandler

1610:         If Len(asCurString) < lWidth Then
1611:          Do Until Len(asCurString) = lWidth
1612:             asCurString = "0" & asCurString
1613:          Loop
1614:         End If
1615:         AddLeadingZeros = asCurString

  Exit Function
ErrorHandler:
  HandleError True, "AddLeadingZeros " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  GetCentroid
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Determines if this feature is annotation feature class then gets the centroid.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:       A Point object
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Public Function GetCentroid(ByRef pFeat As IFeature) As IPoint
  On Error GoTo ErrorHandler

1645:         If pFeat.FeatureType = esriFTAnnotation Or pFeat.FeatureType = esriFTDimension Then
            Dim pArea As IArea
1647:             Set pArea = pFeat.Shape
1648:             Set GetCentroid = pArea.Centroid
1649:         End If

  Exit Function
ErrorHandler:
  HandleError True, "GetCentroid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  CT_GetCenterOfEnvelope
'Initial Author:
'Subsequent Author:
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:       A Point object
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Function CT_GetCenterOfEnvelope(ByRef pEnv As IEnvelope) As IPoint
  On Error GoTo ErrorHandler

    Dim pCenter As IPoint
1681:     Set pCenter = New Point
1682:     pCenter.X = pEnv.XMin + (pEnv.XMax - pEnv.XMin) / 2
1683:     pCenter.Y = pEnv.YMin + (pEnv.YMax - pEnv.YMin) / 2
1684:     Set CT_GetCenterOfEnvelope = pCenter

  Exit Function
ErrorHandler:
  HandleError True, "CT_GetCenterOfEnvelope " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  GetRelatedObjects
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Using the passed in object, get related features through a relationship class
'               This is optimized for anno because there is a single relationship class.
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
Public Function GetRelatedObjects(pObj As IObject) As IFeature
  On Error GoTo ErrorHandler

    Dim pEnumRelClass As IEnumRelationshipClass
    Dim pRelClass As IRelationshipClass
    Dim pParentSet As esriSystem.ISet
    Dim pParentFeat As IFeature
    
1720:     Set pEnumRelClass = pObj.Class.RelationshipClasses(esriRelRoleAny)
1721:     If Not pEnumRelClass Is Nothing Then
1722:       Set pRelClass = pEnumRelClass.Next
1723:       If Not pRelClass Is Nothing Then
1724:           Set pParentSet = pRelClass.GetObjectsRelatedToObject(pObj)
1725:       End If
1726:     Else
1727:         GoTo Process_Exit
1728:     End If
1729:     If Not pParentSet Is Nothing Then
1730:         Set pParentFeat = pParentSet.Next
1731:         If Not pParentFeat Is Nothing Then
1732:             Set GetRelatedObjects = pParentFeat
1733:         End If
1734:     End If

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "GetRelatedObjects " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  GetAnnoSizeByScale
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
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Using Strcomp function to compare strings
'***************************************************************************
Public Function GetAnnoSizeByScale(sFCName As String, lScale As Long) As Double
  On Error GoTo ErrorHandler

    Dim dSize As Double
    '++ New coded added 10/21/05
1768:     With g_pFldnames
'++START JWM 10/11/2006 using strcomp function
1770:         If StrComp(sFCName, .FCTLAcrAnno, vbTextCompare) = 0 Then
        'Determine anno size based on scale
1772:         If lScale = 120 Then dSize = .AnnoSizeTLAcr120
1773:         If lScale = 240 Then dSize = .AnnoSizeTLAcr240
1774:         If lScale = 360 Then dSize = .AnnoSizeTLAcr360
1775:         If lScale = 480 Then dSize = .AnnoSizeTLAcr480
1776:         If lScale = 600 Then dSize = .AnnoSizeTLAcr600
1777:         If lScale = 1200 Then dSize = .AnnoSizeTLAcr1200
1778:         If lScale = 2400 Then dSize = .AnnoSizeTLAcr2400
1779:         If lScale = 4800 Then dSize = .AnnoSizeTLAcr4800
1780:         If lScale = 9600 Then dSize = .AnnoSizeTLAcr9600
1781:         If lScale = 24000 Then dSize = .AnnoSizeTLAcr24000
'++END JWM 10/11/2006
1783:       ElseIf StrComp(sFCName, .FCTLNumAnno, vbTextCompare) = 0 Then
1784:         If lScale = 120 Then dSize = .AnnoSizeTLNum120
1785:         If lScale = 240 Then dSize = .AnnoSizeTLNum240
1786:         If lScale = 360 Then dSize = .AnnoSizeTLNum360
1787:         If lScale = 480 Then dSize = .AnnoSizeTLNum480
1788:         If lScale = 600 Then dSize = .AnnoSizeTLNum600
1789:         If lScale = 1200 Then dSize = .AnnoSizeTLNum1200
1790:         If lScale = 2400 Then dSize = .AnnoSizeTLNum2400
1791:         If lScale = 4800 Then dSize = .AnnoSizeTLNum4800
1792:         If lScale = 9600 Then dSize = .AnnoSizeTLNum9600
1793:         If lScale = 24000 Then dSize = .AnnoSizeTLNum24000
1794:       Else
        'Something not being trapped
1796:         dSize = 10
1797:       End If
1798:     End With
    '++end new code
    'TODO #####
    'Determine a default
1802:     If dSize = 0 Then dSize = 5
1803:     GetAnnoSizeByScale = dSize

  Exit Function
ErrorHandler:
  HandleError True, "GetAnnoSizeByScale " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  FileExists
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
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
'
'***************************************************************************
Public Function FileExists(sPath As String) As Boolean
  On Error GoTo ErrorHandler


    Dim pFSO As Object
1835:     Set pFSO = CreateObject("Scripting.FileSystemObject")
1836:     If Not pFSO.FileExists(sPath) Then
1837:         FileExists = False
1838:     Else
1839:         FileExists = True
1840:     End If


  Exit Function
ErrorHandler:
  HandleError True, "FileExists " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
'***************************************************************************
'Name:  GetAppRef
'Initial Author:        James Moore
'Subsequent Author:     Type your name here.
'Created:       10/11/2006
'Called From:
'Description:   Used to obtain a reference the the Application, which is used throughout the code
'               This is a more complex process with VB code because the code does not live in the MXD.
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
'James Moore    10/11/2006      Initial creation
'***************************************************************************
Public Function GetAppRef() As IApplication
  On Error GoTo ErrorHandler
  
Dim doc As IDocument
Dim app As IApplication
Dim pMXDoc As IMxDocument
Dim pobjectFactory As IObjectFactory
Dim rot As AppROT
Dim strName As String

1878: Set rot = New AppROT
1879: If rot.Count = 1 Then
1880:     Set app = rot.Item(0) 'ArcCatalog
1881: Else
1882:     Set app = rot.Item(1) 'ArcMap
1883: End If
1884: Set pobjectFactory = app

1886: Set GetAppRef = app


  Exit Function
ErrorHandler:
  HandleError True, "GetAppRef " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  GetMXDocRef
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Get a reference to the current map document
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
'
'***************************************************************************
Public Function GetMXDocRef() As IMxDocument
  On Error GoTo ErrorHandler

Dim doc As IDocument
Dim app As IApplication
Dim pMXDoc As IMxDocument
Dim pobjectFactory As IObjectFactory
Dim rot As AppROT
Dim strName As String

1924: Set rot = New AppROT
1925: If rot.Count = 1 Then
1926:     Set app = rot.Item(0) 'ArcCatalog
1927: Else
1928:     Set app = rot.Item(1) 'ArcMap
1929: End If
1930: Set pobjectFactory = app
1931: Set pMXDoc = app.Document

1933: Set GetMXDocRef = pMXDoc


  Exit Function
ErrorHandler:
  HandleError True, "GetMXDocRef " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  LoadFCIntoMap
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:       Loads a feature class into the current map
'Called From:
'Methods:       Feature class must be in the same feature dataset as pOtherFC.
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
Public Sub LoadFCIntoMap(sFCName As String, pOtherFC As IFeatureClass)
  On Error GoTo ErrorHandler

    Dim pWS As IWorkspace
    Dim pFWS As IFeatureWorkspace
    Dim pFC As IFeatureClass
    Dim pFeatLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim pMXDoc As IMxDocument
    Dim pMap As IMap
1971:     Set pWS = pOtherFC.FeatureDataset.Workspace
1972:     Set pFWS = pWS
1973:     Set pFC = pFWS.OpenFeatureClass(sFCName)
1974:     Set pFeatLayer = New FeatureLayer
1975:     Set pFeatLayer.FeatureClass = pFC
1976:     Set pDataset = pFC
1977:     pFeatLayer.Name = pDataset.Name
1978:     Set pMXDoc = g_pApp.Document
1979:     Set pMap = pMXDoc.FocusMap
1980:     pMap.AddLayer pFeatLayer
1981:     pMXDoc.CurrentContentsView.Refresh 0


  Exit Sub
ErrorHandler:
  HandleError True, "LoadFCIntoMap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  IsOrMapFeature
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Determines if a feature class part of the ORMAP design,
'               If not, it will not be used by any code in this project
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
Public Function IsOrMapFeature(obj As esriGeoDatabase.IObject) As Boolean
  On Error GoTo ErrorHandler

    
    Dim pOC As IObjectClass
    Dim pDSet As IDataset
    Dim pName As String
2017:     Set pOC = obj.Class
2018:     Set pDSet = pOC
2019:     pName = LCase(Trim(pDSet.Name))
2020:     If pName = LCase(Trim(g_pFldnames.FCAnno10)) Or pName = LCase(Trim(g_pFldnames.FCAnno100)) Or pName = LCase(Trim(g_pFldnames.FCAnno20)) Or _
        pName = LCase(Trim(g_pFldnames.FCAnno200)) Or pName = LCase(Trim(g_pFldnames.FCAnno2000)) Or pName = LCase(Trim(g_pFldnames.FCAnno30)) Or _
        pName = LCase(Trim(g_pFldnames.FCAnno40)) Or pName = LCase(Trim(g_pFldnames.FCAnno400)) Or pName = LCase(Trim(g_pFldnames.FCAnno50)) Or _
        pName = LCase(Trim(g_pFldnames.FCAnno800)) Or pName = LCase(Trim(g_pFldnames.FCCartoLines)) Or pName = LCase(Trim(g_pFldnames.FCLotsAnno)) Or _
        pName = LCase(Trim(g_pFldnames.FCMapIndex)) Or pName = LCase(Trim(g_pFldnames.FCPlats)) Or pName = LCase(Trim(g_pFldnames.FCReferenceLines)) Or _
        pName = LCase(Trim(g_pFldnames.FCTaxCode)) Or pName = LCase(Trim(g_pFldnames.FCTaxCode)) Or pName = LCase(Trim(g_pFldnames.FCTaxCodeAnno)) Or _
        pName = LCase(Trim(g_pFldnames.FCTaxlot)) Or pName = LCase(Trim(g_pFldnames.FCTaxlotLines)) Or pName = LCase(Trim(g_pFldnames.FCTLAcrAnno)) Or _
        pName = LCase(Trim(g_pFldnames.FCTLNumAnno)) Then
2028:     Else
2029:         IsOrMapFeature = False
2030:     End If


  Exit Function
ErrorHandler:
  HandleError True, "IsOrMapFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  SetAnnoSize
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   If working with anno, determine what size it should be.
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
Public Sub SetAnnoSize(obj As IObject, pFeat As IFeature)
  On Error GoTo ErrorHandler

    Dim sMapNum As String
    Dim pMIFlayer As IFeatureLayer
    Dim pMIFclass As IFeatureClass
    Dim lAnnoMapNumFld As Long
    Dim u As New UID
    Dim pGeometry As IGeometry
    Dim pEnv As IEnvelope
    Dim pCenter As IPoint
    Dim sMapScale As String
    Dim pAnnotationFeature As IAnnotationFeature
    Dim pAnnotationElement As IAnnotationElement
    Dim pElement As IElement
    Dim pTextElement As ITextElement
    Dim pTextSym As ITextSymbol
    Dim pAnnoDset As IDataset
    Dim pAnnoClass As IObjectClass
    Dim dSize As Double
    
    Dim pAnnoFeat As IFeature
    Dim pAOC As IObjectClass
2081:     Set pAOC = obj.Class
2082:     Set pAnnoFeat = obj
    
    'Capture MapNumber for each anno feature created
2085:     lAnnoMapNumFld = LocateFields(obj.Class, g_pFldnames.MIMapNumberFN)
2086:     If lAnnoMapNumFld = -1 Then GoTo Process_Exit
    

    'If new anno feature with no text, determine if it has a shape
    Dim lFld As Long
2091:     lFld = pAnnoFeat.Fields.FindField("TextString")
2092:     If lFld = -1 Then
2093:         MsgBox "Unable to locate textstring field in anno class.  Cannot set size", vbCritical
2094:         GoTo Process_Exit
2095:     End If
    Dim vVal As Variant
2097:     vVal = pAnnoFeat.Value(lFld)
2098:     If IsNull(vVal) Then GoTo Process_Exit
        
    
2101:     Set pFeat = obj
2102:     Set pGeometry = pFeat.Shape
2103:     If pGeometry.IsEmpty Then GoTo Process_Exit
2104:     Set pEnv = pGeometry.Envelope
2105:     Set pCenter = CT_GetCenterOfEnvelope(pEnv)
2106:     Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
2107:     If pMIFlayer Is Nothing Then GoTo Process_Exit
2108:     Set pMIFclass = pMIFlayer.FeatureClass
2109:     sMapNum = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
    
    'Allow existing anno to be moved without changing MapNumber
    'Some anno will reside in another Taxlot, but labels the neighboring taxlot
2113:     If sMapNum = obj.Value(lAnnoMapNumFld) Then
2114:         obj.Value(lAnnoMapNumFld) = sMapNum
    
        'Update the size to reflect current mapscale
2117:         sMapScale = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
2118:         If IsNull(sMapScale) Then GoTo Process_Exit
        
        'Determine which annotation class this is
2121:         Set pAnnoClass = obj.Class
2122:         Set pAnnoDset = pAnnoClass
        'If other anno, don't continue
2124:         If LCase(pAnnoDset.Name) <> LCase(g_pFldnames.FCTLAcrAnno) And LCase(pAnnoDset.Name) <> LCase(g_pFldnames.FCTLNumAnno) Then
2125:             GoTo Process_Exit
2126:         End If
        
2128:         dSize = modUtils.GetAnnoSizeByScale(pAnnoDset.Name, CLng(sMapScale))
        'Get the anno feature, its symbol, set the appropriate size
2130:         Set pAnnotationFeature = obj
2131:         Set pAnnotationElement = pAnnotationFeature.Annotation
2132:         Set pElement = pAnnotationElement
2133:         Set pTextElement = pElement
2134:         Set pTextSym = pTextElement.Symbol
2135:         pTextSym.Size = dSize
2136:         pTextElement.Symbol = pTextSym
2137:         Set pElement = pTextElement
2138:         Set pAnnotationElement = pElement
2139:         pAnnotationFeature.Annotation = pAnnotationElement
2140:     End If
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "SetAnnoSize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  LocateFields
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Return the index (location) of a field within a feature class.
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
Public Function LocateFields(pFClass As IFeatureClass, pFldName As String) As Long
  On Error GoTo ErrorHandler

    '
    Dim lFld As Long
2172:     lFld = pFClass.Fields.FindField(pFldName)
2173:     If lFld > -1 Then
2174:       LocateFields = lFld
2175:     Else
2176:         MsgBox "Unable to locate " & pFldName & " field in " & _
        pFClass.AliasName & " feature class"
2178:     End If


  Exit Function
ErrorHandler:
  HandleError True, "LocateFields " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  UpdateAutoFields
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:
'Description:   Code to update AutoDate and AutoWho.
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
Public Sub UpdateAutoFields(pFeat As IFeature)
  On Error GoTo ErrorHandler


    Dim lAutoDateFld As Long
    Dim lAutoWhoFld As Long
2212:     lAutoDateFld = pFeat.Fields.FindField(g_pFldnames.AutoDateFN)
2213:     If lAutoDateFld > -1 Then
2214:         pFeat.Value(lAutoDateFld) = Now
2215:     End If
2216:     lAutoWhoFld = pFeat.Fields.FindField(g_pFldnames.AutoWhoFN)
2217:     If lAutoWhoFld > -1 Then
2218:         pFeat.Value(lAutoWhoFld) = Environ("USERNAME")
2219:     End If


  Exit Sub
ErrorHandler:
  HandleError True, "UpdateAutoFields " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  Validate5Digits
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
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006      Initial creation
'***************************************************************************
Public Function Validate5Digits(sString As String)
  On Error GoTo ErrorHandler

            'Make sure taxlot number is 5 characters
2252:             If Len(sString) < 5 Then
2253:              Do Until Len(sString) = 5
2254:                 sString = "0" & sString
2255:              Loop
2256:             End If
2257:             Validate5Digits = sString

  Exit Function
ErrorHandler:
  HandleError True, "Validate5Digits " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  GetSpecialInterests
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
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006      Initial creation
'***************************************************************************
Public Function GetSpecialInterests(pFeature As IFeature) As String
  On Error GoTo ErrorHandler

        Dim lTLSpecInterestFld As Long
        Dim sTLSpecVAl As String
2290:         lTLSpecInterestFld = modUtils.LocateFields(pFeature.Class, g_pFldnames.TLSpecInterestFN)
2291:         If lTLSpecInterestFld = -1 Then
2292:             sTLSpecVAl = "00000"
2293:         Else
2294:             If Not IsNull(pFeature.Value(lTLSpecInterestFld)) Then
2295:                 sTLSpecVAl = pFeature.Value(lTLSpecInterestFld)
2296:             Else
2297:                 sTLSpecVAl = "00000"
2298:             End If
            'Verify that it is 5 digits
2300:             If Len(sTLSpecVAl) < 5 Then
2301:              Do Until Len(sTLSpecVAl) = 5
2302:                 sTLSpecVAl = "0" & sTLSpecVAl
2303:              Loop
2304:             End If
2305:         End If
2306:         GetSpecialInterests = sTLSpecVAl

  Exit Function
ErrorHandler:
  HandleError True, "GetSpecialInterests " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  GetMapSufType
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
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006      Initial creation
'***************************************************************************
Public Function GetMapSufType(pFeature As IFeature) As String
  On Error GoTo ErrorHandler

        Dim lTLMapSufTypeFld As Long
        Dim sTLMapSufTypeVAl As String
2339:         lTLMapSufTypeFld = modUtils.LocateFields(pFeature.Class, g_pFldnames.TLSufTypeFN)
2340:         If lTLMapSufTypeFld = -1 Then
2341:             sTLMapSufTypeVAl = "0"
2342:         Else
2343:             If Not IsNull(pFeature.Value(lTLMapSufTypeFld)) Then
2344:                 sTLMapSufTypeVAl = pFeature.Value(lTLMapSufTypeFld)
2345:             Else
2346:                 sTLMapSufTypeVAl = "0"
2347:             End If
                'Verify that it is 1 digit
2349:                 If Len(sTLMapSufTypeVAl) < 1 Then
2350:                     Do Until Len(sTLMapSufTypeVAl) = 1
2351:                        sTLMapSufTypeVAl = "0" & sTLMapSufTypeVAl
2352:                     Loop
2353:                 End If

                'Verify that it isn't more than 1 digit
2356:                 If Len(sTLMapSufTypeVAl) > 1 Then
2357:                     sTLMapSufTypeVAl = Left(sTLMapSufTypeVAl, 1)
2358:                 End If
2359:             End If


2362:         GetMapSufType = sTLMapSufTypeVAl

  Exit Function
ErrorHandler:
  HandleError True, "GetMapSufType " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
'++ END, Laura Gordon, November 29, 2005

'***************************************************************************
'Name:  GetMapSufNum
'Initial Author:        Laura Gordon
'Subsequent Author:     Type your name here.
'Created:       November 29, 2005
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Public Function GetMapSufNum(pFeature As IFeature) As String
  On Error GoTo ErrorHandler

        Dim lTLMapSufNumFld As Long
        Dim sTLMapSufNumVAl As String
2396:         lTLMapSufNumFld = modUtils.LocateFields(pFeature.Class, g_pFldnames.TLSufNumFN)
2397:         If lTLMapSufNumFld = -1 Then
2398:             sTLMapSufNumVAl = "000"
2399:         Else
2400:             If Not IsNull(pFeature.Value(lTLMapSufNumFld)) Then
2401:                 sTLMapSufNumVAl = pFeature.Value(lTLMapSufNumFld)
2402:             Else
2403:                 sTLMapSufNumVAl = "000"
2404:             End If
                'Verify that it is 3 digit
2406:                 If Len(sTLMapSufNumVAl) < 3 Then
2407:                     Do Until Len(sTLMapSufNumVAl) = 3
2408:                        sTLMapSufNumVAl = "0" & sTLMapSufNumVAl
2409:                     Loop
2410:                 End If

                'Verify that it isn't more than 3 digits
2413:                 If Len(sTLMapSufNumVAl) > 3 Then
2414:                     sTLMapSufNumVAl = Left(sTLMapSufNumVAl, 3)
2415:                 End If
2416:             End If


2419:         GetMapSufNum = sTLMapSufNumVAl

  Exit Function
ErrorHandler:
  HandleError True, "GetMapSufNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  CalcOMTLNum
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Calculate ORMAPtaxlot because one if its components may have changed
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
'
'***************************************************************************
Public Function CalcOMTLNum(ByRef sExistOMNum As String, ByRef pFeat As IFeature, ByRef sTLVal As String) As String
  On Error GoTo ErrorHandler

        Dim sShortOMNum As String 'Remove suffixTYpe and suffixNum
        '++ BEGIN, Laura Gordon, November 29, 2005
        '+Dim sTLSpecVAl As String
        '++ END, Laura Gordon, November 29, 2005
        Dim sOMTLNval As String
        '++ BEGIN, Laura Gordon, November 29, 2005
        Dim sTLMapSufTypeVAl As String
        Dim sTLMapSufNumVAl As String
        '++ END, Laura Gordon, November 29, 2005

2459:         sShortOMNum = ShortenOMMapNum(sExistOMNum)
              '++ BEGIN, Laura Gordon, Novemeber 29, 2005
              '+sTLSpecVAl = GetSpecialInterests(pFeat)
              '+sOMTLNval = sShortOMNum & sTLVal & sTLSpecVAl
2463:               sTLMapSufTypeVAl = GetMapSufType(pFeat)
2464:               sTLMapSufNumVAl = GetMapSufNum(pFeat)
2465:               sOMTLNval = sShortOMNum & sTLMapSufTypeVAl & sTLMapSufNumVAl & sTLVal
              '++ END, Laura Gordon, Novemeber 29, 2005
2467:         CalcOMTLNum = sOMTLNval

  Exit Function
ErrorHandler:
  HandleError True, "CalcOMTLNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function ShortenOMMapNum(sOMVal As String) As String
  On Error GoTo ErrorHandler

    'Remove two values from the ORMAPMap number for the purpose of populating ORMAPTaxlog
2478:     ShortenOMMapNum = Left(sOMVal, 20)

  Exit Function
ErrorHandler:
  HandleError True, "ShortenOMMapNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub ZoomToExtent(pEnv As IEnvelope, pMXDoc As IMxDocument)
    'Zooms the current extent to the passed in envelope (i.e. zoom to feature)
    'Works for Layout and Data view
    Dim pMap As IMap
    Dim pActiveView As IActiveView
2490:     Set pMap = pMXDoc.FocusMap
2491:     Set pActiveView = pMap

2493:     pActiveView.Extent = pEnv
2494:     pActiveView.Refresh
End Sub

