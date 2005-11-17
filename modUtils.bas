Attribute VB_Name = "modUtils"
'GENERAL UTILITY MODULE
'MOST COMMONLY USED PROCEDURES ARE LOCATED HERE

Option Explicit
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long
Private m_bContinue As Boolean
Const LB_FINDSTRING = &H18F
Const LB_FINDSTRINGEXACT = &H1A2
Const CB_FINDSTRING = &H14C
Const CB_FINDSTRINGEXACT = &H158
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "C:\active\ModelingWorkshop_01-05-05\CustomCode\ormap\modUtils.bas"


Public Function FindFeatureLayerByDS(DatasetName As String) As IFeatureLayer
  On Error GoTo ErrorHandler

  
    'Return the Feature Layer based on its dataset name
    'This is an easy way to locate a feature layerr in the TOC
    Dim pMXDoc As IMxDocument
    Dim pMap As IMap
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim i As Integer
    
31:     Set pMXDoc = g_pApp.Document
32:     Set pMap = pMXDoc.FocusMap
  
34:     With pMap
35:         For i = 0 To .LayerCount - 1
36:             If TypeOf .Layer(i) Is IFeatureLayer Then
37:                 Set pFeatureLayer = .Layer(i)
38:                 Set pDataset = pFeatureLayer.FeatureClass
39:                 If Not pDataset Is Nothing Then
40:                     If UCase(pDataset.Name) = UCase(DatasetName) Then
41:                         Set FindFeatureLayerByDS = pFeatureLayer
42:                         Exit For
43:                     End If
44:                 End If
45:             End If
46:         Next i
47:     End With
  
49:     If pFeatureLayer Is Nothing Then

51:     End If
  
    Exit Function


  Exit Function
ErrorHandler:
  HandleError True, "FindFeatureLayerByDS " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Public Function GetFWorkspace(pObj As esriGeoDatabase.IObject) As IFeatureWorkspace
  On Error GoTo ErrorHandler


  Dim pFWS As IFeatureWorkspace
  Dim pObjClass As IObjectClass
  Dim pDataset As IDataset
69:   Set pObjClass = pObj.Class
70:   Set pDataset = pObjClass
71:   Set pFWS = pDataset.Workspace
72:   Set GetFWorkspace = pFWS


  Exit Function
ErrorHandler:
  HandleError True, "GetFWorkspace " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function ReadValue(pRow As IRow, pFldName As String, Optional pDataType As String) As Variant
  On Error GoTo ErrorHandler


    'Reads a value from a row, given a field name
    'If a domain field, the descriptive value is returned instead of the stored code
    Dim sVal As String
    Dim lFld As Long
88:     lFld = pRow.Fields.FindField(pFldName)
89:     If lFld > -1 Then
90:       If pDataType = "date" Then
        'If a date and value is null, return a default date value
        '??? How should this be treated?
        Dim pDate As Date
94:         sVal = IIf(IsNull(pRow.Value(lFld)), pDate, pRow.Value(lFld))
95:       Else
96:         sVal = IIf(IsNull(pRow.Value(lFld)), "", pRow.Value(lFld))
97:       End If
      'Determine if domain field
      Dim pField As IField
100:       Set pField = pRow.Fields.Field(lFld)
      Dim pDomain As IDomain
102:       Set pDomain = pField.Domain
103:       If pDomain Is Nothing Then
104:         ReadValue = sVal
        Exit Function
106:       Else
        'Determine type of domain  -If Coded Value, get the description
108:         If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
110:           Set pCVDomain = pDomain
          Dim lCode As Long
          Dim vDomainVal As Variant
113:           vDomainVal = pRow.Value(lFld)
          Dim i As Integer
          'Search the domain for the code
116:           For i = 0 To pCVDomain.CodeCount - 1
117:              If pCVDomain.Value(i) = vDomainVal Then
              'return the description
119:               ReadValue = pCVDomain.Name(i)
              Exit Function
121:             End If
122:           Next i
123:         Else ' If range domain, return the numeric value
124:           ReadValue = sVal
          Exit Function
126:         End If
127:       End If  'If pDomain is nothing/Else
128:       ReadValue = sVal
129:     Else
      'Field not found
131:       ReadValue = ""
132:     End If 'If lFld > -1/Else


  Exit Function
ErrorHandler:
  HandleError True, "ReadValue " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Public Function AddCodesToCmb(pFldName As String, _
                              pFields As IFields, _
                              cboValues As ComboBox, _
                              curVal As Variant, _
                              Optional blnAllowSpace As Boolean) As Boolean
  On Error GoTo ErrorHandler

    'Add the descriptive values from each domain to the drop down comboboxes
149:     If IsMissing(blnAllowSpace) Then blnAllowSpace = True
  
   'Get the Coded Value Domain from the field
      Dim lFld As Long
153:       lFld = pFields.FindField(pFldName)
154:       If lFld > -1 Then
        Dim pField As IField
156:         Set pField = pFields.Field(lFld)
        Dim pDomain As IDomain
158:         Set pDomain = pField.Domain
159:         If pDomain Is Nothing Then
160:           AddCodesToCmb = False
          Exit Function
162:         Else
          'Determine type of domain  -If Coded Value, get the description
164:           If TypeOf pDomain Is ICodedValueDomain Then
            Dim pCVDomain As ICodedValueDomain
166:             Set pCVDomain = pDomain
            ' +++ Get a count of the coded values
            Dim lCodes As Long
            Dim i As Long
170:             lCodes = pCVDomain.CodeCount
            Dim sVal As Variant
            ' +++ Loop through the list of values and add them
            ' +++ and their names to the combo box
174:             If Not blnAllowSpace Then
175:               With cboValues
176:               If .ListCount > 0 Then
177:                 If (.List(0) = "") Or (.List(0) = "") Then
178:                   .RemoveItem (0)
179:                 End If
180:               End If
181:               If .ListCount > 0 Then
182:                 If (.List(.ListCount - 1) = "") Or (.List(.ListCount - 1) = "") Then
183:                   .RemoveItem (.ListCount - 1)
184:                 End If
185:               End If
186:               End With
187:             End If
188:             For i = 0 To lCodes - 1
              'Commented line adds codes and description
              'cboValues.AddItem pCVDomain.Value(i) & ": " & pCVDomain.Name(i)
191:               cboValues.AddItem pCVDomain.Name(i)
192:             Next i
            'Successful completion of addition
            'If current value is null, add an empty string and make it active
195:             If curVal = "" Then
196:               If blnAllowSpace Then
197:                 cboValues.AddItem ""
198:                 cboValues.ListIndex = FindControlString(cboValues, "", 0, True)
                'cboValues.Text = ""
200:               Else
201:                 cboValues.ListIndex = 0
202:               End If
203:             Else 'Otherwise, select the existing value from the list
204:               cboValues.ListIndex = FindControlString(cboValues, curVal, 0, True)
205:             End If
            
207:             AddCodesToCmb = True
208:           Else
            'if Range Domain, do not add values
210:             AddCodesToCmb = False
211:           End If
212:         End If 'if a valid domain
213:       Else 'Field not found
214:         AddCodesToCmb = False
215:       End If


  Exit Function
ErrorHandler:
  HandleError True, "AddCodesToCmb " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'Public Function AddCodesToCmb_org(pFldName As String, _
'                              pFields As IFields, _
'                              cboValues As ComboBox, _
'                              curVal As Variant, _
'                              Optional blnAllowSpace As Boolean) As Boolean
'  On Error GoTo ErrorHandler
'
'  'Changed the " " space to ""
'
'      If IsMissing(blnAllowSpace) Then blnAllowSpace = True
'
'   'Get the Coded Value Domain from the field
'      Dim lFld As Long
'      lFld = pFields.FindField(pFldName)
'      If lFld > -1 Then
'        Dim pField As IField
'        Set pField = pFields.Field(lFld)
'        Dim pDomain As IDomain
'        Set pDomain = pField.Domain
'        If pDomain Is Nothing Then
'          AddCodesToCmb_org = False
'          Exit Function
'        Else
'          'Determine type of domain  -If Coded Value, get the description
'          If TypeOf pDomain Is ICodedValueDomain Then
'            Dim pCVDomain As ICodedValueDomain
'            Set pCVDomain = pDomain
'            ' +++ Get a count of the coded values
'            Dim lCodes As Long
'            Dim i As Long
'            lCodes = pCVDomain.CodeCount
'            Dim sVal As Variant
'            ' +++ Loop through the list of values and add them
'            ' +++ and their names to the combo box
'            If Not blnAllowSpace Then
'              With cboValues
'              If .ListCount > 0 Then
'                If (.List(0) = " ") Or (.List(0) = "") Then
'                  .RemoveItem (0)
'                End If
'              End If
'              If .ListCount > 0 Then
'                If (.List(.ListCount - 1) = " ") Or (.List(.ListCount - 1) = "") Then
'                  .RemoveItem (.ListCount - 1)
'                End If
'              End If
'              End With
'            End If
'            For i = 0 To lCodes - 1
'              'Commented line adds codes and description
'              'cboValues.AddItem pCVDomain.Value(i) & ": " & pCVDomain.Name(i)
'              cboValues.AddItem pCVDomain.Name(i)
'            Next i
'            'Successful completion of addition
'            'If current value is null, add an empty string and make it active
'            If curVal = "" Then
'              If blnAllowSpace Then
'                cboValues.AddItem " "
'                cboValues.ListIndex = FindControlString(cboValues, " ", 0, True)
'                'cboValues.Text = ""
'              Else
'                cboValues.ListIndex = 0
'              End If
'            Else 'Otherwise, select the existing value from the list
'              cboValues.ListIndex = FindControlString(cboValues, curVal, 0, True)
'            End If
'
'            AddCodesToCmb_org = True
'          Else
'            'if Range Domain, do not add values
'            AddCodesToCmb_org = False
'          End If
'        End If 'if a valid domain
'      Else 'Field not found
'        AddCodesToCmb_org = False
'      End If
'
'
'  Exit Function
'ErrorHandler:
'  HandleError True, "AddCodesToCmb_org " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
'End Function
Public Function ConvertCode(pRow As IRow, pFldName As String, pVal As Variant) As Variant
  On Error GoTo ErrorHandler


    'Converts a domain descriptive value to the stored code
    'Domain values chosen from combo boxes must be converted to the code
    'before being stored
    Dim lFld As Long
313:     lFld = pRow.Fields.FindField(pFldName)
314:     If lFld > -1 Then
      'Determine if domain field
      Dim pField As IField
317:       Set pField = pRow.Fields.Field(lFld)
      Dim pDomain As IDomain
319:       Set pDomain = pField.Domain
320:       If pDomain Is Nothing Then
321:         ConvertCode = pVal
        Exit Function
323:       Else
        'Determine type of domain  -If Coded Value, get the description
325:         If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
327:           Set pCVDomain = pDomain
          Dim lCode As Long
          Dim i As Integer
          'Given the description, search the domain for the code
331:           For i = 0 To pCVDomain.CodeCount - 1
332:             If pCVDomain.Name(i) = pVal Then
333:               ConvertCode = pCVDomain.Value(i) 'Return the code value
              Exit Function
335:             End If
336:           Next i
337:         Else ' If range domain, return the numeric value
338:           ConvertCode = pVal
          Exit Function
340:         End If
341:       End If  'If pDomain is nothing/Else
342:       ConvertCode = pVal
343:     Else
      'Field not found
345:       ConvertCode = ""
346:     End If 'If lFld > -1/Else


  Exit Function
ErrorHandler:
  HandleError True, "ConvertCode " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
 
Public Function ConvertToDescription(pFlds As IFields, pFldName As String, pVal As Variant) As Variant
  On Error GoTo ErrorHandler


    'Converts a domain descriptive value to the stored code
    'Domain values chosen from combo boxes must be converted to the code
    'before being stored
    Dim lFld As Long
362:     lFld = pFlds.FindField(pFldName)
363:     If lFld > -1 Then
      'Determine if domain field
      Dim pField As IField
366:       Set pField = pFlds.Field(lFld)
      Dim pDomain As IDomain
368:       Set pDomain = pField.Domain
369:       If pDomain Is Nothing Then
370:         ConvertToDescription = pVal
        Exit Function
372:       Else
        'Determine type of domain  -If Coded Value, get the description
374:         If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
376:           Set pCVDomain = pDomain
          Dim lCode As Long
          Dim i As Integer
          'Given the description, search the domain for the code
380:           For i = 0 To pCVDomain.CodeCount - 1
381:             If pCVDomain.Value(i) = pVal Then
382:               ConvertToDescription = pCVDomain.Name(i) 'Return the code value
              Exit Function
384:             End If
385:           Next i
386:         Else ' If range domain, return the numeric value
387:           ConvertToDescription = pVal
          Exit Function
389:         End If
390:       End If  'If pDomain is nothing/Else
391:       ConvertToDescription = pVal
392:     Else
      'Field not found
394:       ConvertToDescription = ""
395:     End If 'If lFld > -1/Else


  Exit Function
ErrorHandler:
  HandleError True, "ConvertToDescription " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
 
 


Public Sub CompareAndSaveValue(pRow As IRow, pFldName As String, vValNew As Variant, pRowChanged As clsRowChanged)
  On Error GoTo ErrorHandler

    'Compare the descriptive value in the GUI to the original descriptive value
    'Return an object that indicates the status (changed/unchanged) of this row
    Dim vValOrg As Variant
412:     vValOrg = modUtils.ReadValue(pRow, pFldName)
413:     If vValNew <> vValOrg Then
      'Get the Code value that is to be stored in the db
415:       vValNew = modUtils.ConvertCode(pRow, pFldName, vValNew)
      'If the value is changed, update the row
      Dim lFld As Long
418:       lFld = pRow.Fields.FindField(pFldName)
419:       If lFld > -1 Then
        Dim pFldType As esriFieldType
421:         pFldType = pRow.Fields.Field(lFld).Type
422:         If pFldType = esriFieldTypeDouble Then
          Dim dValNew As Double
424:           If IsNumeric(vValNew) Then dValNew = CDbl(vValNew)
425:           If dValNew <> vValOrg Then
426:             pRow.Value(lFld) = dValNew
427:             pRowChanged.RowChanged = True
428:           End If
429:         ElseIf pFldType = esriFieldTypeInteger Or pFldType = esriFieldTypeSmallInteger Then
          Dim iValNew As Long
431:           If IsNumeric(vValNew) Then iValNew = CLng(vValNew)
432:           If iValNew <> vValOrg Then
433:             pRow.Value(lFld) = iValNew
434:             pRowChanged.RowChanged = True
435:           End If
436:         ElseIf pFldType = esriFieldTypeSingle Then
          Dim sValNew As Single
438:           If IsNumeric(vValNew) Then sValNew = CSng(vValNew)
439:           If sValNew <> vValOrg Then
440:             pRow.Value(lFld) = sValNew
441:             pRowChanged.RowChanged = True
442:           End If
443:         ElseIf pFldType = esriFieldTypeDate Then
          Dim dtValNew As Date
445:           If IsDate(vValNew) Then dtValNew = CDate(vValNew)
446:           If dtValNew <> vValOrg Then
447:             pRow.Value(lFld) = dtValNew
448:             pRowChanged.RowChanged = True
449:           End If
450:         ElseIf pFldType = esriFieldTypeString Then
          Dim sgValNew As String
452:           sgValNew = vValNew
453:           If sgValNew <> vValOrg Then
454:             pRow.Value(lFld) = sgValNew
455:             pRowChanged.RowChanged = True
456:           End If
457:         Else
          'Unknown field type
459:         End If
460:      End If
461:   End If


  Exit Sub
ErrorHandler:
  HandleError True, "CompareAndSaveValue " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Public Function GetValueViaOverlay(pGeom As IGeometry, pOverlayFC As IFeatureClass, sFldName As String) As Variant
  On Error GoTo ErrorHandler

  'Overlay the passed in feature with a feature class and return the value from the specified field
475:   GetValueViaOverlay = ""
476:   If Not pGeom Is Nothing And Not pOverlayFC Is Nothing And Not sFldName = "" Then
    Dim pFeatCur As IFeatureCursor
478:     Set pFeatCur = SpatialQuery(pOverlayFC, pGeom, esriSpatialRelIntersects)
479:     If Not pFeatCur Is Nothing Then
      'Get the first feature.  if more than one, let the user decide
      Dim pFeat As IFeature
482:       Set pFeat = pFeatCur.NextFeature
483:       If Not pFeat Is Nothing Then
        Dim lFld As Long
485:         lFld = pFeat.Fields.FindField(sFldName)
486:         If lFld > -1 Then
          'Get the  value
488:           GetValueViaOverlay = IIf(IsNull(pFeat.Value(lFld)), "", pFeat.Value(lFld))
489:         End If
490:       End If
491:     End If
492:   End If


  Exit Function
ErrorHandler:
  HandleError True, "GetValueViaOverlay " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


' Find a string in the control.
' The third argument is the index *after* which to start the search (first item if omitted).
' If the fourth argument is True it searches for an exact match.
' Returns the index of the match, or -1 if not found.

Public Function FindControlString(ctrl As Control, ByVal strSearch As String, Optional lStartIdx As Long = -1, Optional ExactMatch As Boolean) As Long
  On Error GoTo ErrorHandler


  Dim uMsg As Long
511:   If TypeOf ctrl Is ListBox Then
512:     uMsg = IIf(ExactMatch, LB_FINDSTRINGEXACT, LB_FINDSTRING)
513:   ElseIf TypeOf ctrl Is ComboBox Then
514:     uMsg = IIf(ExactMatch, CB_FINDSTRINGEXACT, CB_FINDSTRING)
515:   Else
    Exit Function
517:   End If
518:   FindControlString = SendMessageString(ctrl.hwnd, uMsg, lStartIdx, strSearch)


  Exit Function
ErrorHandler:
  HandleError True, "FindControlString " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function SpatialQuery(pFeatureClassIN As esriGeoDatabase.IFeatureClass, _
                             searchGeometry As esrigeometry.IGeometry, _
                             spatialRelation As esriGeoDatabase.esriSpatialRelEnum, _
                             Optional whereClause As String = "" _
                             ) As esriGeoDatabase.IFeatureCursor
  On Error GoTo ErrorHandler

    'Return a feature cursor based on the results of a spatial query
    'Returns a search cursor (faster than update)
    ' create a spatial query filter
    Dim pSpatialFilter As esriGeoDatabase.ISpatialFilter
537:     Set pSpatialFilter = New esriGeoDatabase.SpatialFilter
    
    ' specify the geometry to query with
540:     Set pSpatialFilter.Geometry = searchGeometry
    
    ' specify what the geometry file is called on the Feature Class that we will be querying against
    Dim strShpFld As String
544:     strShpFld = pFeatureClassIN.ShapeFieldName
545:     pSpatialFilter.GeometryField = strShpFld
    
    'specify the type of spatial operation to use
548:     pSpatialFilter.SpatialRel = spatialRelation

    ' create the where statement
551:     pSpatialFilter.whereClause = whereClause
    
    ' create a cursor that will return the results
    Dim pFeatCursor As esriGeoDatabase.IFeatureCursor
    
    ' perform the query
    Dim pQueryFilter As esriGeoDatabase.IQueryFilter
558:     Set pQueryFilter = pSpatialFilter
559:     Set pFeatCursor = pFeatureClassIN.Search(pQueryFilter, False)
    
561:     Set SpatialQuery = pFeatCursor


  Exit Function
ErrorHandler:
  HandleError True, "SpatialQuery " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Public Function SpatialQueryForEdit(pFeatureClassIN As esriGeoDatabase.IFeatureClass, _
                             searchGeometry As esrigeometry.IGeometry, _
                             spatialRelation As esriGeoDatabase.esriSpatialRelEnum, _
                             Optional whereClause As String = "" _
                             ) As esriGeoDatabase.IFeatureCursor
  On Error GoTo ErrorHandler

    'Same as SpatialQuery, but returns an update cursor
    ' create a spatial query filter
    Dim pSpatialFilter As esriGeoDatabase.ISpatialFilter
580:     Set pSpatialFilter = New esriGeoDatabase.SpatialFilter
    
    ' specify the geometry to query with
583:     Set pSpatialFilter.Geometry = searchGeometry
    
    ' specify what the geometry file is called on the Feature Class that we will be querying against
    Dim strShpFld As String
587:     strShpFld = pFeatureClassIN.ShapeFieldName
588:     pSpatialFilter.GeometryField = strShpFld
    
    'specify the type of spatial operation to use
591:     pSpatialFilter.SpatialRel = spatialRelation

    ' create the where statement
594:     pSpatialFilter.whereClause = whereClause
    
    ' create a cursor that will return the results
    Dim pFeatCursor As esriGeoDatabase.IFeatureCursor
    
    ' perform the query
    Dim pQueryFilter As esriGeoDatabase.IQueryFilter
601:     Set pQueryFilter = pSpatialFilter
    'Set pFeatCursor = pFeatureClassIN.Search(pQueryFilter, False)
603:     Set pFeatCursor = pFeatureClassIN.Update(pQueryFilter, False)
    
605:     Set SpatialQueryForEdit = pFeatCursor


  Exit Function
ErrorHandler:
  HandleError True, "SpatialQueryForEdit " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
Public Function AttributeQuery(pTable As esriGeoDatabase.ITable, _
                               Optional whereClause As String = "" _
                               ) As esriGeoDatabase.ICursor
  On Error GoTo ErrorHandler

'Return a cursor based on an attribute query
' create a query filter
Dim pQueryFilter As esriGeoDatabase.IQueryFilter
620: Set pQueryFilter = New esriGeoDatabase.QueryFilter

' create the where statement
'whereClause = Replace(whereClause, "HYDRO1.", "")
624: pQueryFilter.whereClause = whereClause

' create a cursor that will return the results
Dim pCursor As esriGeoDatabase.ICursor

' query the table passed into the fuction
630: Set pCursor = pTable.Search(pQueryFilter, False)

'Count the number of selected records
Dim selCount As Long
634: selCount = pTable.RowCount(pQueryFilter)
635: If selCount = 0 Then
636:   Set AttributeQuery = Nothing
637: Else
638:   Set AttributeQuery = pCursor
639: End If


  Exit Function
ErrorHandler:
  HandleError True, "AttributeQuery " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Public Function GetDomainDefaultValue(pTable As ITable, sFldName As String) As Variant
  On Error GoTo ErrorHandler

     'Returns the default value if this is a domain field with a default
     Dim lFld As Long
     Dim pField As IField
654:      lFld = pTable.FindField(sFldName)
655:      If lFld > -1 Then
656:         Set pField = pTable.Fields.Field(lFld)
657:      Else
658:         GetDomainDefaultValue = ""
        Exit Function
660:      End If
     Dim pDomain As IDomain
662:      Set pDomain = pField.Domain
663:       If pDomain Is Nothing Then
664:         GetDomainDefaultValue = ""
        Exit Function
666:       Else
        'Determine type of domain  -If Coded Value, get the description
668:         If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
670:           Set pCVDomain = pDomain
          Dim lCode As Long
          Dim vDomainVal As Variant
673:           vDomainVal = pField.DefaultValue
          Dim i As Integer
          'Search the domain for the code
676:           For i = 0 To pCVDomain.CodeCount - 1
677:              If pCVDomain.Value(i) = vDomainVal Then
              'return the description
679:               GetDomainDefaultValue = pCVDomain.Name(i)
              Exit Function
681:             End If
682:           Next i
683:         Else ' If range domain, return the numeric value
684:           GetDomainDefaultValue = pField.DefaultValue
          Exit Function
686:         End If
687:       End If  'If pDomain is nothing/Else


  Exit Function
ErrorHandler:
  HandleError True, "GetDomainDefaultValue " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function GetSelectedFeatures(pFLayer As IFeatureLayer) As IFeatureCursor
  On Error GoTo ErrorHandler


  'return an IFeatureCursor for the selected features
  
  '  exit if not applicable:
702:   If Not TypeOf pFLayer Is IFeatureLayer Then
    Exit Function
704:   End If
  
  Dim pFSelection As IFeatureSelection
707:   Set pFSelection = pFLayer
  
709:   pFSelection.SelectionSet.Search Nothing, False, GetSelectedFeatures
  
  Exit Function


  Exit Function
ErrorHandler:
  HandleError True, "GetSelectedFeatures " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function HasSelectedFeatures(pFLayer As IFeatureLayer2) As Boolean
  On Error GoTo ErrorHandler

  'Determines if the feature layer has a selection
  If pFLayer Is Nothing Then Exit Function
  
  '  exit if not applicable:
726:   If Not TypeOf pFLayer Is IFeatureLayer Then
    Exit Function
728:   End If
  
  Dim pFSelection As IFeatureSelection
  
732:   Set pFSelection = pFLayer
  Dim pFeatCur As IFeatureCursor
734:   pFSelection.SelectionSet.Search Nothing, False, pFeatCur
  Dim pFeat As IFeature
736:   If Not pFeatCur Is Nothing Then
737:     Set pFeat = pFeatCur.NextFeature
738:     If Not pFeat Is Nothing Then 'At least one feature selected
739:         Set pFeat = pFeatCur.NextFeature
740:         If Not pFeat Is Nothing Then 'More than one selected
741:             HasSelectedFeatures = False
            Exit Function
743:         Else
744:             HasSelectedFeatures = True 'Just one selected
            Exit Function
746:         End If
747:     Else 'nothing selected
748:         HasSelectedFeatures = False
        Exit Function
750:     End If
751:   End If
  
  Exit Function


  Exit Function
ErrorHandler:
  HandleError True, "HasSelectedFeatures " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Public Function ParseOMMapNum(sVal As String, sPartName As String) As String
  On Error GoTo ErrorHandler


    'Return specific ORMAP values from this string as the whole number represents
    'multiple entities
768:     If Not Len(sVal) = 24 Then
        'MsgBox "ORMAPMapNumber shoud be 24 characters and instead is " & Len(sVal)
770:         ParseOMMapNum = ""
        Exit Function
772:     End If
    Select Case LCase(sPartName)
        Case "county"
775:             ParseOMMapNum = ExtractString(sVal, 1, 2)
            'If CLng(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "town"
778:             ParseOMMapNum = ExtractString(sVal, 3, 4)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "000"
        Case "townpart"
781:             ParseOMMapNum = ExtractString(sVal, 5, 7)
            'If CDbl(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "towndir"
784:             ParseOMMapNum = ExtractString(sVal, 8, 8)
        Case "range"
786:             ParseOMMapNum = ExtractString(sVal, 9, 10)
            'If CLng(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "rangepart"
789:             ParseOMMapNum = ExtractString(sVal, 11, 13)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "000"
        Case "rangedir"
792:             ParseOMMapNum = ExtractString(sVal, 14, 14)
        Case "section"
794:             ParseOMMapNum = ExtractString(sVal, 15, 16)
            'If CLng(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "qtr"
797:             ParseOMMapNum = ExtractString(sVal, 17, 17)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "0"
        Case "qtrqtr"
800:             ParseOMMapNum = ExtractString(sVal, 18, 18)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "0"
        Case "anomaly"
803:             ParseOMMapNum = ExtractString(sVal, 19, 20)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "00"
        Case "suffixtype"
806:              ParseOMMapNum = ExtractString(sVal, 21, 21)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "0"
        Case "suffixnum"
809:             ParseOMMapNum = ExtractString(sVal, 22, 24)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "000"
        Case Else
            'some handling?
813:     End Select


  Exit Function
ErrorHandler:
  HandleError True, "ParseOMMapNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function FormatOMMapNum(sVal As String, sPartName As String) As String
  On Error GoTo ErrorHandler


    'Return properly formatted part of OM MapNum string
826:     FormatOMMapNum = sVal
    Select Case LCase(sPartName)
        Case "county"
829:             If Len(FormatOMMapNum) <> 2 Then
830:                 FormatOMMapNum = AddLeadingZeros(FormatOMMapNum, 2)
831:             End If
        Case "town"
833:             If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "00"
        Case "townpart"
835:             If Len(FormatOMMapNum) <> 3 Then FormatOMMapNum = "000"
        Case "towndir"
837:             If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "N"
        Case "range"
839:             If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "01"
        Case "rangepart"
841:             If Len(FormatOMMapNum) <> 3 Then FormatOMMapNum = "000"
        Case "rangedir"
843:             If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "W"
        Case "section"
845:             If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "00"
        Case "qtr"
847:             If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "qtrqtr"
849:             If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "suffixtype"
851:             If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "suffixnum"
853:             If Len(FormatOMMapNum) <> 0 And Len(FormatOMMapNum) > 3 Then
854:                 FormatOMMapNum = "000"
                Exit Function
856:             ElseIf Len(FormatOMMapNum) = 1 Then
857:                 FormatOMMapNum = "00" & FormatOMMapNum
                Exit Function
859:             ElseIf Len(FormatOMMapNum) = 2 Then
860:                 FormatOMMapNum = "0" & FormatOMMapNum
                Exit Function
862:             End If
        Case "anomaly"
864:             If Len(FormatOMMapNum) > 2 Or Len(FormatOMMapNum) = 0 Then
865:                 FormatOMMapNum = "00"
                Exit Function
867:             ElseIf Len(FormatOMMapNum) = 1 Then
868:                 FormatOMMapNum = "0" & FormatOMMapNum
                Exit Function
870:             End If
        Case Else
            'Nothing to implement now
873:     End Select


  Exit Function
ErrorHandler:
  HandleError True, "FormatOMMapNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Private Function ExtractString(sFullString As String, llow As Long, lhigh As Long) As String
  On Error GoTo ErrorHandler


    'Use the low and high values to extract the required string
    Dim sVal1 As String
    Dim sVal2 As String
889:     sVal1 = Right(sFullString, Len(sFullString) - (llow - 1))
890:     sVal2 = Left(sVal1, (lhigh - llow) + 1)
891:     ExtractString = sVal2


  Exit Function
ErrorHandler:
  HandleError False, "ExtractString " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function IsTaxlot(obj As IObject) As Boolean
  On Error GoTo ErrorHandler


    'Determines if this feature is in the Taxlot feature class
    'Used by generic functions to determine what has to be done
    Dim pOC As IObjectClass
    Dim pDS As IDataset
907:     Set pOC = obj.Class
908:     Set pDS = pOC
909:     If LCase(pDS.Name) = LCase(g_pFldnames.FCTaxlot) Then
910:         IsTaxlot = True
911:     Else
912:         IsTaxlot = False
913:     End If


  Exit Function
ErrorHandler:
  HandleError True, "IsTaxlot " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function IsAnno(obj As IObject) As Boolean
  On Error GoTo ErrorHandler

924:     IsAnno = False
    'Determines if this feature is annotation feature class
    'Used by generic functions to determine what has to be done
    Dim pOC As IObjectClass
    Dim pDS As IDataset
929:     Set pOC = obj.Class
930:     Set pDS = pOC
931:     If TypeOf obj Is IFeature Then
        Dim pFC As IFeatureClass
933:         Set pFC = pOC
934:         If pFC.FeatureType = esriFTAnnotation Then IsAnno = True
935:     End If


  Exit Function
ErrorHandler:
  HandleError True, "IsAnno " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
Public Function ValidateTaxlotNum(sEnteredTLval As String, pGeometry As IGeometry) As Boolean
  On Error GoTo ErrorHandler


    'Ensure that the numeric taxlot number is unique within the current map index
    Dim pTaxlotFlayer As IFeatureLayer2
    Dim pTaxlotFclass As IFeatureClass
    Dim pMIFlayer As IFeatureLayer2
    Dim pMIFclass As IFeatureClass
951:     Set pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
952:     If pTaxlotFlayer Is Nothing Then
953:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        Exit Function
956:     End If
957:     Set pTaxlotFclass = pTaxlotFlayer.FeatureClass
958:     Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
959:     If pMIFlayer Is Nothing Then
960:         MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
        Exit Function
963:     End If
964:     Set pMIFclass = pMIFlayer.FeatureClass
    'Get fields needed to populate the form
    Dim lMIOMNum As Long
    Dim lTLOMNum As Long
    Dim lTLTaxlot As Long
    Dim sMIOMval As String
    Dim sTLOMval As String
971:     lMIOMNum = modUtils.LocateFields(pMIFclass, g_pFldnames.MIORMAPMapNumberFN)
972:     lTLOMNum = modUtils.LocateFields(pTaxlotFclass, g_pFldnames.TLOrmapMapNumberFN)
973:     lTLTaxlot = modUtils.LocateFields(pTaxlotFclass, g_pFldnames.TLTaxlotFN)
974:     sMIOMval = GetValueViaOverlay(pGeometry, pMIFclass, g_pFldnames.MIORMAPMapNumberFN)
    'if no Mapindex or ORMAP mapnum, then no need to continue
976:     If sMIOMval = "" Then
977:         ValidateTaxlotNum = True
        Exit Function
979:     End If
    'Make sure this number is unique within taxlots with this OM number
    Dim pCursor As ICursor
    Dim sWhere As String
983:     sWhere = g_pFldnames.TLOrmapMapNumberFN & " = '" & sMIOMval & _
            "' and " & g_pFldnames.TLTaxlotFN & " = '" & sEnteredTLval & "'"
985:     Set pCursor = AttributeQuery(pTaxlotFclass, sWhere)
986:     If Not pCursor Is Nothing Then
        Dim pRow As IRow
988:         Set pRow = pCursor.NextRow
989:         If Not pRow Is Nothing Then
990:             ValidateTaxlotNum = False
991:         Else
992:             ValidateTaxlotNum = True
993:         End If
994:     Else
995:         ValidateTaxlotNum = True
996:     End If


  Exit Function
ErrorHandler:
  HandleError True, "ValidateTaxlotNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub CalcTaxlotValues(pFeat As IFeature, pMIFlayer As IFeatureLayer)
  On Error GoTo ErrorHandler

    'Calculates Taxlot vaules from ORMAPMapnum
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
    Dim lTaxlotMapAcres As Long
    Dim lTaxlotShapeArea As Long
    Dim response As Variant
    
1033:     Set pTaxlotFclass = pFeat.Class
    'Find MapIndex
1035:     Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
1036:     If pMIFlayer Is Nothing Then
1037:         response = MsgBox("Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex & ".  " & _
        "Load " & g_pFldnames.FCMapIndex & " automatically?", vbYesNo)
        If response <> vbYes Then Exit Sub
1041:         modUtils.LoadFCIntoMap g_pFldnames.FCMapIndex, pTaxlotFclass
        'Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        If pMIFlayer Is Nothing Then Exit Sub
1044:     End If

    'Find all fields needed
1047:     m_bContinue = True
1048:     lOMTLNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLOrmapTaxlotFN)
1049:     lOMNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLOrmapMapNumberFN)
1050:     lMNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLMapNumberFN)
1051:     lTLCntyFld = LocateFields(pTaxlotFclass, g_pFldnames.TLCountyFN)
1052:     lTaxlotFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTaxlotFN)
1053:     lTLTownFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTownFN)
1054:     lTLTownPartFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTownPartFN)
1055:     lTLTownDirFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTownDirFN)
1056:     lTLRangeFld = LocateFields(pTaxlotFclass, g_pFldnames.TLRangeFN)
1057:     lTLRangePartFld = LocateFields(pTaxlotFclass, g_pFldnames.TLRangePartFN)
1058:     lTLRangeDirFld = LocateFields(pTaxlotFclass, g_pFldnames.TLRangeDirFN)
1059:     lTLSectNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSectNumberFN)
1060:     lTLQtrFld = LocateFields(pTaxlotFclass, g_pFldnames.TLQtrFN)
1061:     lTLQQFld = LocateFields(pTaxlotFclass, g_pFldnames.TLQtrQtrFN)
1062:     lTLMapSufTypeFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSufTypeFN)
1063:     lTLMapSufNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSufNumFN)
1064:     lTLSpecInterestFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSpecInterestFN)
1065:     lTLMapTaxlotFld = LocateFields(pTaxlotFclass, g_pFldnames.TLMapTaxlotFN)
1066:     lTLMapNumberFld = LocateFields(pTaxlotFclass, g_pFldnames.TLMapNumberFN)
1067:     lTaxlotMapAcres = LocateFields(pTaxlotFclass, g_pFldnames.TLMapAcresFN)
1068:     lTaxlotShapeArea = LocateFields(pTaxlotFclass, "SHAPE_Area")
    If Not m_bContinue Then Exit Sub 'If any fields not found
    'Update Acreage
1071:     pFeat.Value(lTaxlotMapAcres) = (pFeat.Value(lTaxlotShapeArea) / 43560)

    'Obtain the map index poly via overlay
    Dim sExistVal As String
    Dim pArea As IArea
    Dim pCenter As IPoint
1077:     Set pArea = pFeat.Shape
1078:     Set pCenter = pArea.Centroid
    Dim sExistOMMapNum As String
    Dim sExistMapNum As String
    'Return the OMMapNum and MapNum and insert values into Taxlot
1082:     sExistOMMapNum = GetValueViaOverlay(pCenter, pMIFlayer.FeatureClass, g_pFldnames.MIORMAPMapNumberFN)
    If sExistOMMapNum = "" Then Exit Sub 'If no value for whatever reason, don't continue
1084:     sExistMapNum = GetValueViaOverlay(pCenter, pMIFlayer.FeatureClass, g_pFldnames.MIMapNumberFN)
    If sExistMapNum = "" Then Exit Sub 'If no value for whatever reason, don't continue
    'Store individual components of map number in taxlot
1087:     pFeat.Value(lOMNumFld) = sExistOMMapNum
1088:     pFeat.Value(lMNumFld) = sExistMapNum
    
    'County
1091:     sExistVal = ParseOMMapNum(sExistOMMapNum, "county")
1092:     sExistVal = ConvertCode(pFeat, g_pFldnames.TLCountyFN, sExistVal)
1093:     pFeat.Value(lTLCntyFld) = CInt(sExistVal) 'Store county in county field
    
    'Town
1096:     sExistVal = ParseOMMapNum(sExistOMMapNum, "town")
1097:     pFeat.Value(lTLTownFld) = CInt(sExistVal)

    'TownPart
1100:     sExistVal = ParseOMMapNum(sExistOMMapNum, "townpart")
1101:     pFeat.Value(lTLTownPartFld) = CDbl(sExistVal)

    'TownDir
1104:     sExistVal = ParseOMMapNum(sExistOMMapNum, "towndir")
1105:     pFeat.Value(lTLTownDirFld) = sExistVal

    'Range
1108:     sExistVal = ParseOMMapNum(sExistOMMapNum, "range")
1109:     pFeat.Value(lTLRangeFld) = CInt(sExistVal)

    'RangePart
1112:     sExistVal = ParseOMMapNum(sExistOMMapNum, "rangepart")
1113:     pFeat.Value(lTLRangePartFld) = CDbl(sExistVal)

    'RangeDir
1116:     sExistVal = ParseOMMapNum(sExistOMMapNum, "rangedir")
1117:     pFeat.Value(lTLRangeDirFld) = sExistVal

    'Section
1120:     sExistVal = ParseOMMapNum(sExistOMMapNum, "section")
1121:     pFeat.Value(lTLSectNumFld) = CInt(sExistVal)
 
    'Qtr
1124:     sExistVal = ParseOMMapNum(sExistOMMapNum, "qtr")
1125:     pFeat.Value(lTLQtrFld) = sExistVal
    
    'QtrQtr
1128:     sExistVal = ParseOMMapNum(sExistOMMapNum, "qtrqtr")
1129:     pFeat.Value(lTLQQFld) = sExistVal

    'MapSuffixType
1132:     sExistVal = ParseOMMapNum(sExistOMMapNum, "suffixtype")
1133:     sExistVal = ConvertCode(pFeat, g_pFldnames.TLSufTypeFN, sExistVal)
1134:     pFeat.Value(lTLMapSufTypeFld) = sExistVal
    
    'MapSuffixNum
1137:     sExistVal = ParseOMMapNum(sExistOMMapNum, "suffixnum")
1138:     pFeat.Value(lTLMapSufNumFld) = sExistVal
    
    'SpecialInterest
1141:     sExistVal = IIf(IsNull(pFeat.Value(lTLSpecInterestFld)), "00000", pFeat.Value(lTLSpecInterestFld))
1142:     If Len(sExistVal) < 5 Then
1143:      Do Until Len(sExistVal) = 5
1144:         sExistVal = "0" & sExistVal
1145:      Loop
1146:     End If
1147:     pFeat.Value(lTLSpecInterestFld) = sExistVal
    
    'Recalculate OMTaxlot
    If IsNull(pFeat.Value(lTaxlotFld)) Then Exit Sub
    Dim sTaxlotVal As String
    'Taxlot has actual taxlot number.  ORMAPTaxlot requires a 5-digit number, so leading zeros have to be added
1153:     sTaxlotVal = pFeat.Value(lTaxlotFld)
1154:     sTaxlotVal = AddLeadingZeros(sTaxlotVal, 5)
    Dim sNewOMTLNum As String
    Dim sExistOMTLNum As String
    If IsNull(pFeat.Value(lOMTLNumFld)) Then Exit Sub
1158:     sExistOMTLNum = pFeat.Value(lOMTLNumFld)
1159:     sNewOMTLNum = CalcOMTLNum(sExistOMTLNum, pFeat, sTaxlotVal)
    'If no changes, don't save value
1161:     If Not sExistOMTLNum = sNewOMTLNum Then
1162:         pFeat.Value(lOMTLNumFld) = sNewOMTLNum
1163:     End If
    

  Exit Sub
ErrorHandler:
  HandleError True, "CalcTaxlotValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function AddLeadingZeros(sCurString As String, lWidth As Long) As String
  On Error GoTo ErrorHandler

        'Add leading zeros if necessary
1175:         If Len(sCurString) < lWidth Then
1176:          Do Until Len(sCurString) = lWidth
1177:             sCurString = "0" & sCurString
1178:          Loop
1179:         End If
1180:         AddLeadingZeros = sCurString

  Exit Function
ErrorHandler:
  HandleError True, "AddLeadingZeros " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function GetCentroid(pFeat As IFeature) As IPoint
  On Error GoTo ErrorHandler

    'Determines if this feature is annotation feature class
1191:         If pFeat.FeatureType = esriFTAnnotation Or pFeat.FeatureType = esriFTDimension Then
            Dim pArea As IArea
1193:             Set pArea = pFeat.Shape
1194:             Set GetCentroid = pArea.Centroid
1195:         End If


  Exit Function
ErrorHandler:
  HandleError True, "GetCentroid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Function CT_GetCenterOfEnvelope(pEnv As IEnvelope) As IPoint
  On Error GoTo ErrorHandler

    Dim pCenter As IPoint
1207:     Set pCenter = New Point
1208:     pCenter.X = pEnv.XMin + (pEnv.XMax - pEnv.XMin) / 2
1209:     pCenter.Y = pEnv.YMin + (pEnv.YMax - pEnv.YMin) / 2
1210:     Set CT_GetCenterOfEnvelope = pCenter

  Exit Function
ErrorHandler:
  HandleError True, "CT_GetCenterOfEnvelope " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function GetRelatedObjects(pObj As IObject) As IFeature
  On Error GoTo ErrorHandler


    'Using the passed in object, get related features through a relationship class
    'This is optimized for anno because there is a single relationship class
    Dim pEnumRelClass As IEnumRelationshipClass
    Dim pRelClass As IRelationshipClass
    Dim pParentSet As esriSystem.ISet
    Dim pParentFeat As IFeature
    
1228:     Set pEnumRelClass = pObj.Class.RelationshipClasses(esriRelRoleAny)
1229:     If Not pEnumRelClass Is Nothing Then
1230:       Set pRelClass = pEnumRelClass.Next
1231:       If Not pRelClass Is Nothing Then
1232:           Set pParentSet = pRelClass.GetObjectsRelatedToObject(pObj)
1233:       End If
1234:     Else
        Exit Function
1236:     End If
1237:     If Not pParentSet Is Nothing Then
1238:         Set pParentFeat = pParentSet.Next
1239:         If Not pParentFeat Is Nothing Then
1240:             Set GetRelatedObjects = pParentFeat
1241:         End If
1242:     End If


  Exit Function
ErrorHandler:
  HandleError True, "GetRelatedObjects " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function GetAnnoSizeByScale(sFCName As String, lScale As Long) As Double
  On Error GoTo ErrorHandler

    Dim dSize As Double
1254:     If LCase(sFCName) = LCase(g_pFldnames.FCTLAcrAnno) Then
            'Determine anno size based on scale
1256:          If lScale = 120 Then dSize = 0.8
1257:          If lScale = 240 Then dSize = 1.6
1258:          If lScale = 360 Then dSize = 2
1259:          If lScale = 480 Then dSize = 3.2
1260:          If lScale = 600 Then dSize = 4
1261:          If lScale = 1200 Then dSize = 8
1262:          If lScale = 2400 Then dSize = 16
1263:          If lScale = 4800 Then dSize = 32
1264:          If lScale = 9600 Then dSize = 48
1265:          If lScale = 24000 Then dSize = 160
1266:     ElseIf LCase(sFCName) = LCase(g_pFldnames.FCTLNumAnno) Then
1267:          If lScale = 120 Then dSize = 1
1268:          If lScale = 240 Then dSize = 2
1269:          If lScale = 360 Then dSize = 3
1270:          If lScale = 480 Then dSize = 4
1271:          If lScale = 600 Then dSize = 5
1272:          If lScale = 1200 Then dSize = 10
1273:          If lScale = 2400 Then dSize = 20
1274:          If lScale = 4800 Then dSize = 40
1275:          If lScale = 9600 Then dSize = 64
1276:          If lScale = 24000 Then dSize = 200
1277:     Else
        'Something not being trapped
1279:         dSize = 10
1280:     End If
    'TODO #####
    'Determine a default
1283:     If dSize = 0 Then dSize = 5
1284:     GetAnnoSizeByScale = dSize

  Exit Function
ErrorHandler:
  HandleError True, "GetAnnoSizeByScale " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function FileExists(sPath As String) As Boolean
  On Error GoTo ErrorHandler


    Dim pFSO As Object
1296:     Set pFSO = CreateObject("Scripting.FileSystemObject")
1297:     If Not pFSO.FileExists(sPath) Then
1298:         FileExists = False
1299:     Else
1300:         FileExists = True
1301:     End If


  Exit Function
ErrorHandler:
  HandleError True, "FileExists " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
Public Function GetAppRef() As IApplication
  On Error GoTo ErrorHandler

'Used to obtain a reference the the Application, which is used throughout the code
'This is a more complex process with VB code because the code does not live in the MXD
Dim doc As IDocument
Dim app As IApplication
Dim pMXDoc As IMxDocument
Dim pobjectFactory As IObjectFactory
Dim rot As AppROT
Dim strName As String

1320: Set rot = New AppROT
1321: If rot.Count = 1 Then
1322:     Set app = rot.Item(0) 'ArcCatalog
1323: Else
1324:     Set app = rot.Item(1) 'ArcMap
1325: End If
1326: Set pobjectFactory = app

1328: Set GetAppRef = app


  Exit Function
ErrorHandler:
  HandleError True, "GetAppRef " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function GetMXDocRef() As IMxDocument
  On Error GoTo ErrorHandler

'Get a reference to the current map document
Dim doc As IDocument
Dim app As IApplication
Dim pMXDoc As IMxDocument
Dim pobjectFactory As IObjectFactory
Dim rot As AppROT
Dim strName As String

1347: Set rot = New AppROT
1348: If rot.Count = 1 Then
1349:     Set app = rot.Item(0) 'ArcCatalog
1350: Else
1351:     Set app = rot.Item(1) 'ArcMap
1352: End If
1353: Set pobjectFactory = app
1354: Set pMXDoc = app.Document

1356: Set GetMXDocRef = pMXDoc


  Exit Function
ErrorHandler:
  HandleError True, "GetMXDocRef " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub LoadFCIntoMap(sFCName As String, pOtherFC As IFeatureClass)
  On Error GoTo ErrorHandler


    'Loads a feature class into the current map
    'Feature class must be in the same feature dataset as pOtherFC
    Dim pWS As IWorkspace
    Dim pFWS As IFeatureWorkspace
    Dim pFC As IFeatureClass
    Dim pFeatLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim pMXDoc As IMxDocument
    Dim pMap As IMap
1377:     Set pWS = pOtherFC.FeatureDataset.Workspace
1378:     Set pFWS = pWS
1379:     Set pFC = pFWS.OpenFeatureClass(sFCName)
1380:     Set pFeatLayer = New FeatureLayer
1381:     Set pFeatLayer.FeatureClass = pFC
1382:     Set pDataset = pFC
1383:     pFeatLayer.Name = pDataset.Name
1384:     Set pMXDoc = g_pApp.Document
1385:     Set pMap = pMXDoc.FocusMap
1386:     pMap.AddLayer pFeatLayer
1387:     pMXDoc.CurrentContentsView.Refresh 0


  Exit Sub
ErrorHandler:
  HandleError True, "LoadFCIntoMap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




Public Function IsOrMapFeature(obj As esriGeoDatabase.IObject) As Boolean
  On Error GoTo ErrorHandler

    'Determines if a feature class part of the ORMAP design,
    'If not, it will not be used by any code in this project
    Dim pOC As IObjectClass
    Dim pDSet As IDataset
    Dim pName As String
1406:     Set pOC = obj.Class
1407:     Set pDSet = pOC
1408:     pName = LCase(Trim(pDSet.Name))
1409:     If pName = LCase(Trim(g_pFldnames.FCAnno10)) Or pName = LCase(Trim(g_pFldnames.FCAnno100)) Or pName = LCase(Trim(g_pFldnames.FCAnno20)) Or _
        pName = LCase(Trim(g_pFldnames.FCAnno200)) Or pName = LCase(Trim(g_pFldnames.FCAnno2000)) Or pName = LCase(Trim(g_pFldnames.FCAnno30)) Or _
        pName = LCase(Trim(g_pFldnames.FCAnno40)) Or pName = LCase(Trim(g_pFldnames.FCAnno400)) Or pName = LCase(Trim(g_pFldnames.FCAnno50)) Or _
        pName = LCase(Trim(g_pFldnames.FCAnno800)) Or pName = LCase(Trim(g_pFldnames.FCCartoLines)) Or pName = LCase(Trim(g_pFldnames.FCLotsAnno)) Or _
        pName = LCase(Trim(g_pFldnames.FCMapIndex)) Or pName = LCase(Trim(g_pFldnames.FCPlats)) Or pName = LCase(Trim(g_pFldnames.FCReferenceLines)) Or _
        pName = LCase(Trim(g_pFldnames.FCTaxCode)) Or pName = LCase(Trim(g_pFldnames.FCTaxCode)) Or pName = LCase(Trim(g_pFldnames.FCTaxCodeAnno)) Or _
        pName = LCase(Trim(g_pFldnames.FCTaxlot)) Or pName = LCase(Trim(g_pFldnames.FCTaxlotLines)) Or pName = LCase(Trim(g_pFldnames.FCTLAcrAnno)) Or _
        pName = LCase(Trim(g_pFldnames.FCTLNumAnno)) Then
1417:         IsOrMapFeature = True
1418:     Else
1419:         IsOrMapFeature = False
1420:     End If


  Exit Function
ErrorHandler:
  HandleError True, "IsOrMapFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub SetAnnoSize(obj As IObject, pFeat As IFeature)
  On Error GoTo ErrorHandler

    'If working with anno, determine what size it should be
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
    
    'Capture MapNumber for each anno feature created
1451:     lAnnoMapNumFld = LocateFields(obj.Class, g_pFldnames.MIMapNumberFN)
    If lAnnoMapNumFld = -1 Then Exit Sub
    'If new anno feature with no text, determine if it has a shape
    Dim pAnnoFeat As IFeature
    Dim pAOC As IObjectClass
1456:     Set pAOC = obj.Class
1457:     Set pAnnoFeat = obj
    Dim lFld As Long
'    Dim g As Long
'    g = 0
'    While g < pAnnoFeat.Fields.FieldCount
'        MsgBox "field = " & pAnnoFeat.Fields.Field(g).Name
'        g = g + 1
'    Wend
1465:     lFld = pAnnoFeat.Fields.FindField("TextString")
1466:     If lFld = -1 Then
1467:         MsgBox "Unable to locate textstring field in anno class.  Cannot set size", vbCritical
        Exit Sub
1469:     End If
    Dim vVal As Variant
1471:     vVal = pAnnoFeat.Value(lFld)
    If IsNull(vVal) Then Exit Sub
        
    
1475:     Set pFeat = obj
1476:     Set pGeometry = pFeat.Shape
    If pGeometry.IsEmpty Then Exit Sub
1478:     Set pEnv = pGeometry.Envelope
1479:     Set pCenter = CT_GetCenterOfEnvelope(pEnv)
1480:     Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then Exit Sub
1482:     Set pMIFclass = pMIFlayer.FeatureClass
1483:     sMapNum = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
1484:     obj.Value(lAnnoMapNumFld) = sMapNum
    
    'Update the size to reflect current mapscale
1487:     sMapScale = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
    If IsNull(sMapScale) Then Exit Sub
    
    'Determine which annotation class this is
1491:     Set pAnnoClass = obj.Class
1492:     Set pAnnoDset = pAnnoClass
    'If other anno, don't continue
1494:     If LCase(pAnnoDset.Name) <> LCase(g_pFldnames.FCTLAcrAnno) And LCase(pAnnoDset.Name) <> LCase(g_pFldnames.FCTLNumAnno) Then
        Exit Sub
1496:     End If
    
1498:     dSize = modUtils.GetAnnoSizeByScale(pAnnoDset.Name, CLng(sMapScale))
    'Get the anno feature, its symbol, set the appropriate size
1500:     Set pAnnotationFeature = obj
1501:     Set pAnnotationElement = pAnnotationFeature.Annotation
1502:     Set pElement = pAnnotationElement
1503:     Set pTextElement = pElement
1504:     Set pTextSym = pTextElement.Symbol
'MsgBox "Size = " & dSize
1506:     pTextSym.Size = dSize
1507:     pTextElement.Symbol = pTextSym
1508:     Set pElement = pTextElement
1509:     Set pAnnotationElement = pElement
1510:     pAnnotationFeature.Annotation = pAnnotationElement


  Exit Sub
ErrorHandler:
  HandleError True, "SetAnnoSize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function LocateFields(pFClass As IFeatureClass, pFldName As String) As Long
  On Error GoTo ErrorHandler

    'Return the index (location) of a field within a feature class
    Dim lFld As Long
1523:     lFld = pFClass.Fields.FindField(pFldName)
1524:     If lFld > -1 Then
1525:       LocateFields = lFld
1526:     Else
1527:         MsgBox "Unable to locate " & pFldName & " field in " & _
        pFClass.AliasName & " feature class"
1529:     End If


  Exit Function
ErrorHandler:
  HandleError True, "LocateFields " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub UpdateAutoFields(pFeat As IFeature)
  On Error GoTo ErrorHandler


'Code to update AutoDate and AutoWho
    Dim lAutoDateFld As Long
    Dim lAutoWhoFld As Long
1544:     lAutoDateFld = pFeat.Fields.FindField(g_pFldnames.AutoDateFN)
1545:     If lAutoDateFld > -1 Then
1546:         pFeat.Value(lAutoDateFld) = Now
1547:     End If
1548:     lAutoWhoFld = pFeat.Fields.FindField(g_pFldnames.AutoWhoFN)
1549:     If lAutoWhoFld > -1 Then
1550:         pFeat.Value(lAutoWhoFld) = Environ("USERNAME")
1551:     End If


  Exit Sub
ErrorHandler:
  HandleError True, "UpdateAutoFields " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function Validate5Digits(sString As String)
  On Error GoTo ErrorHandler

            'Make sure taxlot number is 5 characters
1563:             If Len(sString) < 5 Then
1564:              Do Until Len(sString) = 5
1565:                 sString = "0" & sString
1566:              Loop
1567:             End If
1568:             Validate5Digits = sString

  Exit Function
ErrorHandler:
  HandleError True, "Validate5Digits " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function GetSpecialInterests(pFeature As IFeature) As String
  On Error GoTo ErrorHandler

        Dim lTLSpecInterestFld As Long
        Dim sTLSpecVAl As String
1580:         lTLSpecInterestFld = modUtils.LocateFields(pFeature.Class, g_pFldnames.TLSpecInterestFN)
1581:         If lTLSpecInterestFld = -1 Then
1582:             sTLSpecVAl = "00000"
1583:         Else
1584:             If Not IsNull(pFeature.Value(lTLSpecInterestFld)) Then
1585:                 sTLSpecVAl = pFeature.Value(lTLSpecInterestFld)
1586:             Else
1587:                 sTLSpecVAl = "00000"
1588:             End If
            'Verify that it is 5 digits
1590:             If Len(sTLSpecVAl) < 5 Then
1591:              Do Until Len(sTLSpecVAl) = 5
1592:                 sTLSpecVAl = "0" & sTLSpecVAl
1593:              Loop
1594:             End If
1595:         End If
1596:         GetSpecialInterests = sTLSpecVAl

  Exit Function
ErrorHandler:
  HandleError True, "GetSpecialInterests " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function CalcOMTLNum(sExistOMNum As String, pFeat As IFeature, sTLVal As String) As String
  On Error GoTo ErrorHandler

        'Calculate ORMAPtaxlot because one if its components may have changed
        Dim sShortOMNum As String 'Remove suffixTYpe and suffixNum
        Dim sTLSpecVAl As String
        Dim sOMTLNval As String
1610:         sShortOMNum = ShortenOMMapNum(sExistOMNum)
1611:         sTLSpecVAl = GetSpecialInterests(pFeat)
1612:         sOMTLNval = sShortOMNum & sTLVal & sTLSpecVAl
1613:         CalcOMTLNum = sOMTLNval

  Exit Function
ErrorHandler:
  HandleError True, "CalcOMTLNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function ShortenOMMapNum(sOMVal As String) As String
  On Error GoTo ErrorHandler

    'Remove two values from the ORMAPMap number for the purpose of populating ORMAPTaxlog
1624:     ShortenOMMapNum = Left(sOMVal, 20)

  Exit Function
ErrorHandler:
  HandleError True, "ShortenOMMapNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub ZoomToExtent(pEnv As IEnvelope, pMXDoc As IMxDocument)
    'Zooms the current extent to the passed in envelope (i.e. zoom to feature)
    'Works for Layout and Data view
    Dim pMap As IMap
    Dim pActiveView As IActiveView
1636:     Set pMap = pMXDoc.FocusMap
1637:     Set pActiveView = pMap

1639:     pActiveView.Extent = pEnv
1640:     pActiveView.Refresh
End Sub

