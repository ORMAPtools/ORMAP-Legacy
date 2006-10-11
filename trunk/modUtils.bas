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
    
    Set pMXDoc = g_pApp.Document
    Set pMap = pMXDoc.FocusMap
  
    With pMap
        For i = 0 To .LayerCount - 1
            If TypeOf .Layer(i) Is IFeatureLayer Then
                Set pFeatureLayer = .Layer(i)
                Set pDataset = pFeatureLayer.FeatureClass
                If Not pDataset Is Nothing Then
'++ JWM 10/11/2006 using strcomp function
                    If StrComp(pDataset.Name, DatasetName, vbTextCompare) = 0 Then
                        Set FindFeatureLayerByDS = pFeatureLayer
                        Exit For
                    End If
                End If
            End If
        Next i
    End With
  
    If pFeatureLayer Is Nothing Then

    End If
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
  Set pObjClass = pObj.Class
  Set pDataset = pObjClass
  Set pFWS = pDataset.Workspace
  Set GetFWorkspace = pFWS

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
    lFld = pRow.Fields.FindField(pFldName)
    If lFld > -1 Then
      If pDataType = "date" Then
        'If a date and value is null, return a default date value
        '??? How should this be treated?
        Dim pDate As Date
        sVal = IIf(IsNull(pRow.Value(lFld)), pDate, pRow.Value(lFld))
      Else
        sVal = IIf(IsNull(pRow.Value(lFld)), "", pRow.Value(lFld))
      End If
      'Determine if domain field
      Dim pField As IField
      Set pField = pRow.Fields.Field(lFld)
      Dim pDomain As IDomain
      Set pDomain = pField.Domain
      If pDomain Is Nothing Then
        ReadValue = sVal
        GoTo Process_Exit
      Else
        'Determine type of domain  -If Coded Value, get the description
        If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
          Set pCVDomain = pDomain
          Dim lCode As Long
          Dim vDomainVal As Variant
          vDomainVal = pRow.Value(lFld)
          Dim i As Integer
          'Search the domain for the code
          For i = 0 To pCVDomain.CodeCount - 1
             If pCVDomain.Value(i) = vDomainVal Then
              'return the description
              ReadValue = pCVDomain.Name(i)
              GoTo Process_Exit
            End If
          Next i
        Else ' If range domain, return the numeric value
          ReadValue = sVal
          GoTo Process_Exit:
        End If
      End If  'If pDomain is nothing/Else
      ReadValue = sVal
    Else
      'Field not found
      ReadValue = ""
    End If 'If lFld > -1/Else

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
      lFld = pFields.FindField(pFldName)
      If lFld > -1 Then
        Dim pField As IField
        Set pField = pFields.Field(lFld)
        Dim pDomain As IDomain
        Set pDomain = pField.Domain
        If pDomain Is Nothing Then
          AddCodesToCmb = False
          GoTo Process_Exit
        Else
          'Determine type of domain  -If Coded Value, get the description
          If TypeOf pDomain Is ICodedValueDomain Then
            Dim pCVDomain As ICodedValueDomain
            Set pCVDomain = pDomain
            ' +++ Get a count of the coded values
            Dim lCodes As Long
            Dim i As Long
            lCodes = pCVDomain.CodeCount
            Dim sVal As Variant
            ' +++ Loop through the list of values and add them
            ' +++ and their names to the combo box
            If Not blnAllowSpace Then
              With cboValues
              If .ListCount > 0 Then
                If (.List(0) = "") Or (.List(0) = "") Then
                  .RemoveItem (0)
                End If
              End If
              If .ListCount > 0 Then
'++ JWM 10/11/2006 Is this if statement comparing against the same thing ?
                If (.List(.ListCount - 1) = "") Or (.List(.ListCount - 1) = "") Then
                  .RemoveItem (.ListCount - 1)
                End If
              End If
              End With
            End If
            For i = 0 To lCodes - 1
              'Commented line adds codes and description
              'cboValues.AddItem pCVDomain.Value(i) & ": " & pCVDomain.Name(i)
              cboValues.AddItem pCVDomain.Name(i)
            Next i
            'Successful completion of addition
            'If current value is null, add an empty string and make it active
            If curVal = "" Then
              If blnAllowSpace Then
                cboValues.AddItem ""
                cboValues.ListIndex = FindControlString(cboValues, "", 0, True)
                'cboValues.Text = ""
              Else
                cboValues.ListIndex = 0
              End If
            Else 'Otherwise, select the existing value from the list
              cboValues.ListIndex = FindControlString(cboValues, curVal, 0, True)
            End If
            
            AddCodesToCmb = True
          Else
            'if Range Domain, do not add values
            AddCodesToCmb = False
          End If
        End If 'if a valid domain
      Else 'Field not found
        AddCodesToCmb = False
      End If

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
    lFld = pRow.Fields.FindField(pFldName)
    If lFld > -1 Then
      'Determine if domain field
      Dim pField As IField
      Set pField = pRow.Fields.Field(lFld)
      Dim pDomain As IDomain
      Set pDomain = pField.Domain
      If pDomain Is Nothing Then
        ConvertCode = pVal
        GoTo Process_Exit
      Else
        'Determine type of domain  -If Coded Value, get the description
        If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
          Set pCVDomain = pDomain
          Dim lCode As Long
          Dim i As Integer
          'Given the description, search the domain for the code
          For i = 0 To pCVDomain.CodeCount - 1
            If pCVDomain.Name(i) = pVal Then
              ConvertCode = pCVDomain.Value(i) 'Return the code value
              GoTo Process_Exit
            End If
          Next i
        Else ' If range domain, return the numeric value
          ConvertCode = pVal
          GoTo Process_Exit
        End If
      End If  'If pDomain is nothing/Else
      ConvertCode = pVal
    Else
      'Field not found
      ConvertCode = ""
    End If 'If lFld > -1/Else

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
    lFld = pFlds.FindField(pFldName)
    If lFld > -1 Then
      'Determine if domain field
      Dim pField As IField
      Set pField = pFlds.Field(lFld)
      Dim pDomain As IDomain
      Set pDomain = pField.Domain
      If pDomain Is Nothing Then
        ConvertToDescription = pVal
        GoTo Process_Exit
      Else
        'Determine type of domain  -If Coded Value, get the description
        If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
          Set pCVDomain = pDomain
          Dim lCode As Long
          Dim i As Integer
          'Given the description, search the domain for the code
          For i = 0 To pCVDomain.CodeCount - 1
            If pCVDomain.Value(i) = pVal Then
              ConvertToDescription = pCVDomain.Name(i) 'Return the code value
              GoTo Process_Exit
            End If
          Next i
        Else ' If range domain, return the numeric value
          ConvertToDescription = pVal
          GoTo Process_Exit
        End If
      End If  'If pDomain is nothing/Else
      ConvertToDescription = pVal
    Else
      'Field not found
      ConvertToDescription = ""
    End If 'If lFld > -1/Else

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
    vValOrg = modUtils.ReadValue(pRow, pFldName)
    If vValNew <> vValOrg Then
      'Get the Code value that is to be stored in the db
      vValNew = modUtils.ConvertCode(pRow, pFldName, vValNew)
      'If the value is changed, update the row
      Dim lFld As Long
      lFld = pRow.Fields.FindField(pFldName)
      If lFld > -1 Then
        Dim pFldType As esriFieldType
        pFldType = pRow.Fields.Field(lFld).Type
        If pFldType = esriFieldTypeDouble Then
          Dim dValNew As Double
          If IsNumeric(vValNew) Then dValNew = CDbl(vValNew)
          If dValNew <> vValOrg Then
            pRow.Value(lFld) = dValNew
            pRowChanged.RowChanged = True
          End If
        ElseIf pFldType = esriFieldTypeInteger Or pFldType = esriFieldTypeSmallInteger Then
          Dim iValNew As Long
          If IsNumeric(vValNew) Then iValNew = CLng(vValNew)
          If iValNew <> vValOrg Then
            pRow.Value(lFld) = iValNew
            pRowChanged.RowChanged = True
          End If
        ElseIf pFldType = esriFieldTypeSingle Then
          Dim sValNew As Single
          If IsNumeric(vValNew) Then sValNew = CSng(vValNew)
          If sValNew <> vValOrg Then
            pRow.Value(lFld) = sValNew
            pRowChanged.RowChanged = True
          End If
        ElseIf pFldType = esriFieldTypeDate Then
          Dim dtValNew As Date
          If IsDate(vValNew) Then dtValNew = CDate(vValNew)
          If dtValNew <> vValOrg Then
            pRow.Value(lFld) = dtValNew
            pRowChanged.RowChanged = True
          End If
        ElseIf pFldType = esriFieldTypeString Then
          Dim sgValNew As String
          sgValNew = vValNew
          If sgValNew <> vValOrg Then
            pRow.Value(lFld) = sgValNew
            pRowChanged.RowChanged = True
          End If
        Else
          'Unknown field type
        End If
     End If
  End If

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

  GetValueViaOverlay = ""
  If Not pGeom Is Nothing And Not pOverlayFC Is Nothing And Not sFldName = "" Then
    Dim pFeatCur As IFeatureCursor
    Set pFeatCur = SpatialQuery(pOverlayFC, pGeom, esriSpatialRelIntersects)
    If Not pFeatCur Is Nothing Then
      'Get the first feature.  if more than one, let the user decide
      Dim pFeat As IFeature
      Set pFeat = pFeatCur.NextFeature
      If Not pFeat Is Nothing Then
        Dim lFld As Long
        lFld = pFeat.Fields.FindField(sFldName)
        If lFld > -1 Then
          'Get the  value
          GetValueViaOverlay = IIf(IsNull(pFeat.Value(lFld)), "", pFeat.Value(lFld))
        End If
      End If
    End If
  End If

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
  If TypeOf ctrl Is ListBox Then
    uMsg = IIf(ExactMatch, LB_FINDSTRINGEXACT, LB_FINDSTRING)
  ElseIf TypeOf ctrl Is ComboBox Then
    uMsg = IIf(ExactMatch, CB_FINDSTRINGEXACT, CB_FINDSTRING)
  Else
    GoTo Process_Exit
  End If
  FindControlString = SendMessageString(ctrl.hwnd, uMsg, lStartIdx, strSearch)

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
    Set pSpatialFilter = New esriGeoDatabase.SpatialFilter
    
    ' specify the geometry to query with
    Set pSpatialFilter.Geometry = searchGeometry
    
    ' specify what the geometry file is called on the Feature Class that we will be querying against
    Dim strShpFld As String
    strShpFld = pFeatureClassIN.ShapeFieldName
    pSpatialFilter.GeometryField = strShpFld
    
    'specify the type of spatial operation to use
    pSpatialFilter.SpatialRel = spatialRelation

    ' create the where statement
    pSpatialFilter.whereClause = whereClause
    
    ' create a cursor that will return the results
    Dim pFeatCursor As esriGeoDatabase.IFeatureCursor
    
    ' perform the query
    Dim pQueryFilter As esriGeoDatabase.IQueryFilter
    Set pQueryFilter = pSpatialFilter
    Set pFeatCursor = pFeatureClassIN.Search(pQueryFilter, False)
    
    Set SpatialQuery = pFeatCursor

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
    Set pSpatialFilter = New esriGeoDatabase.SpatialFilter
    
    ' specify the geometry to query with
    Set pSpatialFilter.Geometry = searchGeometry
    
    ' specify what the geometry file is called on the Feature Class that we will be querying against
    Dim strShpFld As String
    strShpFld = pFeatureClassIN.ShapeFieldName
    pSpatialFilter.GeometryField = strShpFld
    
    'specify the type of spatial operation to use
    pSpatialFilter.SpatialRel = spatialRelation

    ' create the where statement
    pSpatialFilter.whereClause = whereClause
    
    ' create a cursor that will return the results
    Dim pFeatCursor As esriGeoDatabase.IFeatureCursor
    
    ' perform the query
    Dim pQueryFilter As esriGeoDatabase.IQueryFilter
    Set pQueryFilter = pSpatialFilter
    'Set pFeatCursor = pFeatureClassIN.Search(pQueryFilter, False)
    Set pFeatCursor = pFeatureClassIN.Update(pQueryFilter, False)
    
    Set SpatialQueryForEdit = pFeatCursor

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
Set pQueryFilter = New esriGeoDatabase.QueryFilter

' create the where statement
'whereClause = Replace(whereClause, "HYDRO1.", "")
pQueryFilter.whereClause = whereClause

' create a cursor that will return the results
Dim pCursor As esriGeoDatabase.ICursor

' query the table passed into the fuction
Set pCursor = pTable.Search(pQueryFilter, False)

'Count the number of selected records
Dim selCount As Long
selCount = pTable.RowCount(pQueryFilter)
If selCount = 0 Then
  Set AttributeQuery = Nothing
Else
  Set AttributeQuery = pCursor
End If

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
     lFld = pTable.FindField(sFldName)
     If lFld > -1 Then
        Set pField = pTable.Fields.Field(lFld)
     Else
        GetDomainDefaultValue = ""
        GoTo Process_Exit
     End If
     Dim pDomain As IDomain
     Set pDomain = pField.Domain
      If pDomain Is Nothing Then
        GetDomainDefaultValue = ""
        GoTo Process_Exit
      Else
        'Determine type of domain  -If Coded Value, get the description
        If TypeOf pDomain Is ICodedValueDomain Then
          Dim pCVDomain As ICodedValueDomain
          Set pCVDomain = pDomain
          Dim lCode As Long
          Dim vDomainVal As Variant
          vDomainVal = pField.DefaultValue
          Dim i As Integer
          'Search the domain for the code
          For i = 0 To pCVDomain.CodeCount - 1
             If pCVDomain.Value(i) = vDomainVal Then
              'return the description
              GetDomainDefaultValue = pCVDomain.Name(i)
              GoTo Process_Exit
            End If
          Next i
        Else ' If range domain, return the numeric value
          GetDomainDefaultValue = pField.DefaultValue
          GoTo Process_Exit
        End If
      End If  'If pDomain is nothing/Else

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
  If Not TypeOf pFLayer Is IFeatureLayer Then
    GoTo Process_Exit
  End If
  
  Dim pFSelection As IFeatureSelection
  Set pFSelection = pFLayer
  
  pFSelection.SelectionSet.Search Nothing, False, GetSelectedFeatures
  
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
  If pFLayer Is Nothing Then GoTo Process_Exit
  
  '  exit if not applicable:
  If Not TypeOf pFLayer Is IFeatureLayer Then
    GoTo Process_Exit
  End If
  
  Dim pFSelection As IFeatureSelection
  
  Set pFSelection = pFLayer
  Dim pFeatCur As IFeatureCursor
  pFSelection.SelectionSet.Search Nothing, False, pFeatCur
  Dim pFeat As IFeature
  If Not pFeatCur Is Nothing Then
    Set pFeat = pFeatCur.NextFeature
    If Not pFeat Is Nothing Then 'At least one feature selected
        Set pFeat = pFeatCur.NextFeature
        If Not pFeat Is Nothing Then 'More than one selected
            HasSelectedFeatures = False
            GoTo Process_Exit
        Else
            HasSelectedFeatures = True 'Just one selected
            GoTo Process_Exit
        End If
    Else 'nothing selected
        HasSelectedFeatures = False
        GoTo Process_Exit
    End If
  End If
  
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
    
    If Not Len(sVal) = 24 Then
        'MsgBox "ORMAPMapNumber shoud be 24 characters and instead is " & Len(sVal)
        ParseOMMapNum = ""
        GoTo Process_Exit
    End If
    Select Case LCase(sPartName)
        Case "county"
            ParseOMMapNum = ExtractString(sVal, 1, 2)
            'If CLng(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "town"
            ParseOMMapNum = ExtractString(sVal, 3, 4)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "000"
        Case "townpart"
            
            ParseOMMapNum = ExtractString(sVal, 5, 7)
            'If CDbl(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "towndir"
            ParseOMMapNum = ExtractString(sVal, 8, 8)
        Case "range"
            ParseOMMapNum = ExtractString(sVal, 9, 10)
            'If CLng(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "rangepart"
            ParseOMMapNum = ExtractString(sVal, 11, 13)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "000"
        Case "rangedir"
            ParseOMMapNum = ExtractString(sVal, 14, 14)
        Case "section"
            ParseOMMapNum = ExtractString(sVal, 15, 16)
            'If CLng(ParseOMMapNum) < 10 Then ParseOMMapNum = "0" & ParseOMMapNum
        Case "qtr"
            ParseOMMapNum = ExtractString(sVal, 17, 17)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "0"
        Case "qtrqtr"
            ParseOMMapNum = ExtractString(sVal, 18, 18)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "0"
        Case "anomaly"
            ParseOMMapNum = ExtractString(sVal, 19, 20)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "00"
        Case "suffixtype"
             ParseOMMapNum = ExtractString(sVal, 21, 21)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "0"
        Case "suffixnum"
            ParseOMMapNum = ExtractString(sVal, 22, 24)
            'If Len(ParseOMMapNum) = 0 Then ParseOMMapNum = "000"
        Case Else
            'some handling?
    End Select

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
    FormatOMMapNum = asVal
    Select Case LCase(asPartName)
        Case "county"
            If Len(FormatOMMapNum) <> 2 Then
                FormatOMMapNum = AddLeadingZeros(FormatOMMapNum, 2)
            End If
        Case "town"
            If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "00"
        Case "townpart"
            FormatOMMapNum = Replace(FormatOMMapNum, "0.", ".")
            'If Len(FormatOMMapNum) <> 3 Then FormatOMMapNum = "000"
        Case "towndir"
            If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "N"
        Case "range"
            If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "01"
        Case "rangepart"
            FormatOMMapNum = Replace(FormatOMMapNum, "0.", ".")
            'If Len(FormatOMMapNum) <> 3 Then FormatOMMapNum = "000"
            'If Len(sVal) <> 3 Then FormatOMMapNum = "000"
        Case "rangedir"
            If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "W"
        Case "section"
            If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "00"
        Case "qtr"
            If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "qtrqtr"
            If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "suffixtype"
            If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "suffixnum"
            If Len(FormatOMMapNum) <> 0 And Len(FormatOMMapNum) > 3 Then
                FormatOMMapNum = "000"
                GoTo Process_Exit
            ElseIf Len(FormatOMMapNum) = 1 Then
                FormatOMMapNum = "00" & FormatOMMapNum
                GoTo Process_Exit
            ElseIf Len(FormatOMMapNum) = 2 Then
                FormatOMMapNum = "0" & FormatOMMapNum
                GoTo Process_Exit
            End If
        Case "anomaly"
            If Len(FormatOMMapNum) > 2 Or Len(FormatOMMapNum) = 0 Then
                FormatOMMapNum = "00"
                GoTo Process_Exit
            ElseIf Len(FormatOMMapNum) = 2 Then
            
            ElseIf Len(FormatOMMapNum) = 1 Then
                FormatOMMapNum = "0" & FormatOMMapNum
                GoTo Process_Exit
            End If
        Case Else
            'Nothing to implement now
    End Select

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
    sVal1 = Right(sFullString, Len(sFullString) - (llow - 1))
    sVal2 = Left(sVal1, (lhigh - llow) + 1)
    ExtractString = sVal2

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
    Set pOC = obj.Class
    Set pDS = pOC
    If LCase(pDS.Name) = LCase(g_pFldnames.FCTaxlot) Then
        IsTaxlot = True
    Else
        IsTaxlot = False
    End If

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
    Set pOC = obj.Class
    Set pDS = pOC
    If LCase(pDS.Name) = LCase(g_pFldnames.FCMapIndex) Then
        IsMapIndex = True
    Else
        IsMapIndex = False
    End If


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

    IsAnno = False

    Dim pOC As IObjectClass
    Dim pDS As IDataset
    Set pOC = obj.Class
    Set pDS = pOC
    If TypeOf obj Is IFeature Then
        Dim pFC As IFeatureClass
        Set pFC = pOC
        If pFC.FeatureType = esriFTAnnotation Then IsAnno = True
    End If


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
    Set pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If pTaxlotFlayer Is Nothing Then
        MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        GoTo Process_Exit
    End If
    Set pTaxlotFclass = pTaxlotFlayer.FeatureClass
    Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then
        MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
        GoTo Process_Exit
    End If
    Set pMIFclass = pMIFlayer.FeatureClass
    'Get fields needed to populate the form
    Dim lMIOMNum As Long
    Dim lTLOMNum As Long
    Dim lTLTaxlot As Long
    Dim sMIOMval As String
    Dim sTLOMval As String
    lMIOMNum = modUtils.LocateFields(pMIFclass, g_pFldnames.MIORMAPMapNumberFN)
    lTLOMNum = modUtils.LocateFields(pTaxlotFclass, g_pFldnames.TLOrmapMapNumberFN)
    lTLTaxlot = modUtils.LocateFields(pTaxlotFclass, g_pFldnames.TLTaxlotFN)
    sMIOMval = GetValueViaOverlay(pGeometry, pMIFclass, g_pFldnames.MIORMAPMapNumberFN)
    'if no Mapindex or ORMAP mapnum, then no need to continue
    If sMIOMval = "" Then
        ValidateTaxlotNum = True
        GoTo Process_Exit
    End If
    'Make sure this number is unique within taxlots with this OM number
    Dim pCursor As ICursor
    Dim sWhere As String
    sWhere = g_pFldnames.TLOrmapMapNumberFN & " = '" & sMIOMval & _
            "' and " & g_pFldnames.TLTaxlotFN & " = '" & sEnteredTLval & "'"
    Set pCursor = AttributeQuery(pTaxlotFclass, sWhere)
    If Not pCursor Is Nothing Then
        Dim pRow As IRow
        Set pRow = pCursor.NextRow
        If Not pRow Is Nothing Then
            ValidateTaxlotNum = False
        Else
            ValidateTaxlotNum = True
        End If
    Else
        ValidateTaxlotNum = True
    End If

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
    
    Set pTaxlotFclass = pFeat.Class
    'Find MapIndex
    Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then
        response = MsgBox("Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex & ".  " & _
        "Load " & g_pFldnames.FCMapIndex & " automatically?", vbYesNo)
        If response <> vbYes Then GoTo Process_Exit
        modUtils.LoadFCIntoMap g_pFldnames.FCMapIndex, pTaxlotFclass
        'Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        If pMIFlayer Is Nothing Then GoTo Process_Exit
    End If

    'Find all fields needed
    m_bContinue = True
    lOMTLNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLOrmapTaxlotFN)
    lOMNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLOrmapMapNumberFN)
    lMNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLMapNumberFN)
    lTLCntyFld = LocateFields(pTaxlotFclass, g_pFldnames.TLCountyFN)
    lTaxlotFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTaxlotFN)
    lTLTownFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTownFN)
    lTLTownPartFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTownPartFN)
    lTLTownDirFld = LocateFields(pTaxlotFclass, g_pFldnames.TLTownDirFN)
    lTLRangeFld = LocateFields(pTaxlotFclass, g_pFldnames.TLRangeFN)
    lTLRangePartFld = LocateFields(pTaxlotFclass, g_pFldnames.TLRangePartFN)
    lTLRangeDirFld = LocateFields(pTaxlotFclass, g_pFldnames.TLRangeDirFN)
    lTLSectNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSectNumberFN)
    lTLQtrFld = LocateFields(pTaxlotFclass, g_pFldnames.TLQtrFN)
    lTLQQFld = LocateFields(pTaxlotFclass, g_pFldnames.TLQtrQtrFN)
    lTLMapSufTypeFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSufTypeFN)
    lTLMapSufNumFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSufNumFN)
    lTLSpecInterestFld = LocateFields(pTaxlotFclass, g_pFldnames.TLSpecInterestFN)
    lTLMapTaxlotFld = LocateFields(pTaxlotFclass, g_pFldnames.TLMapTaxlotFN)
    lTLMapNumberFld = LocateFields(pTaxlotFclass, g_pFldnames.TLMapNumberFN)
    lTaxlotMapAcres = LocateFields(pTaxlotFclass, g_pFldnames.TLMapAcresFN)
    lTLAnomalyFld = LocateFields(pTaxlotFclass, g_pFldnames.TLAnomalyFN)
    If Not m_bContinue Then GoTo Process_Exit 'If any fields not found

    'Obtain the map index poly via overlay
    Dim sExistVal As String
    Dim pArea As IArea
    Dim pCenter As IPoint
    Dim sExistOMMapNum As String
    Dim sExistMapNum As String
    
    Set pArea = pFeat.Shape
    Set pCenter = pArea.Centroid
    
    'Update Acreage
    pFeat.Value(lTaxlotMapAcres) = pArea.Area / 43560  '(pFeat.Value(lTaxlotShapeArea) / 43560)
    'Return the OMMapNum and MapNum and insert values into Taxlot
    sExistOMMapNum = GetValueViaOverlay(pCenter, pMIFlayer.FeatureClass, g_pFldnames.MIORMAPMapNumberFN)
    If sExistOMMapNum = "" Then GoTo Process_Exit 'If no value for whatever reason, don't continue
    sExistMapNum = GetValueViaOverlay(pCenter, pMIFlayer.FeatureClass, g_pFldnames.MIMapNumberFN)
    If sExistMapNum = "" Then GoTo Process_Exit 'If no value for whatever reason, don't continue
    'Store individual components of map number in taxlot
    pFeat.Value(lOMNumFld) = sExistOMMapNum
    pFeat.Value(lMNumFld) = sExistMapNum
    
    'County
    sExistVal = ParseOMMapNum(sExistOMMapNum, "county")
    sExistVal = ConvertCode(pFeat, g_pFldnames.TLCountyFN, sExistVal)
    pFeat.Value(lTLCntyFld) = CInt(sExistVal) 'Store county in county field
    
    'Town
    sExistVal = ParseOMMapNum(sExistOMMapNum, "town")
    pFeat.Value(lTLTownFld) = CInt(sExistVal)

    'TownPart
    sExistVal = ParseOMMapNum(sExistOMMapNum, "townpart")
    pFeat.Value(lTLTownPartFld) = CDbl(sExistVal)

    'TownDir
    sExistVal = ParseOMMapNum(sExistOMMapNum, "towndir")
    pFeat.Value(lTLTownDirFld) = sExistVal

    'Range
    sExistVal = ParseOMMapNum(sExistOMMapNum, "range")
    pFeat.Value(lTLRangeFld) = CInt(sExistVal)

    'RangePart
    sExistVal = ParseOMMapNum(sExistOMMapNum, "rangepart")
    pFeat.Value(lTLRangePartFld) = CDbl(sExistVal)

    'RangeDir
    sExistVal = ParseOMMapNum(sExistOMMapNum, "rangedir")
    pFeat.Value(lTLRangeDirFld) = sExistVal

    'Section
    sExistVal = ParseOMMapNum(sExistOMMapNum, "section")
    pFeat.Value(lTLSectNumFld) = CInt(sExistVal)
 
    'Qtr
    sExistVal = ParseOMMapNum(sExistOMMapNum, "qtr")
    pFeat.Value(lTLQtrFld) = sExistVal
    
    'QtrQtr
    sExistVal = ParseOMMapNum(sExistOMMapNum, "qtrqtr")
    pFeat.Value(lTLQQFld) = sExistVal

    'MapSuffixType
    sExistVal = ParseOMMapNum(sExistOMMapNum, "suffixtype")
    sExistVal = ConvertCode(pFeat, g_pFldnames.TLSufTypeFN, sExistVal)
    pFeat.Value(lTLMapSufTypeFld) = sExistVal
    
    'MapSuffixNum
    sExistVal = ParseOMMapNum(sExistOMMapNum, "suffixnum")
    pFeat.Value(lTLMapSufNumFld) = sExistVal
    
    'Anomaly
    sExistVal = ParseOMMapNum(sExistOMMapNum, "anomaly")
    pFeat.Value(lTLAnomalyFld) = sExistVal
    
    'SpecialInterest
    sExistVal = IIf(IsNull(pFeat.Value(lTLSpecInterestFld)), "00000", pFeat.Value(lTLSpecInterestFld))
    If Len(sExistVal) < 5 Then
     Do Until Len(sExistVal) = 5
        sExistVal = "0" & sExistVal
     Loop
    End If
    pFeat.Value(lTLSpecInterestFld) = sExistVal
    
    'Recalculate OMTaxlot
    If IsNull(pFeat.Value(lTaxlotFld)) Then GoTo Process_Exit
    Dim sTaxlotVal As String
    'Taxlot has actual taxlot number.  ORMAPTaxlot requires a 5-digit number, so leading zeros have to be added
    sTaxlotVal = pFeat.Value(lTaxlotFld)
    sTaxlotVal = AddLeadingZeros(sTaxlotVal, 5)
    Dim sNewOMTLNum As String
    Dim sExistOMTLNum As String
    If IsNull(pFeat.Value(lOMTLNumFld)) Then GoTo Process_Exit
    sExistOMTLNum = pFeat.Value(lOMTLNumFld)
    sNewOMTLNum = CalcOMTLNum(sExistOMTLNum, pFeat, sTaxlotVal)
    'If no changes, don't save value
    If Not sExistOMTLNum = sNewOMTLNum Then
        pFeat.Value(lOMTLNumFld) = sNewOMTLNum
    End If
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

        If Len(asCurString) < lWidth Then
         Do Until Len(asCurString) = lWidth
            asCurString = "0" & asCurString
         Loop
        End If
        AddLeadingZeros = asCurString

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

        If pFeat.FeatureType = esriFTAnnotation Or pFeat.FeatureType = esriFTDimension Then
            Dim pArea As IArea
            Set pArea = pFeat.Shape
            Set GetCentroid = pArea.Centroid
        End If

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
    Set pCenter = New Point
    pCenter.X = pEnv.XMin + (pEnv.XMax - pEnv.XMin) / 2
    pCenter.Y = pEnv.YMin + (pEnv.YMax - pEnv.YMin) / 2
    Set CT_GetCenterOfEnvelope = pCenter

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
    
    Set pEnumRelClass = pObj.Class.RelationshipClasses(esriRelRoleAny)
    If Not pEnumRelClass Is Nothing Then
      Set pRelClass = pEnumRelClass.Next
      If Not pRelClass Is Nothing Then
          Set pParentSet = pRelClass.GetObjectsRelatedToObject(pObj)
      End If
    Else
        GoTo Process_Exit
    End If
    If Not pParentSet Is Nothing Then
        Set pParentFeat = pParentSet.Next
        If Not pParentFeat Is Nothing Then
            Set GetRelatedObjects = pParentFeat
        End If
    End If

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
    With g_pFldnames
'++START JWM 10/11/2006 using strcomp function
        If StrComp(sFCName, .FCTLAcrAnno, vbTextCompare) = 0 Then
        'Determine anno size based on scale
        If lScale = 120 Then dSize = .AnnoSizeTLAcr120
        If lScale = 240 Then dSize = .AnnoSizeTLAcr240
        If lScale = 360 Then dSize = .AnnoSizeTLAcr360
        If lScale = 480 Then dSize = .AnnoSizeTLAcr480
        If lScale = 600 Then dSize = .AnnoSizeTLAcr600
        If lScale = 1200 Then dSize = .AnnoSizeTLAcr1200
        If lScale = 2400 Then dSize = .AnnoSizeTLAcr2400
        If lScale = 4800 Then dSize = .AnnoSizeTLAcr4800
        If lScale = 9600 Then dSize = .AnnoSizeTLAcr9600
        If lScale = 24000 Then dSize = .AnnoSizeTLAcr24000
'++END JWM 10/11/2006
      ElseIf StrComp(sFCName, .FCTLNumAnno, vbTextCompare) = 0 Then
        If lScale = 120 Then dSize = .AnnoSizeTLNum120
        If lScale = 240 Then dSize = .AnnoSizeTLNum240
        If lScale = 360 Then dSize = .AnnoSizeTLNum360
        If lScale = 480 Then dSize = .AnnoSizeTLNum480
        If lScale = 600 Then dSize = .AnnoSizeTLNum600
        If lScale = 1200 Then dSize = .AnnoSizeTLNum1200
        If lScale = 2400 Then dSize = .AnnoSizeTLNum2400
        If lScale = 4800 Then dSize = .AnnoSizeTLNum4800
        If lScale = 9600 Then dSize = .AnnoSizeTLNum9600
        If lScale = 24000 Then dSize = .AnnoSizeTLNum24000
      Else
        'Something not being trapped
        dSize = 10
      End If
    End With
    '++end new code
    'TODO #####
    'Determine a default
    If dSize = 0 Then dSize = 5
    GetAnnoSizeByScale = dSize

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
    Set pFSO = CreateObject("Scripting.FileSystemObject")
    If Not pFSO.FileExists(sPath) Then
        FileExists = False
    Else
        FileExists = True
    End If


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

Set rot = New AppROT
If rot.Count = 1 Then
    Set app = rot.Item(0) 'ArcCatalog
Else
    Set app = rot.Item(1) 'ArcMap
End If
Set pobjectFactory = app

Set GetAppRef = app


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

Set rot = New AppROT
If rot.Count = 1 Then
    Set app = rot.Item(0) 'ArcCatalog
Else
    Set app = rot.Item(1) 'ArcMap
End If
Set pobjectFactory = app
Set pMXDoc = app.Document

Set GetMXDocRef = pMXDoc


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
    Set pWS = pOtherFC.FeatureDataset.Workspace
    Set pFWS = pWS
    Set pFC = pFWS.OpenFeatureClass(sFCName)
    Set pFeatLayer = New FeatureLayer
    Set pFeatLayer.FeatureClass = pFC
    Set pDataset = pFC
    pFeatLayer.Name = pDataset.Name
    Set pMXDoc = g_pApp.Document
    Set pMap = pMXDoc.FocusMap
    pMap.AddLayer pFeatLayer
    pMXDoc.CurrentContentsView.Refresh 0


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
    Set pOC = obj.Class
    Set pDSet = pOC
    pName = LCase(Trim(pDSet.Name))
    If pName = LCase(Trim(g_pFldnames.FCAnno10)) Or pName = LCase(Trim(g_pFldnames.FCAnno100)) Or pName = LCase(Trim(g_pFldnames.FCAnno20)) Or _
        pName = LCase(Trim(g_pFldnames.FCAnno200)) Or pName = LCase(Trim(g_pFldnames.FCAnno2000)) Or pName = LCase(Trim(g_pFldnames.FCAnno30)) Or _
        pName = LCase(Trim(g_pFldnames.FCAnno40)) Or pName = LCase(Trim(g_pFldnames.FCAnno400)) Or pName = LCase(Trim(g_pFldnames.FCAnno50)) Or _
        pName = LCase(Trim(g_pFldnames.FCAnno800)) Or pName = LCase(Trim(g_pFldnames.FCCartoLines)) Or pName = LCase(Trim(g_pFldnames.FCLotsAnno)) Or _
        pName = LCase(Trim(g_pFldnames.FCMapIndex)) Or pName = LCase(Trim(g_pFldnames.FCPlats)) Or pName = LCase(Trim(g_pFldnames.FCReferenceLines)) Or _
        pName = LCase(Trim(g_pFldnames.FCTaxCode)) Or pName = LCase(Trim(g_pFldnames.FCTaxCode)) Or pName = LCase(Trim(g_pFldnames.FCTaxCodeAnno)) Or _
        pName = LCase(Trim(g_pFldnames.FCTaxlot)) Or pName = LCase(Trim(g_pFldnames.FCTaxlotLines)) Or pName = LCase(Trim(g_pFldnames.FCTLAcrAnno)) Or _
        pName = LCase(Trim(g_pFldnames.FCTLNumAnno)) Then
    Else
        IsOrMapFeature = False
    End If


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
    Set pAOC = obj.Class
    Set pAnnoFeat = obj
    
    'Capture MapNumber for each anno feature created
    lAnnoMapNumFld = LocateFields(obj.Class, g_pFldnames.MIMapNumberFN)
    If lAnnoMapNumFld = -1 Then GoTo Process_Exit
    

    'If new anno feature with no text, determine if it has a shape
    Dim lFld As Long
    lFld = pAnnoFeat.Fields.FindField("TextString")
    If lFld = -1 Then
        MsgBox "Unable to locate textstring field in anno class.  Cannot set size", vbCritical
        GoTo Process_Exit
    End If
    Dim vVal As Variant
    vVal = pAnnoFeat.Value(lFld)
    If IsNull(vVal) Then GoTo Process_Exit
        
    
    Set pFeat = obj
    Set pGeometry = pFeat.Shape
    If pGeometry.IsEmpty Then GoTo Process_Exit
    Set pEnv = pGeometry.Envelope
    Set pCenter = CT_GetCenterOfEnvelope(pEnv)
    Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then GoTo Process_Exit
    Set pMIFclass = pMIFlayer.FeatureClass
    sMapNum = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
    
    'Allow existing anno to be moved without changing MapNumber
    'Some anno will reside in another Taxlot, but labels the neighboring taxlot
    If sMapNum = obj.Value(lAnnoMapNumFld) Then
        obj.Value(lAnnoMapNumFld) = sMapNum
    
        'Update the size to reflect current mapscale
        sMapScale = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
        If IsNull(sMapScale) Then GoTo Process_Exit
        
        'Determine which annotation class this is
        Set pAnnoClass = obj.Class
        Set pAnnoDset = pAnnoClass
        'If other anno, don't continue
        If LCase(pAnnoDset.Name) <> LCase(g_pFldnames.FCTLAcrAnno) And LCase(pAnnoDset.Name) <> LCase(g_pFldnames.FCTLNumAnno) Then
            GoTo Process_Exit
        End If
        
        dSize = modUtils.GetAnnoSizeByScale(pAnnoDset.Name, CLng(sMapScale))
        'Get the anno feature, its symbol, set the appropriate size
        Set pAnnotationFeature = obj
        Set pAnnotationElement = pAnnotationFeature.Annotation
        Set pElement = pAnnotationElement
        Set pTextElement = pElement
        Set pTextSym = pTextElement.Symbol
        pTextSym.Size = dSize
        pTextElement.Symbol = pTextSym
        Set pElement = pTextElement
        Set pAnnotationElement = pElement
        pAnnotationFeature.Annotation = pAnnotationElement
    End If
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
    lFld = pFClass.Fields.FindField(pFldName)
    If lFld > -1 Then
      LocateFields = lFld
    Else
        MsgBox "Unable to locate " & pFldName & " field in " & _
        pFClass.AliasName & " feature class"
    End If


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
    lAutoDateFld = pFeat.Fields.FindField(g_pFldnames.AutoDateFN)
    If lAutoDateFld > -1 Then
        pFeat.Value(lAutoDateFld) = Now
    End If
    lAutoWhoFld = pFeat.Fields.FindField(g_pFldnames.AutoWhoFN)
    If lAutoWhoFld > -1 Then
        pFeat.Value(lAutoWhoFld) = Environ("USERNAME")
    End If


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
            If Len(sString) < 5 Then
             Do Until Len(sString) = 5
                sString = "0" & sString
             Loop
            End If
            Validate5Digits = sString

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
        lTLSpecInterestFld = modUtils.LocateFields(pFeature.Class, g_pFldnames.TLSpecInterestFN)
        If lTLSpecInterestFld = -1 Then
            sTLSpecVAl = "00000"
        Else
            If Not IsNull(pFeature.Value(lTLSpecInterestFld)) Then
                sTLSpecVAl = pFeature.Value(lTLSpecInterestFld)
            Else
                sTLSpecVAl = "00000"
            End If
            'Verify that it is 5 digits
            If Len(sTLSpecVAl) < 5 Then
             Do Until Len(sTLSpecVAl) = 5
                sTLSpecVAl = "0" & sTLSpecVAl
             Loop
            End If
        End If
        GetSpecialInterests = sTLSpecVAl

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
        lTLMapSufTypeFld = modUtils.LocateFields(pFeature.Class, g_pFldnames.TLSufTypeFN)
        If lTLMapSufTypeFld = -1 Then
            sTLMapSufTypeVAl = "0"
        Else
            If Not IsNull(pFeature.Value(lTLMapSufTypeFld)) Then
                sTLMapSufTypeVAl = pFeature.Value(lTLMapSufTypeFld)
            Else
                sTLMapSufTypeVAl = "0"
            End If
                'Verify that it is 1 digit
                If Len(sTLMapSufTypeVAl) < 1 Then
                    Do Until Len(sTLMapSufTypeVAl) = 1
                       sTLMapSufTypeVAl = "0" & sTLMapSufTypeVAl
                    Loop
                End If

                'Verify that it isn't more than 1 digit
                If Len(sTLMapSufTypeVAl) > 1 Then
                    sTLMapSufTypeVAl = Left(sTLMapSufTypeVAl, 1)
                End If
            End If


        GetMapSufType = sTLMapSufTypeVAl

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
        lTLMapSufNumFld = modUtils.LocateFields(pFeature.Class, g_pFldnames.TLSufNumFN)
        If lTLMapSufNumFld = -1 Then
            sTLMapSufNumVAl = "000"
        Else
            If Not IsNull(pFeature.Value(lTLMapSufNumFld)) Then
                sTLMapSufNumVAl = pFeature.Value(lTLMapSufNumFld)
            Else
                sTLMapSufNumVAl = "000"
            End If
                'Verify that it is 3 digit
                If Len(sTLMapSufNumVAl) < 3 Then
                    Do Until Len(sTLMapSufNumVAl) = 3
                       sTLMapSufNumVAl = "0" & sTLMapSufNumVAl
                    Loop
                End If

                'Verify that it isn't more than 3 digits
                If Len(sTLMapSufNumVAl) > 3 Then
                    sTLMapSufNumVAl = Left(sTLMapSufNumVAl, 3)
                End If
            End If


        GetMapSufNum = sTLMapSufNumVAl

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

        sShortOMNum = ShortenOMMapNum(sExistOMNum)
              '++ BEGIN, Laura Gordon, Novemeber 29, 2005
              '+sTLSpecVAl = GetSpecialInterests(pFeat)
              '+sOMTLNval = sShortOMNum & sTLVal & sTLSpecVAl
              sTLMapSufTypeVAl = GetMapSufType(pFeat)
              sTLMapSufNumVAl = GetMapSufNum(pFeat)
              sOMTLNval = sShortOMNum & sTLMapSufTypeVAl & sTLMapSufNumVAl & sTLVal
              '++ END, Laura Gordon, Novemeber 29, 2005
        CalcOMTLNum = sOMTLNval

  Exit Function
ErrorHandler:
  HandleError True, "CalcOMTLNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Function ShortenOMMapNum(sOMVal As String) As String
  On Error GoTo ErrorHandler

    'Remove two values from the ORMAPMap number for the purpose of populating ORMAPTaxlog
    ShortenOMMapNum = Left(sOMVal, 20)

  Exit Function
ErrorHandler:
  HandleError True, "ShortenOMMapNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Public Sub ZoomToExtent(pEnv As IEnvelope, pMXDoc As IMxDocument)
    'Zooms the current extent to the passed in envelope (i.e. zoom to feature)
    'Works for Layout and Data view
    Dim pMap As IMap
    Dim pActiveView As IActiveView
    Set pMap = pMXDoc.FocusMap
    Set pActiveView = pMap

    pActiveView.Extent = pEnv
    pActiveView.Refresh
End Sub

