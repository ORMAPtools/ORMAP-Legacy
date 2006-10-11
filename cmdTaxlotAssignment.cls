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


  Set m_pBitmap = LoadResPicture(101, 0)
  Set m_pCursor = LoadResPicture(101, 2)


  Exit Sub
ErrorHandler:
  HandleError True, "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Class_Terminate()
  On Error GoTo ErrorHandler

  
  Unload frmTaxlotAssignment
  Set m_pBitmap = Nothing
  Set m_pCursor = Nothing
  Set m_pBitmap = Nothing
  Set m_pCursor = Nothing
  Set m_pApp = Nothing
  Set m_pDoc = Nothing
  Set m_pMap = Nothing
  Set m_pTaxlotFlayer = Nothing
  Set m_pTaxlotFClass = Nothing
  Set m_pMIFlayer = Nothing
  Set m_pMIFclass = Nothing
  Set m_pExt = Nothing
  Set m_pEditorEvents = Nothing
  Set m_pEditor = Nothing


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


  If m_bEnabled = True Then
    ICommand_Enabled = True
  End If


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler


     ICommand_Name = "TaxlotAssignment"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler


     ICommand_Caption = "Taxlot Assignment"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler


     ICommand_Tooltip = "Taxlot Assignment Tool"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
  On Error GoTo ErrorHandler

    
  ICommand_Message = "Set starting value and increment value to populate values in Taxlots"


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

    
   ICommand_Bitmap = m_pBitmap


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Category() As String
  On Error GoTo ErrorHandler


     ICommand_Category = "ORMAP"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
  On Error GoTo ErrorHandler

    
  Set m_pApp = hook
  If g_pApp Is Nothing Then
    Set g_pApp = hook
  End If
  Set m_pDoc = m_pApp.Document
  Set m_pMap = m_pDoc.FocusMap
  
  'Get a reference to the editor extension
  Dim pID As New UID
  pID = "esriCore.Editor"
  Set m_pApp = hook
  Set m_pEditor = m_pApp.FindExtensionByCLSID(pID)
  Set m_pEditorEvents = m_pEditor


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
    Set m_pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If m_pTaxlotFlayer Is Nothing Then
        MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        GoTo Process_Exit
    End If
    Set m_pTaxlotFClass = m_pTaxlotFlayer.FeatureClass
    'Find MapIndex
    Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If m_pMIFlayer Is Nothing Then
        response = MsgBox("Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex & ".  " & _
        "Load " & g_pFldnames.FCMapIndex & " automatically?", vbYesNo)
        If response <> vbYes Then GoTo Process_Exit
        modUtils.LoadFCIntoMap g_pFldnames.FCMapIndex, m_pTaxlotFClass
        Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        If m_pMIFlayer Is Nothing Then GoTo Process_Exit
    End If
    Set m_pMIFclass = m_pMIFlayer.FeatureClass
    m_lOMTLNumFld = m_pTaxlotFClass.FindField(g_pFldnames.TLOrmapTaxlotFN)
    If m_lOMTLNumFld = -1 Then
        MsgBox "Unable to locate ORMAPTaxlot field.  Check ini file"
        GoTo Process_Exit
    End If
    m_lOMNumFld = m_pTaxlotFClass.FindField(g_pFldnames.TLOrmapMapNumberFN)
    If m_lOMNumFld = -1 Then
        MsgBox "Unable to locate ORMAPMapnumber field.  Check ini file"
        GoTo Process_Exit
    End If
    m_lTLTaxlotFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLTaxlotFN)
    m_lTLAnomalyFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLAnomalyFN)
    'Initialize the TaxlotAssignment form
    frmTaxlotAssignment.Show vbModal

Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
  On Error GoTo ErrorHandler

    
    ITool_Cursor = m_pCursor
    

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
    
  If Button = vbLeftButton Then
    'Save values from Form to variables and increment value in form
    Dim lvalue As Long
    Dim lincrement As Long
    Dim lNextValue As Long
    Dim sFieldName As String

    'If "NUMBER" selected, then make sure value is numeric
    If StrComp(frmTaxlotAssignment.cmbTaxlotNum.Text, "NUMBER", vbTextCompare) = 0 Then
        lvalue = frmTaxlotAssignment.txtTaxlotNum.Text
        If frmTaxlotAssignment.tglBy1.Value = True Then
            lincrement = 1
        ElseIf frmTaxlotAssignment.tglBy10.Value = True Then
            lincrement = 10
        Else
            lincrement = 100
        End If
    Else 'If one of the non-numeric categories is selected
        ' dont have a else if there is nothing to do

    End If

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

    Set pMap = m_pEditor.Map
    Set pPoint = m_pEditor.Display.DisplayTransformation.ToMapPoint(X, Y)
    Set pGeometry = m_pEditor.CreateSearchShape(pPoint)
    If frmTaxlotAssignment.cmbTaxlotNum.Text = "NUMBER" Then
        'Verify that within this map index, this taxlot number is unique
        'If not unique, prompt user to enter a new value
        If Not modUtils.ValidateTaxlotNum(frmTaxlotAssignment.txtTaxlotNum.Text, pGeometry) Then
            MsgBox "The current Taxlot value (" & frmTaxlotAssignment.txtTaxlotNum.Text & _
            ") is not unique withing this MapIndex.  Please enter a new number"
            frmTaxlotAssignment.Show vbModal
            GoTo Process_Exit
        End If
    End If
    Set pEnv = pGeometry 'QI
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pGeometry
    strShapeFieldName = m_pTaxlotFClass.ShapeFieldName
    Set pSpatialFilter.OutputSpatialReference(strShapeFieldName) = pMap.SpatialReference
    pSpatialFilter.GeometryField = m_pTaxlotFClass.ShapeFieldName
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    Set pFCursor = m_pTaxlotFClass.Search(pSpatialFilter, False)
    If pFCursor Is Nothing Then GoTo Process_Exit 'If no taxlots selected
    Set pFeature = pFCursor.NextFeature
    If Not pFeature Is Nothing Then
      'Update the feature
        m_pEditor.StartOperation
        Dim sOMTLNval As String
        Dim sExistOMNum As String
        sExistOMNum = IIf(IsNull(pFeature.Value(m_lOMNumFld)), "", pFeature.Value(m_lOMNumFld))
        If sExistOMNum = "" Then 'Obtain it from MapIndex polygon
            modUtils.CalcTaxlotValues pFeature, m_pMIFlayer
            sExistOMNum = IIf(IsNull(pFeature.Value(m_lOMNumFld)), "", pFeature.Value(m_lOMNumFld))
            If sExistOMNum = "" Then
                MsgBox "ORMAPMapNumber not present in this taxlot or MapIndex.  Use the MapIndex tool to create one"
                GoTo Process_Exit
            End If
        End If
        'Assign Taxlot value
        Dim sExistTLNumVal As String
        sExistTLNumVal = IIf(IsNull(pFeature.Value(m_lTLTaxlotFld)), "", pFeature.Value(m_lTLTaxlotFld))
        If Not sExistTLNumVal = "" Then
            Dim response As Integer
            response = MsgBox("Taxlot currently has a Taxlot value (" & sExistTLNumVal & ").  Update it?", vbYesNo)
            If response = vbNo Then GoTo Process_Exit
        End If
        'Taxlot can be less than 5-digits
        'The Taxlot value in OrMapMapNum must be 5 digits.
        'Two versions of the taxlot number will be used for these purposes.
        Dim sNewTLNum As String
        Dim sNewTLNum_5digit As String
        If StrComp(frmTaxlotAssignment.cmbTaxlotNum.Text, "NUMBER", vbTextCompare) = 0 Then
            sNewTLNum = frmTaxlotAssignment.txtTaxlotNum.Text 'User entered number
            sNewTLNum_5digit = sNewTLNum
            'Make sure number is 5 characters
            If Len(sNewTLNum_5digit) < 5 Then
             Do Until Len(sNewTLNum_5digit) = 5
                sNewTLNum_5digit = "0" & sNewTLNum_5digit
             Loop
            End If
        Else
            'Remove leading Zeros for taxlot number if any exist
            sNewTLNum_5digit = frmTaxlotAssignment.cmbTaxlotNum.Text 'Predefined selection
            'sNewTLNum_5digit = sNewTLNum
            sNewTLNum = Replace(sNewTLNum_5digit, "0", "")
        End If
        '++ BEGIN, Laura Gordon, November 29, 2005
        'Determine if Special Interests field is something other than default
        'If so, include it in ORMAPtaxlot
        '+Dim sTLSpecVAl As String
        '+sTLSpecVAl = GetSpecialInterests(pFeature)
        '+If IsNull(sTLSpecVAl) Then sTLSpecVAl = "00000"
        Dim sTLMapSufTypeVAl As String
        Dim sTLMapSufNumVAl As String
        sTLMapSufTypeVAl = GetMapSufType(pFeature)
        sTLMapSufNumVAl = GetMapSufNum(pFeature)
        If IsNull(sTLMapSufTypeVAl) Then sTLMapSufTypeVAl = "0"
        If IsNull(sTLMapSufNumVAl) Then sTLMapSufNumVAl = "000"
        '++ END, Laura Gordon, November 29, 2005

        pFeature.Value(m_lTLTaxlotFld) = sNewTLNum
        Dim sShortOMNum As String 'Remove suffixTYpe and suffixNum
        sShortOMNum = ShortenOMMapNum(sExistOMNum)
        '++ BEGIN, Laura Gordon, November 29, 2005
        '+sOMTLNval = sShortOMNum & sNewTLNum_5digit & sTLSpecVAl '"00000"
        sOMTLNval = sShortOMNum & sTLMapSufTypeVAl & sTLMapSufNumVAl & sNewTLNum_5digit
        '++ END, Laura Gordon, November 29, 2005

        'Copy Anomaly from MapIndex
        Dim sAnom As String
        sAnom = ParseOMMapNum(sExistOMNum, "anomaly")
        pFeature.Value(m_lTLAnomalyFld) = sAnom
        'Assign OrmapTaxlot value
        pFeature.Value(m_lOMTLNumFld) = sOMTLNval
        pFeature.Store
        'AutoIncrement if necessary
        If frmTaxlotAssignment.tglAutoYes.Value = True And StrComp(frmTaxlotAssignment.cmbTaxlotNum.Text, "NUMBER", vbTextCompare) = 0 Then
            lNextValue = lvalue + lincrement
            'Add leading zeros if necessary
            Dim sNextValue As String
            sNextValue = CStr(lNextValue)
            'If Len(sNextValue) < 5 Then
             'Do Until Len(sNextValue) = 5
            ''    sNextValue = "0" & sNextValue
             'Loop
           ' End If
            frmTaxlotAssignment.txtTaxlotNum.Text = lNextValue
        End If
      m_pEditor.StopOperation "AutoIncrement Attribute"
    If frmTaxlotAssignment.tglAutoYes.Value = False Then
        'Deactivate the tool
        Set m_pApp.CurrentTool = Nothing
    End If
    Else
        'Deactivate the tool
        Set m_pApp.CurrentTool = Nothing
        GoTo Process_Exit 'Don't do anything if no feature
    End If
    Set pActiveView = pMap
    pActiveView.PartialRefresh esriViewGeography, m_pTaxlotFlayer, pFeature.Extent.Envelope
  End If
    
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

    
  If keyCode = vbKeyO Then
    frmTaxlotAssignment.Show vbModal
  End If

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

    
    ITool_Deactivate = True


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
    UpdateAutoFields obj
        
    If Not modUtils.IsOrMapFeature(obj) Then Exit Sub
    Dim u As New UID
    Dim pFeat As IFeature
    
    If modUtils.IsTaxlot(obj) Then
        'Obtain ORMAPMapnum via overlay and calculate other field values
        modUtils.CalcTaxlotValues obj, m_pMIFlayer
    ElseIf modUtils.IsAnno(obj) Then
        Dim pAnnotationFeature As IAnnotationFeature
        Set pAnnotationFeature = obj
        
        'Get the parent feature so mapnumber can be obtained
        Dim lParentID As Long
        lParentID = pAnnotationFeature.LinkedFeatureID
        If lParentID > -1 Then 'Feature linked
            Set pFeat = GetRelatedObjects(obj)
            If pFeat Is Nothing Then Exit Sub
        Else
            'Not feature linked anno, so we can use the feature as is
            Set pFeat = obj
        End If
    
        'Set anno size
        SetAnnoSize obj, pFeat

    End If

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
    UpdateAutoFields obj
        
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

    lMapNumFld = obj.Fields.FindField(g_pFldnames.MIMapNumberFN)
    lMapScaleFld = obj.Fields.FindField(g_pFldnames.MIMapScaleFN)
    If modUtils.IsTaxlot(obj) Then
        'Obtain ORMAPMapnum via overlay and calculate other field values
        modUtils.CalcTaxlotValues obj, m_pMIFlayer
        
    ElseIf modUtils.IsAnno(obj) Then
        'Capture MapNumber for each anno feature created
        lAnnoMapNumFld = LocateFields(obj.Class, g_pFldnames.MIMapNumberFN)
        If lAnnoMapNumFld = -1 Then GoTo Process_Exit
        Set pAnnotationFeature = obj
        
        'Get the parent feature so mapnumber can be obtained
        lParentID = pAnnotationFeature.LinkedFeatureID
        If lParentID > -1 Then 'Feature linked
            Set pFeat = GetRelatedObjects(obj)
            If pFeat Is Nothing Then GoTo Process_Exit
        Else
            'Not feature linked anno, so we can use the feature as is
            Set pFeat = obj
        End If

        Set pGeometry = pFeat.Shape
        If pGeometry.IsEmpty Then GoTo Process_Exit
        Set pEnv = pGeometry.Envelope
        Set pCenter = CT_GetCenterOfEnvelope(pEnv)
        Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        If pMIFlayer Is Nothing Then GoTo Process_Exit
        Set pMIFclass = pMIFlayer.FeatureClass
        sMapNum = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
        obj.Value(lAnnoMapNumFld) = sMapNum
        If lMapScaleFld > -1 Then
            vMapScaleVal = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
            If Not IsNull(vMapScaleVal) Then
                obj.Value(lMapScaleFld) = vMapScaleVal
            End If
        End If
        'Set size based on mapscale
        SetAnnoSize obj, pFeat
    Else
        'Update MapScale for all features with a MapScale field (except MapIndex)
        If lMapScaleFld > -1 And Not modUtils.IsMapIndex(obj) Then
            Set pFeat = obj
            Set pGeometry = pFeat.Shape
            If pGeometry.IsEmpty Then GoTo Process_Exit
            Set pEnv = pGeometry.Envelope
            Set pCenter = CT_GetCenterOfEnvelope(pEnv)
            Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
            If pMIFlayer Is Nothing Then GoTo Process_Exit
            Set pMIFclass = pMIFlayer.FeatureClass
            vMapScaleVal = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
            If Not IsNull(vMapScaleVal) Then
                obj.Value(lMapScaleFld) = vMapScaleVal
            End If
            'If a dataset with MapNumber, populate it
            If lMapNumFld > -1 Then
                vMapNumVal = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
                If Not IsNull(vMapNumVal) Then
                    obj.Value(lMapNumFld) = vMapNumVal
                End If
            End If
        End If
    End If
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
    If Not IsOrMapFeature(obj) Then GoTo Process_Exit
    If modUtils.IsTaxlot(obj) Then
            Dim pDSet As IDataset
            Dim pCNTable As ITable
            Dim pWSpace As IWorkspace
            Dim pFWSpace As IFeatureWorkspace
            Dim pFeat As IFeature
            Dim pRow As IRow
            Set pFeat = obj
            Set m_pTaxlotFClass = pFeat.Class
            Set pDSet = m_pTaxlotFClass
            Set pWSpace = pDSet.Workspace
            Set pFWSpace = pWSpace
            Set pCNTable = pFWSpace.OpenTable(g_pFldnames.TabCancelled)
            If pCNTable Is Nothing Then
                'Could privide warning that table is missing, but could get annoying
                GoTo Process_Exit
            End If
            m_lTLTaxlotFld = m_pTaxlotFClass.FindField(g_pFldnames.TLTaxlotFN)
            m_lTLMNFld = m_pTaxlotFClass.FindField(g_pFldnames.TLMapNumberFN)
            m_lCNTaxlotFld = pCNTable.FindField(g_pFldnames.TLTaxlotFN)
            m_lCNMNumFld = pCNTable.FindField(g_pFldnames.TLMapNumberFN)
            If m_lCNTaxlotFld = -1 Or m_lCNMNumFld = -1 Then GoTo Process_Exit
            'If no null values, copy them to Cancelled numbers
            If Not IsNull(obj.Value(m_lTLTaxlotFld)) And Not IsNull(obj.Value(m_lTLMNFld)) Then
                Set pRow = pCNTable.CreateRow
                If pRow Is Nothing Then GoTo Process_Exit
                pRow.Value(m_lCNTaxlotFld) = obj.Value(m_lTLTaxlotFld)
                pRow.Value(m_lCNMNumFld) = obj.Value(m_lTLMNFld)
                pRow.Store
                Set pRow = Nothing
            End If
        End If
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnDeleteFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub m_pEditorEvents_OnStartEditing()
  On Error GoTo ErrorHandler

    
  m_bEnabled = True


  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnStartEditing " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub m_pEditorEvents_OnStopEditing(ByVal save As Boolean)
  On Error GoTo ErrorHandler


  m_bEnabled = False


  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnStopEditing " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




