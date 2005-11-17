VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private m_pExt As IExtensionConfig
Private WithEvents m_pEditorEvents As Editor
Attribute m_pEditorEvents.VB_VarHelpID = -1
Private m_pEditor As IEditor2
Private m_bContinue As Boolean
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument
Private m_pMap As IMap
Const c_sModuleFileName = "C:\active\ModelingWorkshop_01-05-05\CustomCode\ormap\clsAutoUpdate.cls"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms
Implements ICommand

Private Sub Class_Initialize()
  On Error GoTo ErrorHandler


  Exit Sub
ErrorHandler:
  HandleError True, "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub UpdateAutoFields(pFeat As IFeature)
  On Error GoTo ErrorHandler


'Code to update AutoDate and AutoWho
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
  HandleError False, "UpdateAutoFields " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



 
Private Sub Class_Terminate()
  On Error GoTo ErrorHandler

  Set m_pExt = Nothing
  Set m_pEditorEvents = Nothing
  Set m_pEditor = Nothing
  Set m_pMxDoc = Nothing
  Set m_pMap = Nothing

  Exit Sub
ErrorHandler:
  HandleError True, "Class_Terminate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




Private Sub m_pEditorEvents_OnChangeFeature(ByVal obj As esriGeoDatabase.IObject)
  On Error GoTo ErrorHandler

    If Not IsOrMapFeature(obj) Then Exit Sub
    Dim u As New UID
    Dim pFeat As IFeature
    
'    If m_pExt Is Nothing Then
'      u.Value = "TaxlotEditing.ORMAPEditingExt"
'      Set m_pExt = m_pApp.FindExtensionByCLSID(u)
'    End If
    UpdateAutoFields obj
    If IsTaxlot(obj) Then
        'Obtain ORMAPMapnum via overlay and calculate other field values
        modUtils.CalcTaxlotValues obj
    ElseIf IsAnno(obj) Then
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

Private Sub m_pEditorEvents_OnCreateFeature(ByVal obj As esriGeoDatabase.IObject)
  On Error GoTo ErrorHandler
  
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
    
    MsgBox "In Extension"
    UpdateAutoFields obj
    If IsTaxlot(obj) Then
        'Obtain ORMAPMapnum via overlay and calculate other field values
        modUtils.CalcTaxlotValues obj
        
    ElseIf IsAnno(obj) Then
        'Capture MapNumber for each anno feature created
        lAnnoMapNumFld = FindField(obj.Class, g_pFldnames.MIMapNumberFN)
        If lAnnoMapNumFld = -1 Then Exit Sub
        Set pAnnotationFeature = obj
        
        'Get the parent feature so mapnumber can be obtained
        lParentID = pAnnotationFeature.LinkedFeatureID
        If lParentID > -1 Then 'Feature linked
            Set pFeat = GetRelatedObjects(obj)
            If pFeat Is Nothing Then Exit Sub
        Else
            'Not feature linked anno, so we can use the feature as is
            Set pFeat = obj
        End If

        Set pGeometry = pFeat.Shape
        If pGeometry.IsEmpty Then Exit Sub
        Set pEnv = pGeometry.Envelope
        Set pCenter = CT_GetCenterOfEnvelope(pEnv)
        Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        If pMIFlayer Is Nothing Then Exit Sub
        Set pMIFclass = pMIFlayer.FeatureClass
        sMapNum = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
        obj.Value(lAnnoMapNumFld) = sMapNum
        
        'Set size based on mapscale
        SetAnnoSize obj, pFeat
    End If
    
  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnCreateFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function FindField(pFClass As IFeatureClass, pFldName As String) As Long
  On Error GoTo ErrorHandler


    Dim lFld As Long
    lFld = pFClass.Fields.FindField(pFldName)
    If lFld > -1 Then
      FindField = lFld
    Else
        MsgBox "Unable to locate " & pFldName & " field in " & _
        pFClass.AliasName & " feature class"
    End If


  Exit Function
ErrorHandler:
  HandleError True, "FindField " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function



Private Function IsOrMapFeature(obj As esriGeoDatabase.IObject) As Boolean
  On Error GoTo ErrorHandler

    Dim pOC As IObjectClass
    Dim pDSet As IDataset
    Dim pName As String
    Set pOC = obj.Class
    Set pDSet = pOC
    pName = pDSet.Name
    If pName = g_pFldnames.FCAnno10 Or pName = g_pFldnames.FCAnno100 Or pName = g_pFldnames.FCAnno20 Or _
        pName = g_pFldnames.FCAnno200 Or pName = g_pFldnames.FCAnno2000 Or pName = g_pFldnames.FCAnno30 Or _
        pName = g_pFldnames.FCAnno40 Or pName = g_pFldnames.FCAnno400 Or pName = g_pFldnames.FCAnno50 Or _
        pName = g_pFldnames.FCAnno800 Or pName = g_pFldnames.FCCartoLines Or pName = g_pFldnames.FCLotsAnno Or _
        pName = g_pFldnames.FCMapIndex Or pName = g_pFldnames.FCPlats Or pName = g_pFldnames.FCReferenceLines Or _
        pName = g_pFldnames.FCTaxCode Or pName = g_pFldnames.FCTaxCode Or pName = g_pFldnames.FCTaxCodeAnno Or _
        pName = g_pFldnames.FCTaxlot Or pName = g_pFldnames.FCTaxlotLines Or pName = g_pFldnames.FCTLAcrAnno Or _
        pName = g_pFldnames.FCTLNumAnno Then
        IsOrMapFeature = True
    Else
        IsOrMapFeature = False
    End If
        

  Exit Function
ErrorHandler:
  HandleError False, "IsOrMapFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub SetAnnoSize(obj As IObject, pFeat As IFeature)
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
    
    'Capture MapNumber for each anno feature created
    lAnnoMapNumFld = FindField(obj.Class, g_pFldnames.MIMapNumberFN)
    If lAnnoMapNumFld = -1 Then Exit Sub
    'If new anno feature with no text, determine if it has a shape
    Dim pAnnoFeat As IFeature
    Dim pAOC As IObjectClass
    Set pAOC = obj.Class
    Set pAnnoFeat = obj
    Dim lFld As Long
'    Dim g As Long
'    g = 0
'    While g < pAnnoFeat.Fields.FieldCount
'        MsgBox "field = " & pAnnoFeat.Fields.Field(g).Name
'        g = g + 1
'    Wend
    lFld = pAnnoFeat.Fields.FindField("TextString")
    If lFld = -1 Then
        MsgBox "Unable to locate textstring field in anno class.  Cannot set size", vbCritical
        Exit Sub
    End If
    Dim vVal As Variant
    vVal = pAnnoFeat.Value(lFld)
    If IsNull(vVal) Then Exit Sub
        
      
'    Dim pAFeat As IAnnotationFeature
'    Dim pAFeat1 As IFeature
'    Set pAFeat1 = obj
'    If IsEmpty(pAFeat1) Then Exit Sub
'    If IsNull(pAFeat1.Shape) Then Exit Sub
'    Set pAFeat = pAFeat1
'    Dim pAElement As IElement
'    Set pAElement = pAFeat.Annotation
'    Dim pTElement As ITextElement
'    Set pTElement = pAElement
'    If IsNull(pTElement.Text) Then Exit Sub
    'If pTElement.Text = "" Then Exit Sub
    
    Set pFeat = obj
    Set pGeometry = pFeat.Shape
    If pGeometry.IsEmpty Then Exit Sub
    Set pEnv = pGeometry.Envelope
    Set pCenter = CT_GetCenterOfEnvelope(pEnv)
    Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then Exit Sub
    Set pMIFclass = pMIFlayer.FeatureClass
    sMapNum = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
    obj.Value(lAnnoMapNumFld) = sMapNum
    
    'Update the size to reflect current mapscale
    sMapScale = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
    If IsNull(sMapScale) Then Exit Sub
    
    'Determine which annotation class this is
    Set pAnnoClass = obj.Class
    Set pAnnoDset = pAnnoClass
    'If other anno, don't continue
    If LCase(pAnnoDset.Name) <> LCase(g_pFldnames.FCTLAcrAnno) And LCase(pAnnoDset.Name) <> LCase(g_pFldnames.FCTLNumAnno) Then
        Exit Sub
    End If
    
    dSize = modUtils.GetAnnoSizeByScale(pAnnoDset.Name, CLng(sMapScale))
    'Get the anno feature, its symbol, set the appropriate size
    Set pAnnotationFeature = obj
    Set pAnnotationElement = pAnnotationFeature.Annotation
    Set pElement = pAnnotationElement
    Set pTextElement = pElement
    Set pTextSym = pTextElement.Symbol
'MsgBox "Size = " & dSize
    pTextSym.Size = dSize
    pTextElement.Symbol = pTextSym
    Set pElement = pTextElement
    Set pAnnotationElement = pElement
    pAnnotationFeature.Annotation = pAnnotationElement

  Exit Sub
ErrorHandler:
  HandleError False, "SetAnnoSize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

 
Private Property Get ICommand_Enabled() As Boolean
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
     ICommand_Enabled = True
    
    Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Checked() As Boolean
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    ' ICommand_Checked =
    
    Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Name() As String
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    ICommand_Name = "AutoUpdate"
    
    Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    ' ICommand_Caption =
    
    Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    ' ICommand_Tooltip =
    
    Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    ' ICommand_Message =
    
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
    
    ' TODO: Add your implementation here
    ' ICommand_Bitmap =
    
    Exit Property
ErrorHandler:
  HandleError True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Category() As String
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
     ICommand_Category = "ORMAP"
    
    Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
    On Error GoTo ErrorHandler
    
 Set m_pApp = modUtils.GetAppRef
  If g_pApp Is Nothing Then
    Set g_pApp = m_pApp
  End If
  Set m_pMxDoc = m_pApp.Document
  Set m_pMap = m_pMxDoc.FocusMap
''
''  'Get the extension
  Dim u As New UID
  u.Value = "TaxlotEditing.ORMAPEditingExt"
  Set m_pExt = m_pApp.FindExtensionByCLSID(u)
  
  'copied
  'Get a reference to the editor extension
  Dim pID As New UID
  pID = "esriCore.Editor"
  Set m_pEditor = m_pApp.FindExtensionByCLSID(pID)
  Set m_pEditorEvents = m_pEditor
    MsgBox "command created"
    Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ICommand_OnClick()
    On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here
    
    Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
