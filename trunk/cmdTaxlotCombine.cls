VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdTaxlotCombine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CLASS MODULE FOR TAXLOTCOMBINE COMMAND
'OPENS frmCombine
Implements ICommand

Private m_pApp As esriFramework.IApplication
Private m_pDoc As esriArcMapUI.IMxDocument
Private m_pMap As IMap
Private m_pExt As IExtensionConfig
Private m_bEnabled As Boolean
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "C:\active\ModelingWorkshop_01-05-05\CustomCode\ormap\cmdTaxlotCombine.cls"

 
Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler

    'Enable if editing and at least two taxlots selected
18:     m_bEnabled = False
19:     Set m_pDoc = m_pApp.Document
20:     Set m_pMap = m_pDoc.FocusMap
    'This is called before the mxd is open, so need to check if any layers present
22:     If m_pMap.LayerCount > 0 Then
        Dim pWorkspaceEdit As IWorkspaceEdit
        Dim pFeatureLayer As IFeatureLayer
        Dim pDataset As IDataset
26:         Set pFeatureLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
27:         If pFeatureLayer Is Nothing Then
28:             m_bEnabled = False
            Exit Property
30:         End If
31:         Set pDataset = pFeatureLayer.FeatureClass
        If pDataset Is Nothing Then Exit Property
33:         Set pWorkspaceEdit = pDataset.Workspace
34:         If pWorkspaceEdit.IsBeingEdited Then 'Check if being edited
            Dim pFeatCur As IFeatureCursor
36:             Set pFeatCur = modUtils.GetSelectedFeatures(pFeatureLayer) 'Make sure more than one selected
37:             If Not pFeatCur Is Nothing Then
                Dim pFeat As IFeature
39:                 Set pFeat = pFeatCur.NextFeature 'first feature
40:                 If Not pFeat Is Nothing Then
41:                     Set pFeat = pFeatCur.NextFeature 'second feature
42:                     If Not pFeat Is Nothing Then
43:                         m_bEnabled = True
44:                     End If
45:                 End If
46:             End If
47:         End If
48:     End If
    
    'If conditions are met, enable the command
51:     If m_bEnabled = True Then
52:         ICommand_Enabled = True
53:     End If
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Checked() As Boolean
  On Error GoTo ErrorHandler


65:      ICommand_Checked = False
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler


77:      ICommand_Name = "Combine"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler


89:      ICommand_Caption = "Combine Selected Taxlots"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler


101:      ICommand_Tooltip = "Taxlot Consolidation Tool"
    

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


165:      ICommand_Category = "ORMAP"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
  On Error GoTo ErrorHandler

    
177:   Set m_pApp = hook
178:   If g_pApp Is Nothing Then
179:     Set g_pApp = hook
180:   End If
181:   Set m_pDoc = m_pApp.Document
182:   Set m_pMap = m_pDoc.FocusMap


  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

    
194:     frmCombine.Show vbModal
    

  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
