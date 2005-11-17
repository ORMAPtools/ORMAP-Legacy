VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdMapIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CLASS MODULE FOR MAPINDEX COMMAND
'OPENS frmMapIndex
Implements ICommand

Private m_pApp As esriFramework.IApplication
Private m_pDoc As esriArcMapUI.IMxDocument
Private m_pMap As IMap
Private m_pExt As IExtensionConfig
Private m_bEnabled As Boolean
Private WithEvents m_pEditorEvents As Editor
Attribute m_pEditorEvents.VB_VarHelpID = -1
Private m_pEditor As IEditor2
Const c_sModuleFileName As String = "C:\active\ModelingWorkshop_01-05-05\CustomCode\ormap\clsMapIndex.cls"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms

 
Private Sub Class_Initialize()
  On Error GoTo ErrorHandler

19:     m_bEnabled = False

  Exit Sub
ErrorHandler:
  HandleError True, "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub Class_Terminate()
  On Error GoTo ErrorHandler

29:     Set m_pEditorEvents = Nothing
30:     Set m_pEditor = Nothing

  Exit Sub
ErrorHandler:
  HandleError True, "Class_Terminate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler


41:      ICommand_Enabled = m_bEnabled
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Checked() As Boolean
  On Error GoTo ErrorHandler

    
53:      ICommand_Checked = False
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler

    
65:      ICommand_Name = "MapIndex"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler

    
77:      ICommand_Caption = "Map Index"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler

    
89:      ICommand_Tooltip = "Map Index Tool"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Message() As String
  On Error GoTo ErrorHandler

    
101:      ICommand_Message = ""
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_HelpFile() As String
  On Error GoTo ErrorHandler

    
113:      ICommand_HelpFile = ""
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_HelpContextID() As Long
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here
    ' ICommand_HelpContextID =
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpContextID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here
    ' ICommand_Bitmap =
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Category() As String
  On Error GoTo ErrorHandler

    
151:      ICommand_Category = "ORMAP"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
  On Error GoTo ErrorHandler

    
163:   Set m_pApp = hook
164:   If g_pApp Is Nothing Then
165:     Set g_pApp = hook
166:   End If
167:   Set m_pDoc = m_pApp.Document
168:   Set m_pMap = m_pDoc.FocusMap

  'Get a reference to the editor extension
  Dim pID As New UID
172:   pID = "esriCore.Editor"
173:   Set m_pEditor = m_pApp.FindExtensionByCLSID(pID)
174:   Set m_pEditorEvents = m_pEditor
    

  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub
 
Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

185:     If frmMapIndex.InitForm Then
186:         frmMapIndex.Show vbModal
187:     End If
    

  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub m_pEditorEvents_OnSelectionChanged()
  On Error GoTo ErrorHandler
    'Determine if MapIndex command should be enabled
    Dim pMIFlayer As IFeatureLayer2
    Dim pMIFclass As IFeatureClass
200:     Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
201:     If pMIFlayer Is Nothing Then
202:         m_bEnabled = False
        Exit Sub
204:     End If
205:     If modUtils.HasSelectedFeatures(pMIFlayer) Then
206:         m_bEnabled = True
207:     Else
208:         m_bEnabled = False
209:     End If

  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnSelectionChanged " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

