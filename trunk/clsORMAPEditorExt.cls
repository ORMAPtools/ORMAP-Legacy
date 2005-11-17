VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsORMAPEditorExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IExtension
Implements IExtensionConfig
Private m_extState As esriExtensionState
Private m_pEditor As IEditor

Private Property Get IExtension_Name() As String

    IExtension_Name = "ORMAPEditor"
    
End Property
 
Private Sub IExtension_Startup(ByRef initializationData As Variant)
    
        'copied
    'Get a reference to the editor extension
    Dim pID As New UID
    pID = "esriCore.Editor"
    Set m_pEditor = m_pApp.FindExtensionByCLSID(pID)
    Set g_pApp = initializationData

End Sub
 
Private Sub IExtension_Shutdown()
    
    Set g_pApp = Nothing
    Set m_pEditor = Nothing
    
End Sub

Private Property Get IExtensionConfig_Description() As String

    IExtensionConfig_Description = "ORMAP Editing Tools"

End Property

Private Property Get IExtensionConfig_ProductName() As String

    IExtensionConfig_ProductName = "ORMAP Taxlot Editor"

End Property

'Private Property Let IExtensionConfig_State(ByVal RHS As esriSystem.esriExtensionState)
'
'  m_extState = RHS 'ExtensionState
'  Dim m_pDoc As IDocument
'  Set m_pDoc = g_pApp.Document
'
'  Dim pbars As ICommandBars
'  Set pbars = m_pDoc.CommandBars
'  Dim pUID As UID
'  Set pUID = New UID
'  pUID.Value = "TaxlotEditing.clsTaxLotEditingToolbar"
'  Dim pbar As ICommandBar
'  Set pbar = pbars.Find(pUID)
'
'  If m_extState = esriESEnabled Then
'    pbar.Dock esriDockFloat
'  Else
'    pbar.Dock esriDockHide
'  End If
'
'End Property
'
'Private Property Get IExtensionConfig_State() As esriSystem.esriExtensionState
'
'    IExtensionConfig_State = m_extState
'
'End Property


Private Property Get IExtensionConfig_State() As esriSystem.esriExtensionState
   IExtensionConfig_State = m_extState
End Property
 
Private Property Let IExtensionConfig_State(ByVal ExtensionState As esriSystem.esriExtensionState)
   m_extState = ExtensionState
End Property

