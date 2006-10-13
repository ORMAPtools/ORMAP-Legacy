VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdLocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CLASS MODULE FOR THE LOCATE COMMAND
'OPENS frmLocate


Implements ICommand

Private m_pApp As esriFramework.IApplication
Private m_pDoc As esriArcMapUI.IMxDocument
Private m_pMap As IMap
Private m_pExt As IExtensionConfig

 
Private Property Get ICommand_Enabled() As Boolean

15:     ICommand_Enabled = True
    
End Property
 
Private Property Get ICommand_Checked() As Boolean

21:     ICommand_Checked = False
    
End Property
 
Private Property Get ICommand_Name() As String

27:      ICommand_Name = "Locate"
    
End Property
 
Private Property Get ICommand_Caption() As String

33:      ICommand_Caption = "Locate"
    
End Property
 
Private Property Get ICommand_Tooltip() As String

39:      ICommand_Tooltip = "Locate Tool"
    
End Property
 
Private Property Get ICommand_Message() As String
    
    ' TODO: Add your implementation here
    ' ICommand_Message =
    
End Property
 
Private Property Get ICommand_HelpFile() As String
    
    ' TODO: Add your implementation here
    ' ICommand_HelpFile =
    
End Property
 
Private Property Get ICommand_HelpContextID() As Long
    
    ' TODO: Add your implementation here
    ' ICommand_HelpContextID =
    
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    
    ' TODO: Add your implementation here
    ' ICommand_Bitmap =
    
End Property
 
Private Property Get ICommand_Category() As String

73:      ICommand_Category = "ORMAP"
    
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
78:   Set m_pApp = hook
79:   If g_pApp Is Nothing Then
80:     Set g_pApp = hook
81:   End If
82:   Set m_pDoc = m_pApp.Document
83:   Set m_pMap = m_pDoc.FocusMap

End Sub
 
Private Sub ICommand_OnClick()

    'Get the MapIndex and taxlot feature layers
90:     Set m_pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
91:     If m_pTaxlotFlayer Is Nothing Then
92:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        Exit Sub
95:     End If
96:     Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
97:     If m_pMIFlayer Is Nothing Then
98:         MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
        Exit Sub
101:     End If
    'Initialize the Locate form
103:     frmLocate.Show vbModal
    
End Sub
