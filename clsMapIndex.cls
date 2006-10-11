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
'
' File name:            cmdMapIndex
'
' Initial Author:       Type your name here
'
' Date Created:
'
' Description: CLASS MODULE FOR MAPINDEX COMMAND
'
'
'
' Entry points:
'       List the public variables and their purposes.
'       List the properties and routines that the module exposes to the rest of the program.
'
' Dependencies: OPENS frmMapIndex
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
Private m_pApp As esriFramework.IApplication
Private m_pDoc As esriArcMapUI.IMxDocument
Private m_pMap As IMap
Private m_pExt As IExtensionConfig
Private m_bEnabled As Boolean
Private WithEvents m_pEditorEvents As Editor
Attribute m_pEditorEvents.VB_VarHelpID = -1
Private m_pEditor As IEditor2
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms
'------------------------------
'Private Constants and Enums
'------------------------------
Const c_sModuleFileName As String = "clsMapIndex.cls"

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------

'implements
Implements ICommand

 
Private Sub Class_Initialize()
  On Error GoTo ErrorHandler

    m_bEnabled = False

  Exit Sub
ErrorHandler:
  HandleError True, "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub Class_Terminate()
  On Error GoTo ErrorHandler

    Set m_pEditorEvents = Nothing
    Set m_pEditor = Nothing
    Set g_pApp = Nothing
    Set g_pFldnames = Nothing
  Exit Sub
ErrorHandler:
  HandleError True, "Class_Terminate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler


     ICommand_Enabled = m_bEnabled
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Checked() As Boolean
  On Error GoTo ErrorHandler

    
     ICommand_Checked = False
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler

    
     ICommand_Name = "MapIndex"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler

    
     ICommand_Caption = "Map Index"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler

    
     ICommand_Tooltip = "Map Index Tool"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_Message() As String
  On Error GoTo ErrorHandler

    
     ICommand_Message = ""
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Property
 
Private Property Get ICommand_HelpFile() As String
  On Error GoTo ErrorHandler

    
     ICommand_HelpFile = ""
    

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

    
     ICommand_Category = "ORMAP"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
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
  Set m_pEditor = m_pApp.FindExtensionByCLSID(pID)
  Set m_pEditorEvents = m_pEditor
    

  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub
 
Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

    If frmMapIndex.InitForm Then
        frmMapIndex.Show vbModal
    End If
    

  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub m_pEditorEvents_OnSelectionChanged()
  On Error GoTo ErrorHandler
    'Determine if MapIndex command should be enabled
    Dim pMIFlayer As IFeatureLayer2
    Dim pMIFclass As IFeatureClass
    Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then
        m_bEnabled = False
        Exit Sub
    End If
    If modUtils.HasSelectedFeatures(pMIFlayer) Then
        m_bEnabled = True
    Else
        m_bEnabled = False
    End If

  Exit Sub
ErrorHandler:
  HandleError False, "m_pEditorEvents_OnSelectionChanged " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

