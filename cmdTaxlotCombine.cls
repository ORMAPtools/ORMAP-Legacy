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
'
' File name:            cmdTaxlotCombine
'
' Initial Author:       Type your name here
'
' Date Created:     10/11/2006
'
' Description: CLASS MODULE FOR TAXLOTCOMBINE COMMAND
'               OPENS frmCombine
'
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
'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "cmdTaxlotCombine.cls" ' Variables used by the Error handler function - DO NOT REMOVE

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------


Implements ICommand


Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler

    'Enable if editing and at least two taxlots selected
    m_bEnabled = False
    Set m_pDoc = m_pApp.Document
    Set m_pMap = m_pDoc.FocusMap
    'This is called before the mxd is open, so need to check if any layers present
    If m_pMap.LayerCount > 0 Then
        Dim pWorkspaceEdit As IWorkspaceEdit
        Dim pFeatureLayer As IFeatureLayer
        Dim pDataset As IDataset
        Set pFeatureLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
        If pFeatureLayer Is Nothing Then
            m_bEnabled = False
            GoTo Process_Exit
        End If
        Set pDataset = pFeatureLayer.FeatureClass
        If pDataset Is Nothing Then GoTo Process_Exit
        Set pWorkspaceEdit = pDataset.Workspace
        If pWorkspaceEdit.IsBeingEdited Then 'Check if being edited
            Dim pFeatCur As IFeatureCursor
            Set pFeatCur = modUtils.GetSelectedFeatures(pFeatureLayer) 'Make sure more than one selected
            If Not pFeatCur Is Nothing Then
                Dim pFeat As IFeature
                Set pFeat = pFeatCur.NextFeature 'first feature
                If Not pFeat Is Nothing Then
                    Set pFeat = pFeatCur.NextFeature 'second feature
                    If Not pFeat Is Nothing Then
                        m_bEnabled = True
                    End If
                End If
            End If
        End If
    End If
    
    'If conditions are met, enable the command
    If m_bEnabled = True Then
        ICommand_Enabled = True
    End If
    
Process_Exit:
  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Checked() As Boolean
  On Error GoTo ErrorHandler


     ICommand_Checked = False
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler


     ICommand_Name = "Combine"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler


     ICommand_Caption = "Combine Selected Taxlots"
    

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler


     ICommand_Tooltip = "Taxlot Consolidation Tool"
    

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


  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

    
    frmCombine.Show vbModal
    

  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
