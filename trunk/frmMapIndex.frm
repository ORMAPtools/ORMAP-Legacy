VERSION 5.00
Begin VB.Form frmMapIndex 
   Caption         =   "Map Index"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbSufftype 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtSuffNum 
      Height          =   285
      Left            =   4440
      MaxLength       =   3
      TabIndex        =   37
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtAnomaly 
      Height          =   285
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   35
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtPage 
      Height          =   285
      Left            =   8640
      TabIndex        =   32
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox txtMapNum 
      Height          =   285
      Left            =   7560
      MaxLength       =   20
      TabIndex        =   31
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtORMAPMapNum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   26
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   25
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAssign 
      Caption         =   "Assign"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   23
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox cmbQtrQtr 
      Height          =   315
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox cmbRangeDir 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cmbTownDir 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox cmbScale 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox cmbQtr 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox cmbRangePart 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cmbTownPart 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox cmbReliability 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cmbSection 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox cmbRange 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cmbTown 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox cmbCounty 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblAnomaly 
      Caption         =   "Anomaly:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   36
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblPage 
      Caption         =   "Page:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   34
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblMapNum 
      Caption         =   "MapNum:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   33
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblQtrQtr 
      Caption         =   "QtrQtr:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   30
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblRangeDir 
      Caption         =   "RangeDir:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   29
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblTownDir 
      Caption         =   "TownDir:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   28
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblOMMapNumber 
      Caption         =   "ORMAPMapNumber:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   27
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblSufNum 
      Caption         =   "SufNum:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   22
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblScale 
      Caption         =   "Scale:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblQtr 
      Caption         =   "Qtr:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblRangePart 
      Caption         =   "RangePart:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblTownPart 
      Caption         =   "TownPart:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblReliability 
      Caption         =   "Reliability:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblSufType 
      Caption         =   "SufType:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblSection 
      Caption         =   "Section:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblRange 
      Caption         =   "Range:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblTown 
      Caption         =   "Town:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblCounty 
      Caption         =   "County:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMapIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' File name:            frmMapIndex
'
' Initial Author:
'
' Date Created:
'
' Description: FORM USED TO CAPTURE ATTRIBUTES FOR MAPINDEX FEATURES
'THESE ATTRIBUTES USED TO CONSTRUCT ORMAPMAPNUMBER
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
'JWM 10/11/2006 Added comment headers


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
Dim m_pMIFlayer As IFeatureLayer2
Dim m_pMIFclass As IFeatureClass
Dim m_pMIFields As IFields2
Dim m_pMIFeat As IFeature
Dim m_pTaxlotFlayer As IFeatureLayer2
Dim m_pTaxlotFClass As IFeatureClass

Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument

Dim m_lOMMapNumFld As Long
Dim m_lReliabFld As Long
Dim m_lScaleFld As Long
Dim m_lMapNumFld As Long
Dim m_lPageFld As Long
Dim m_bContinue As Boolean
Dim m_bSuccess As Boolean
Dim m_bPossiblyChanged As Boolean
Private m_ParentHWND As Long ' Set this to get correct parenting of Error handler forms
'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "frmMapIndex.frm"

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------


Private Sub cmbCounty_Click()
  On Error GoTo ErrorHandler

101:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbCounty_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbQtr_Click()
  On Error GoTo ErrorHandler

111:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbQtr_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbQtrQtr_Click()
  On Error GoTo ErrorHandler

121:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbQtrQtr_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbRange_Click()
  On Error GoTo ErrorHandler

131:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbRange_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbRangeDir_Click()
  On Error GoTo ErrorHandler

141:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbRangeDir_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbRangePart_Click()
  On Error GoTo ErrorHandler

151:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbRangePart_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbReliability_Click()
  On Error GoTo ErrorHandler

161:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbReliability_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbScale_Click()
  On Error GoTo ErrorHandler

171:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbScale_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbSection_Click()
  On Error GoTo ErrorHandler

181:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbSection_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub
Private Sub cmbSufNum_Click()
  On Error GoTo ErrorHandler

190:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbSufNum_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbSufType_Click()
  On Error GoTo ErrorHandler

200:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbSufType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbTown_Click()
  On Error GoTo ErrorHandler

210:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbTown_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub
Private Sub cmbTownDir_Click()
  On Error GoTo ErrorHandler

219:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbTownDir_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbTownPart_Click()
  On Error GoTo ErrorHandler

229:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbTownPart_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

'***************************************************************************
'Name:  cmdAssign_Click
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Assign attributes to the current MapIndex polygon

'Description:   Verify that all fields have values entered.  Don't allow changes to
'               be applied without all values present
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
Private Sub cmdAssign_Click()
  On Error GoTo ErrorHandler
    
    Dim bValuesPresent As Boolean
262:     bValuesPresent = True
    Dim ctl As Control
264:     For Each ctl In Me.Controls
265:         If TypeOf ctl Is TextBox Then
            Dim pTxtBox As TextBox
267:             Set pTxtBox = ctl
268:             If Len(pTxtBox.Text) = 0 Then
269:                 If StrComp(pTxtBox.Name, "txtORMAPMapNum", vbTextCompare) <> 0 And StrComp(pTxtBox.Name, "txtPage", vbTextCompare) <> 0 Then
'                If Not pTxtBox.Name = "txtORMAPMapNum" And Not pTxtBox.Name = "txtPage" Then
271:                     bValuesPresent = False
272:                 End If
273:             End If
274:         ElseIf TypeOf ctl Is ComboBox Then
            Dim pCmb As ComboBox
276:             Set pCmb = ctl
277:             If Len(pCmb.Text) = 0 Then
278:                bValuesPresent = False
279:             End If
280:         ElseIf TypeOf ctl Is ListBox Then
'++  JWM 10/11/2006 why is this elseif here
282:             Debug.Assert True
283:             MsgBox "listbox"
284:         End If
285:     Next ctl
    
287:     If Not bValuesPresent Then
288:        MsgBox "All fields must be filled in before assigning", vbOKOnly
289:        GoTo Process_Exit
290:     End If

    'Otherwise, save values
    
    'Save values if necessary
295:     If Not m_bPossiblyChanged Then GoTo Process_Exit
    Dim sExistOMMapNum As String
    Dim sCurOMMapNum As String
    Dim sCurMapNum As String
    Dim sCurReliabil As String
    Dim sCurScale As String
    Dim sCurPage As String
    Dim sCurCounty As String
    Dim sCurTown As String
    Dim sCurTownPart As String
    Dim sCurTownDir As String
    Dim sCurRange As String
    Dim sCurRangePart As String
    Dim sCurRangeDir As String
    Dim sCurSection As String
    Dim sCurQtr As String
    Dim sCurQtrQtr As String
    Dim sCurSuffixType As String
    Dim sCurSuffixNum As String
    Dim sCurAnomaly As String
    Dim pWSEdit As IWorkspaceEdit
    Dim pDataset As IDataset
317:     Set pDataset = m_pMIFclass
318:     Set pWSEdit = pDataset.Workspace
319:     pWSEdit.StartEditOperation
    
    'Get a Taxlot feature, so its domains can be referenced
    Dim pFeatCur As IFeatureCursor
323:     Set pFeatCur = m_pTaxlotFlayer.Search(Nothing, True)
    Dim pTLFeat As IFeature
325:     Set pTLFeat = pFeatCur.NextFeature
    'This functionality was originally set up to work with the MapIndex feature
    'currently being edited.  The db design changed, but the structure of the code
    'has not been changed.  To obtain the domains necessary to display and save
    'values in MapIndex, taxlots are used.  This requires that at least one taxlot
    'feature exist.
331:     If pTLFeat Is Nothing Then
332:         MsgBox "No taxlot features present.  Please create at least one taxlot", vbOKOnly
333:         GoTo Process_Exit
334:     End If
    
    'MapNumber
337:     sCurMapNum = Me.txtMapNum.Text
338:     m_pMIFeat.Value(m_lMapNumFld) = sCurMapNum

    'Reliability
341:     sCurReliabil = ConvertCode(m_pMIFeat, g_pFldnames.MIReliabFN, Me.cmbReliability)
342:     m_pMIFeat.Value(m_lReliabFld) = CInt(sCurReliabil)
    
    'Scale
345:     sCurScale = ConvertCode(m_pMIFeat, g_pFldnames.MIMapScaleFN, Me.cmbScale.Text)
346:     m_pMIFeat.Value(m_lScaleFld) = CLng(sCurScale)

    'Page
349:     sCurPage = Me.txtPage.Text
350:     If IsNumeric(sCurPage) Then
        Dim lCurPage As Long
352:         lCurPage = CLng(sCurPage)
353:         m_pMIFeat.Value(m_lPageFld) = lCurPage
354:     Else
        'If null, enter a null value
        Dim nullVal As Variant
357:         m_pMIFeat.Value(m_lPageFld) = nullVal
358:     End If

    'Compile values below into the OrMAPMapNumber value
    'County
362:     sCurCounty = ConvertCode(pTLFeat, g_pFldnames.TLCountyFN, Me.cmbCounty.Text)
363:     sCurCounty = FormatOMMapNum(sCurCounty, "county")
    
    'Town
366:     sCurTown = Me.cmbTown.Text  'ConvertCode(pTLFeat, g_pFldnames.TLTownFN, Me.cmbTown.Text)
367:     sCurTown = FormatOMMapNum(sCurTown, "town")

    'TownPart
370:     sCurTownPart = Me.cmbTownPart.Text 'ConvertCode(pTLFeat, g_pFldnames.TLTownPartFN, Me.cmbTownPart.Text)
371:     sCurTownPart = FormatOMMapNum(sCurTownPart, "townpart")
    'If Len(sCurTownPart) = 3 Then
        'sCurTownPart = Replace(sCurTownPart, ".", "")
        'sCurTownPart = Left(sCurTownPart, 1) & "." & Right(sCurTownPart, 2)
    'End If
    'TownDir
377:     sCurTownDir = Me.cmbTownDir.Text
378:     sCurTownDir = FormatOMMapNum(sCurTownDir, "towndir")

    'Range
381:     sCurRange = Me.cmbRange.Text
382:     sCurRange = FormatOMMapNum(sCurRange, "range")

    'RangePart
385:     sCurRangePart = Me.cmbRangePart.Text
386:     sCurRangePart = FormatOMMapNum(sCurRangePart, "rangepart")
    'If Len(sCurRangePart) = 3 Then
    'sCurRangePart = Replace(sCurRangePart, ".", "")
        'sCurRangePart = Left(sCurRangePart, 1) & "." & Right(sCurRangePart, 2)
    'End If
    'RangeDir
392:     sCurRangeDir = Me.cmbRangeDir.Text
393:     sCurRangeDir = FormatOMMapNum(sCurRangeDir, "rangedir")

    'Section
396:     sCurSection = Me.cmbSection.Text
397:     sCurSection = FormatOMMapNum(sCurSection, "section")
 
    'Qtr
400:     sCurQtr = Me.cmbQtr.Text
401:     sCurQtr = FormatOMMapNum(sCurQtr, "qtr")

    'QtrQtr
404:     sCurQtrQtr = Me.cmbQtrQtr.Text
405:     sCurQtrQtr = FormatOMMapNum(sCurQtrQtr, "qtrqtr")

    'MapSuffixType
408:     sCurSuffixType = ConvertCode(pTLFeat, g_pFldnames.TLSufTypeFN, cmbSufftype.Text)
409:     sCurSuffixType = FormatOMMapNum(sCurSuffixType, "suffixtype")
    
    'MapSuffixNum
412:     sCurSuffixNum = Me.txtSuffNum.Text
413:     sCurSuffixNum = FormatOMMapNum(sCurSuffixNum, "suffixnum")

    'Anomaly
416:     sCurAnomaly = Me.txtAnomaly.Text
417:     sCurAnomaly = FormatOMMapNum(sCurAnomaly, "anomaly")
    
    'Concatenate everything and compare to existing ORMAPMapNumber
    'ORMAPMapNumber
421:     sCurOMMapNum = sCurCounty & sCurTown & sCurTownPart & sCurTownDir & _
        sCurRange & sCurRangePart & sCurRangeDir & sCurSection & sCurQtr & _
        sCurQtrQtr & sCurAnomaly & sCurSuffixType & sCurSuffixNum
        
425:     Me.txtORMAPMapNum.Text = sCurOMMapNum
426:     m_pMIFeat.Value(m_lOMMapNumFld) = sCurOMMapNum
    
428:     Set pTLFeat = Nothing
429:     Set pFeatCur = Nothing
430:     m_pMIFeat.Store
431:     pWSEdit.StopEditOperation
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "cmdAssign_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmdHelp_Click()
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
442:     sFilePath = app.Path & "\" & "MapIndex_help.rtf"
443:     If modUtils.FileExists(sFilePath) Then
444:     Debug.Assert True 'Need a better way to open files
445:         If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
446:             Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
447:         End If
448:     Else
449:         MsgBox "No help available"
450:     End If
End Sub

Private Sub cmdQuit_Click()
  On Error GoTo ErrorHandler

    'Prompt for save if necessary
    
458:     Unload frmMapIndex

  Exit Sub
ErrorHandler:
  HandleError True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
  

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Function LocateFields(pFldName As String, pFClass As IFeatureClass) As Long
  On Error GoTo ErrorHandler

    'Generic function to locate a field and warn user if it can't be found
    Dim lFld As Long
479:     lFld = pFClass.Fields.FindField(pFldName)
480:     If lFld > -1 Then
481:       LocateFields = lFld
482:     Else
483:         MsgBox "Unable to locate " & g_pFldnames.MICountyFN & " field in " & _
        g_pFldnames.FCMapIndex & " feature class"
485:         m_bContinue = False
486:     End If


  Exit Function
ErrorHandler:
  HandleError False, "FindField " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function


Private Sub txtAnomaly_Change()
    'This is now able to contain non numeric values
    'If Not IsNumeric(txtAnomaly.Text) Then txtAnomaly.Text = ""
  On Error GoTo ErrorHandler

500:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "txtAnomaly_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub

Private Sub txtMapNum_Change()
509:     m_bPossiblyChanged = True
End Sub

Private Sub txtSuffNum_Change()
513:     If Not IsNumeric(txtSuffNum.Text) Then txtSuffNum.Text = ""
    
515:     m_bPossiblyChanged = True
End Sub

'***************************************************************************
'Name:  InitForm
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Called From:   cmbMapIndex.ICommand_OnClick
'Description:   Populate the drop down comboboxes with domain values. Set defaults if a new feature.
'               Select current values if an existing feature.
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
Public Function InitForm() As Boolean

  'Get a reference to the MXDocument
542:   Set m_pMxDoc = modUtils.GetMXDocRef

    Dim bOpenForm As Boolean
545:     Me.Refresh
    Dim response As Variant
547:     Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
548:     If m_pMIFlayer Is Nothing Then
549:         MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
551:     End If
552:     Set m_pMIFclass = m_pMIFlayer.FeatureClass
    'Get the MapIndex feature layer and fclass
554:     Set m_pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
555:     If m_pTaxlotFlayer Is Nothing Then
556:         response = MsgBox("Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot & ".  " & _
        "Load " & g_pFldnames.FCMapIndex & " automatically?", vbYesNo)
559:         If response <> vbYes Then
560:             InitForm = False
561:             GoTo Process_Exit
562:         Else
563:             modUtils.LoadFCIntoMap g_pFldnames.FCTaxlot, m_pMIFclass
564:             Set m_pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
565:         End If
566:     End If
567:     Set m_pTaxlotFClass = m_pTaxlotFlayer.FeatureClass

    'Get fields needed to populate the form
570:     Set m_pMIFields = m_pMIFclass.Fields
571:     m_bContinue = True
572:     m_lOMMapNumFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIORMAPMapNumberFN)
573:     m_lReliabFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIReliabFN)
574:     m_lScaleFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIMapScaleFN)
575:     m_lMapNumFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIMapNumberFN)
576:     m_lPageFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIPageFN)
577:     If Not m_bContinue Then
578:         InitForm = False
579:         GoTo Process_Exit 'If any fields not found
580:     End If
    
    'Get the selected feature and its attributes
    Dim sExistOMMapNum As String
    Dim sExistVal As String
    Dim pFeatCur As IFeatureCursor
586:     Set pFeatCur = modUtils.GetSelectedFeatures(m_pMIFlayer)
587:     If pFeatCur Is Nothing Then
588:         InitForm = False
589:         GoTo Process_Exit
590:     End If
591:     Set m_pMIFeat = pFeatCur.NextFeature
    
    'Get a Taxlot feature, so its domains can be referenced
    Dim pTLFeatCur As IFeatureCursor
595:     Set pTLFeatCur = m_pTaxlotFlayer.Search(Nothing, True)
    Dim pTLFeat As IFeature
597:     Set pTLFeat = pTLFeatCur.NextFeature
    
    'Populate the form with domain values
    'ORMAPMapNumber
601:     sExistOMMapNum = ReadValue(m_pMIFeat, g_pFldnames.MIORMAPMapNumberFN)
    'Verify that the number is the right length.  If not, load default values
    'into the fields below
604:     If Not Len(sExistOMMapNum) = 24 Then
605:         Me.txtORMAPMapNum.Text = ""
606:         m_bSuccess = AddCodesToCmb(g_pFldnames.MIReliabFN, m_pMIFields, Me.cmbReliability, "", True)
607:         m_bSuccess = AddCodesToCmb(g_pFldnames.MIMapScaleFN, m_pMIFields, Me.cmbScale, "", True)
608:         Me.txtPage.Text = ""
        'Convert default county to description
        Dim sDefCntyDesc As String
611:         sDefCntyDesc = modUtils.ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLCountyFN, CLng(g_pFldnames.DefCounty))
        
613:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLCountyFN, m_pTaxlotFClass.Fields, Me.cmbCounty, sDefCntyDesc, True)
614:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownFN, m_pTaxlotFClass.Fields, Me.cmbTown, "", True)
615:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownPartFN, m_pTaxlotFClass.Fields, Me.cmbTownPart, modUtils.ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLTownPartFN, CDbl(g_pFldnames.DefTownPart)), True)
616:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownDirFN, m_pTaxlotFClass.Fields, Me.cmbTownDir, g_pFldnames.DefTownDir, True)
617:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeFN, m_pTaxlotFClass.Fields, Me.cmbRange, "", True)
618:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangePartFN, m_pTaxlotFClass.Fields, Me.cmbRangePart, modUtils.ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLRangePartFN, CDbl(g_pFldnames.DefRangePart)), True)
619:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeDirFN, m_pTaxlotFClass.Fields, Me.cmbRangeDir, g_pFldnames.DefRangeDir, True)
620:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLSectNumberFN, m_pTaxlotFClass.Fields, Me.cmbSection, "", True)
621:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtr, g_pFldnames.DefQtr, True)
622:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtrQtr, g_pFldnames.DefQtrQtr, True)
623:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLSufTypeFN, m_pTaxlotFClass.Fields, Me.cmbSufftype, modUtils.ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.MIMapSuffNumFN, g_pFldnames.DefSuffType))
        
625:         Me.txtSuffNum.Text = g_pFldnames.DefSuffNum
626:         Me.txtAnomaly.Text = g_pFldnames.DefAnomaly
        'txtAnomaly.Text = ""
628:     Else
629:         Me.txtORMAPMapNum.Text = sExistOMMapNum
        'm_bSuccess = AddCodesToCmb(g_pFldnames.MICountyFN, m_pMIFields, Me.cmbCounty, sExistVal)
631:         sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIMapNumberFN)
632:         Me.txtMapNum.Text = sExistVal
        'Reliability
634:         sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIReliabFN)
635:         m_bSuccess = AddCodesToCmb(g_pFldnames.MIReliabFN, m_pMIFields, Me.cmbReliability, sExistVal, True)
        'Scale
637:         sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIMapScaleFN)
638:         m_bSuccess = AddCodesToCmb(g_pFldnames.MIMapScaleFN, m_pMIFields, Me.cmbScale, sExistVal, True)
        'Page
640:         sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIPageFN)
641:         Me.txtPage.Text = sExistVal
        'County
643:         sExistVal = ParseOMMapNum(sExistOMMapNum, "county")
644:         sExistVal = ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLCountyFN, Int(sExistVal))
645:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLCountyFN, m_pTaxlotFClass.Fields, Me.cmbCounty, sExistVal, True)
        'Town
647:         sExistVal = ParseOMMapNum(sExistOMMapNum, "town")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLTownFN)
649:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownFN, m_pTaxlotFClass.Fields, Me.cmbTown, sExistVal, True)
        'TownPart
651:         sExistVal = ParseOMMapNum(sExistOMMapNum, "townpart")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLTownPartFN)
        'If Len(sExistVal) = 3 Then sExistVal = Left(sExistVal, 1) & "." & Right(sExistVal, 2)
654:         If Len(sExistVal) = 3 And Left(sExistVal, 1) = "." Then sExistVal = "0" & sExistVal
655:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownPartFN, m_pTaxlotFClass.Fields, Me.cmbTownPart, sExistVal, True)
        'TownDir
657:         sExistVal = ParseOMMapNum(sExistOMMapNum, "towndir")
658:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownDirFN, m_pTaxlotFClass.Fields, Me.cmbTownDir, sExistVal, True)
        'Range
660:         sExistVal = ParseOMMapNum(sExistOMMapNum, "range")
661:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeFN, m_pTaxlotFClass.Fields, Me.cmbRange, sExistVal, True)
        'RangePart
663:         sExistVal = ParseOMMapNum(sExistOMMapNum, "rangepart")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLRangePartFN)
        'If Len(sExistVal) = 3 Then sExistVal = Left(sExistVal, 1) & "." & Right(sExistVal, 2)
666:         If Len(sExistVal) = 3 And Left(sExistVal, 1) = "." Then sExistVal = "0" & sExistVal
667:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangePartFN, m_pTaxlotFClass.Fields, Me.cmbRangePart, sExistVal, True)
        'RangeDir
669:         sExistVal = ParseOMMapNum(sExistOMMapNum, "rangedir")
670:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeDirFN, m_pTaxlotFClass.Fields, Me.cmbRangeDir, sExistVal, True)
        'Section
672:         sExistVal = ParseOMMapNum(sExistOMMapNum, "section")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLSectNumberFN)
674:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLSectNumberFN, m_pTaxlotFClass.Fields, Me.cmbSection, sExistVal, True)
        'Qtr
676:         sExistVal = ParseOMMapNum(sExistOMMapNum, "qtr")
677:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtr, sExistVal, True)
        'QtrQtr
679:         sExistVal = ParseOMMapNum(sExistOMMapNum, "qtrqtr")
680:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtrQtr, sExistVal, True)
        'MapSuffixType
682:         sExistVal = modUtils.ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLSufTypeFN, ParseOMMapNum(sExistOMMapNum, "suffixtype"))
683:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLSufTypeFN, m_pTaxlotFClass.Fields, Me.cmbSufftype, sExistVal, True)
        'MapSuffixNum
685:         sExistVal = ParseOMMapNum(sExistOMMapNum, "suffixnum")
686:         Me.txtSuffNum.Text = sExistVal
        'Anomaly
688:         sExistVal = ParseOMMapNum(sExistOMMapNum, "anomaly")
689:         txtAnomaly.Text = sExistVal
690:     End If
691:     m_bPossiblyChanged = False
692:     InitForm = True
Process_Exit:
Exit Function

End Function
