VERSION 5.00
Begin VB.Form frmLocate 
   Caption         =   "Locate"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox cmbMapNumber 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtTaxlot 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Taxlot:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Map Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmLocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' File name:            frmLocate
'
' Initial Author:       Type your name here
'
' Date Created:     10/11/2006
'
' Description: FORM USED TO LOXATE TAXLOTS OR MAPINDEX EXTENTS BASED ON INPUT MAPNUMBER OR TAXLOT
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
Private m_pTaxlotFlayer As IFeatureLayer
Private m_pTaxlotFClass As IFeatureClass
Private m_pMIFlayer As IFeatureLayer
Private m_pMIFclass As IFeatureClass
Private m_pMIFields As IFields
Private m_lMapNumFld As Long
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument
'------------------------------
'Private Constants and Enums
'------------------------------
' Variables used by the Error handler function - DO NOT REMOVE
Private Const c_sModuleFileName As String = "frmLocate.frm"
'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------

'***************************************************************************
'Name:  cmdApply_Click
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   Process the Locate query and zoom to MapIndex or Taxlot

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
Private Sub cmdApply_Click()
'
  On Error GoTo ErrorHandler
  Dim pMIFlayer As IFeatureLayer
  Dim pMIFclass As IFeatureClass
  Dim pFeatureCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim sTLNum As String
  Dim sMNum As String
  Dim sMsg As String
  Dim pQueryFilter As IQueryFilter
115:   Set pQueryFilter = New QueryFilter
  
117:   sMsg = "Please enter a Map Number and optionally, a Taxlot number"
  
'++ START JWM 10/11/2006 trim and then test for length
120:   sTLNum = Trim$(frmLocate.txtTaxlot)
121:   sMNum = Trim$(frmLocate.cmbMapNumber.Text)
  
123: If Len(sTLNum) = 0 Then 'Just Query MapIndex
'  If sTLNum = "" Then
125:     If Len(sMNum) = 0 Then 'If both empty
'    If sMNum = "" Then
127:         MsgBox sMsg, vbOKOnly
128:         GoTo Process_Exit
129:     Else 'MapNumber entered with no Taxlot
130:         Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
131:         If pMIFlayer Is Nothing Then
132:               MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
              "This process requires a feature class called " & g_pFldnames.FCMapIndex
134:               GoTo Process_Exit
135:         End If
136:         Set pMIFclass = pMIFlayer.FeatureClass
137:         pQueryFilter.whereClause = "[" & g_pFldnames.MIMapNumberFN & "] = '" & frmLocate.cmbMapNumber.Text & "'"
138:         Set pFeatureCursor = pMIFclass.Search(pQueryFilter, False)
139:     End If
140:   ElseIf Len(sTLNum) > 0 And Len(sMNum) > 0 Then 'Both values entered
141:         pQueryFilter.whereClause = "[" & g_pFldnames.TLMapNumberFN & "] = '" & frmLocate.cmbMapNumber.Text & "' and [" & g_pFldnames.TLTaxlotFN & "]= '" & frmLocate.txtTaxlot & "'"
142:         Set pFeatureCursor = m_pTaxlotFClass.Search(pQueryFilter, False)
143:   Else 'Only a taxlot entered
144:         MsgBox sMsg, vbOKOnly
145:         GoTo Process_Exit
146:     End If
147:   If pFeatureCursor Is Nothing Then GoTo Process_Exit
148:   Set pFeature = pFeatureCursor.NextFeature

150:   If pFeature Is Nothing Then
151:     If Len(sTLNum) = 0 Then
152:         MsgBox "Map Index could not be found.", vbInformation, "Try Again"
153:     Else
154:         MsgBox "Taxlot could not be found.", vbInformation, "Try Again"
155:     End If
    '++ END JWM 10/11/2006
157:     frmLocate.txtTaxlot = ""
158:     frmLocate.txtTaxlot.SetFocus
159:     GoTo Process_Exit
160:   Else
    'Zoom to selected feature
    Dim pEnvelope As IEnvelope
163:     Set pEnvelope = pFeature.Shape.Envelope
    
165:     modUtils.ZoomToExtent pEnvelope, m_pMxDoc

167:   End If
  
169:   Unload Me
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "cmdApply_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

179:     Unload Me

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdHelp_Click()
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
190:     sFilePath = app.Path & "\" & "Locate_help.rtf"
191:     If modUtils.FileExists(sFilePath) Then
192:     Debug.Assert True 'need a better method to open rtf files
193:         If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
194:             Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
195:         End If
196:     Else
197:         MsgBox "No help available"
198:     End If
End Sub

Private Sub Form_Initialize()
  On Error GoTo ErrorHandler



  Exit Sub
ErrorHandler:
  HandleError True, "Form_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  Form_Load
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:

'Methods:       Populate the MapIndex dropdown list so user can choose from all available MapIndex values.
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
Private Sub Form_Load()
On Error GoTo Err_Handler
    '
    Dim pApp As IApplication
234:     Set pApp = modUtils.GetAppRef ' AppRef
235:     Set m_pMxDoc = pApp.Document
    'Get the MapIndex feature layer and fclass
237:     Set m_pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
238:     If m_pTaxlotFlayer Is Nothing Then
239:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
241:         GoTo Process_Exit
242:     End If
243:     Set m_pTaxlotFClass = m_pTaxlotFlayer.FeatureClass
244:     Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
245:     If m_pMIFlayer Is Nothing Then
246:         MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
248:         Process_Exit
249:     End If
250:     Set m_pMIFclass = m_pMIFlayer.FeatureClass
    'Get fields needed to populate the form
252:     Set m_pMIFields = m_pMIFclass.Fields
253:     m_bContinue = True
254:     m_lMapNumFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIMapNumberFN)
  Dim pQueryDef As IQueryDef
  Dim pRow As IRow
  Dim pCursor As ICursor
  Dim pFeatureWorkspace As IFeatureWorkspace
  Dim pDataset As IDataset
  
261:   Set pDataset = m_pMIFlayer 'm_pTaxlotFlayer
  
  Dim sFieldName As String
264:   sFieldName = g_pFldnames.TLMapNumberFN
  
266:   Set pFeatureWorkspace = pDataset.Workspace
267:   Set pQueryDef = pFeatureWorkspace.CreateQueryDef
268:   With pQueryDef
269:     .Tables = pDataset.Name ' Fully qualified table name
         'Problems with some values -- prevents the form from loading
271:     .SubFields = "DISTINCT(" & sFieldName & ")"
272:     Set pCursor = .Evaluate
273:   End With
  
275:   Set pRow = pCursor.NextRow
276:   Do Until pRow Is Nothing
277:     If Not IsNull(pRow.Value(0)) Then
278:         frmLocate.cmbMapNumber.AddItem pRow.Value(0) ' Note only one field in the cursor
279:     End If
280:     Set pRow = pCursor.NextRow
281:   Loop
  
Proc_Exit:
Exit Sub
Err_Handler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
