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
'FORM USED TO LOXATE TAXLOTS OR MAPINDEX EXTENTS BASED ON INPUT MAPNUMBER OR TAXLOT

Private m_pTaxlotFlayer As IFeatureLayer
Private m_pTaxlotFClass As IFeatureClass
Private m_pMIFlayer As IFeatureLayer
Private m_pMIFclass As IFeatureClass
Private m_pMIFields As IFields
Private m_lMapNumFld As Long
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "C:\active\ModelingWorkshop_01-05-05\CustomCode\ormap\frmLocate.frm"



Private Sub cmdApply_Click()
'Process the Locate query and zoom to MapIndex or Taxlot
  On Error GoTo ErrorHandler
  Dim pMIFlayer As IFeatureLayer
  Dim pMIFclass As IFeatureClass
  Dim pFeatureCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim sTLNum As String
  Dim sMNum As String
  Dim pQueryFilter As IQueryFilter
26:   Set pQueryFilter = New QueryFilter
27:   sTLNum = frmLocate.txtTaxlot
28:   sMNum = frmLocate.cmbMapNumber.Text

30:   If sTLNum = "" Then 'Just Query MapIndex
31:     If sMNum = "" Then 'If both empty
32:         MsgBox "Please enter a Map Number and optionally, a Taxlot number", vbOKOnly
        Exit Sub
34:     Else 'MapNumber entered with no Taxlot
35:         Set pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
36:         If pMIFlayer Is Nothing Then
37:               MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
              "This process requires a feature class called " & g_pFldnames.FCMapIndex
              Exit Sub
40:         End If
41:         Set pMIFclass = pMIFlayer.FeatureClass
42:         pQueryFilter.whereClause = "[" & g_pFldnames.MIMapNumberFN & "] = '" & frmLocate.cmbMapNumber.Text & "'"
43:         Set pFeatureCursor = pMIFclass.Search(pQueryFilter, False)
44:     End If
45:   ElseIf sTLNum <> "" And sMNum <> "" Then 'Both values entered
46:         pQueryFilter.whereClause = "[" & g_pFldnames.TLMapNumberFN & "] = '" & frmLocate.cmbMapNumber.Text & "' and [" & g_pFldnames.TLTaxlotFN & "]= '" & frmLocate.txtTaxlot & "'"
47:         Set pFeatureCursor = m_pTaxlotFClass.Search(pQueryFilter, False)
48:   Else 'Only a taxlot entered
49:         MsgBox "Please enter a MapNumber and optionally, a Taxlot numver", vbOKOnly
        Exit Sub
51:     End If
  If pFeatureCursor Is Nothing Then Exit Sub
53:   Set pFeature = pFeatureCursor.NextFeature

55:   If pFeature Is Nothing Then
56:     If sTLNum = "" Then
57:         MsgBox "Map Index could not be found.", vbInformation, "Try Again"
58:     Else
59:         MsgBox "Taxlot could not be found.", vbInformation, "Try Again"
60:     End If
61:     frmLocate.txtTaxlot = ""
62:     frmLocate.txtTaxlot.SetFocus
    Exit Sub
64:   Else
    'Zoom to selected feature
    Dim pEnvelope As IEnvelope
67:     Set pEnvelope = pFeature.Shape.Envelope
    
69:     modUtils.ZoomToExtent pEnvelope, m_pMxDoc

71:   End If
  
73:   Unload Me

  Exit Sub
ErrorHandler:
  HandleError True, "cmdApply_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

83:     Unload Me

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdHelp_Click()
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
94:     sFilePath = app.Path & "\" & "Locate_help.rtf"
95:     If modUtils.FileExists(sFilePath) Then
96:         If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
97:             Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
98:         End If
99:     Else
100:         MsgBox "No help available"
101:     End If
End Sub

Private Sub Form_Initialize()
  On Error GoTo ErrorHandler



  Exit Sub
ErrorHandler:
  HandleError True, "Form_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
    'Populate the MapIndex dropdown list so user can choose from all available MapIndex values
    Dim pApp As IApplication
117:     Set pApp = modUtils.GetAppRef ' AppRef
118:     Set m_pMxDoc = pApp.Document
    'Get the MapIndex feature layer and fclass
120:     Set m_pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
121:     If m_pTaxlotFlayer Is Nothing Then
122:         MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        Exit Sub
125:     End If
126:     Set m_pTaxlotFClass = m_pTaxlotFlayer.FeatureClass
127:     Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
128:     If m_pMIFlayer Is Nothing Then
129:         MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
        Exit Sub
132:     End If
133:     Set m_pMIFclass = m_pMIFlayer.FeatureClass
    'Get fields needed to populate the form
135:     Set m_pMIFields = m_pMIFclass.Fields
136:     m_bContinue = True
137:     m_lMapNumFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIMapNumberFN)
  Dim pQueryDef As IQueryDef
  Dim pRow As IRow
  Dim pCursor As ICursor
  Dim pFeatureWorkspace As IFeatureWorkspace
  Dim pDataset As IDataset
  
144:   Set pDataset = m_pMIFlayer 'm_pTaxlotFlayer
  
  Dim sFieldName As String
147:   sFieldName = g_pFldnames.TLMapNumberFN
  
149:   Set pFeatureWorkspace = pDataset.Workspace
150:   Set pQueryDef = pFeatureWorkspace.CreateQueryDef
151:   With pQueryDef
152:     .Tables = pDataset.Name ' Fully qualified table name
         'Problems with some values -- prevents the form from loading
154:     .SubFields = "DISTINCT(" & sFieldName & ")"
155:     Set pCursor = .Evaluate
156:   End With
  
158:   Set pRow = pCursor.NextRow
159:   Do Until pRow Is Nothing
160:     If Not IsNull(pRow.Value(0)) Then
161:         frmLocate.cmbMapNumber.AddItem pRow.Value(0) ' Note only one field in the cursor
162:     End If
163:     Set pRow = pCursor.NextRow
164:   Loop
End Sub
