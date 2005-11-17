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
      ItemData        =   "frmMapIndex.frx":0000
      Left            =   1440
      List            =   "frmMapIndex.frx":0002
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
'FORM USED TO CAPTURE ATTRIBUTES FOR MAPINDEX FEATURES
'THESE ATTRIBUTES USED TO CONSTRUCT ORMAPMAPNUMBER

Dim m_pMIFlayer As IFeatureLayer2
Dim m_pMIFclass As IFeatureClass
Dim m_pMIFields As IFields2
Dim m_pMIFeat As IFeature
Dim m_pTaxlotFlayer As IFeatureLayer2
Dim m_pTaxlotFClass As IFeatureClass
Dim m_lOMMapNumFld As Long
Dim m_lReliabFld As Long
Dim m_lScaleFld As Long
Dim m_lMapNumFld As Long
Dim m_lPageFld As Long
Dim m_bContinue As Boolean
Dim m_bSuccess As Boolean
Dim m_bPossiblyChanged As Boolean
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument
Const c_sModuleFileName As String = "C:\active\ModelingWorkshop_01-05-05\CustomCode\ormap\frmMapIndex.frm"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms


Private Sub cmbCounty_Click()
  On Error GoTo ErrorHandler

27:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbCounty_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbQtr_Click()
  On Error GoTo ErrorHandler

37:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbQtr_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbQtrQtr_Click()
  On Error GoTo ErrorHandler

47:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbQtrQtr_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbRange_Click()
  On Error GoTo ErrorHandler

57:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbRange_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbRangeDir_Click()
  On Error GoTo ErrorHandler

67:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbRangeDir_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbRangePart_Click()
  On Error GoTo ErrorHandler

77:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbRangePart_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbReliability_Click()
  On Error GoTo ErrorHandler

87:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbReliability_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbScale_Click()
  On Error GoTo ErrorHandler

97:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbScale_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbSection_Click()
  On Error GoTo ErrorHandler

107:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbSection_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub
Private Sub cmbSufNum_Click()
  On Error GoTo ErrorHandler

116:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbSufNum_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbSufType_Click()
  On Error GoTo ErrorHandler

126:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbSufType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbTown_Click()
  On Error GoTo ErrorHandler

136:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbTown_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub
Private Sub cmbTownDir_Click()
  On Error GoTo ErrorHandler

145:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbTownDir_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbTownPart_Click()
  On Error GoTo ErrorHandler

155:     m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbTownPart_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmdAssign_Click()
  On Error GoTo ErrorHandler
    'Assign attributes to the current MapIndex polygon
    'Verify that all fields have values entered.  Don't allow changes to
    'be applied without all values present
    Dim bValuesPresent As Boolean
168:     bValuesPresent = True
    Dim ctl As Control
170:     For Each ctl In Me.Controls
171:         If TypeOf ctl Is TextBox Then
            Dim pTxtBox As TextBox
173:             Set pTxtBox = ctl
174:             If pTxtBox.Text = "" Then
175:                 If Not pTxtBox.Name = "txtORMAPMapNum" And Not pTxtBox.Name = "txtPage" Then
176:                     bValuesPresent = False
177:                 End If
178:             End If
179:         ElseIf TypeOf ctl Is ComboBox Then
            Dim pCmb As ComboBox
181:             Set pCmb = ctl
182:             If pCmb.Text = "" Then
183:                bValuesPresent = False
184:             End If
185:         ElseIf TypeOf ctl Is ListBox Then
186:             MsgBox "listbox"
187:         End If
188:     Next ctl
    
190:     If Not bValuesPresent Then
191:        MsgBox "All fields must be filled in before assigning", vbOKOnly
       Exit Sub
193:     End If

    'Otherwise, save values
    
    'Save values if necessary
    If Not m_bPossiblyChanged Then Exit Sub
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
220:     Set pDataset = m_pMIFclass
221:     Set pWSEdit = pDataset.Workspace
222:     pWSEdit.StartEditOperation
    
    'Get a Taxlot feature, so its domains can be referenced
    Dim pFeatCur As IFeatureCursor
226:     Set pFeatCur = m_pTaxlotFlayer.Search(Nothing, True)
    Dim pTLFeat As IFeature
228:     Set pTLFeat = pFeatCur.NextFeature
    'This functionality was originally set up to work with the MapIndex feature
    'currently being edited.  The db design changed, but the structure of the code
    'has not been changed.  To obtain the domains necessary to display and save
    'values in MapIndex, taxlots are used.  This requires that at least one taxlot
    'feature exist.
234:     If pTLFeat Is Nothing Then
235:         MsgBox "No taxlot features present.  Please create at least one taxlot", vbOKOnly
        Exit Sub
237:     End If
        
    
    'MapNumber
241:     sCurMapNum = Me.txtMapNum.Text
242:     m_pMIFeat.Value(m_lMapNumFld) = sCurMapNum

    'Reliability
245:     sCurReliabil = ConvertCode(m_pMIFeat, g_pFldnames.MIReliabFN, Me.cmbReliability)
246:     m_pMIFeat.Value(m_lReliabFld) = CInt(sCurReliabil)
    
    'Scale
249:     sCurScale = ConvertCode(m_pMIFeat, g_pFldnames.MIMapScaleFN, Me.cmbScale.Text)
250:     m_pMIFeat.Value(m_lScaleFld) = CLng(sCurScale)

    'Page
253:     sCurPage = Me.txtPage.Text
254:     If IsNumeric(sCurPage) Then
        Dim lCurPage As Long
256:         lCurPage = CLng(sCurPage)
257:         m_pMIFeat.Value(m_lPageFld) = lCurPage
258:     Else
        'If null, enter a null value
        Dim nullVal As Variant
261:         m_pMIFeat.Value(m_lPageFld) = nullVal
262:     End If

    'Compile values below into the OrMAPMapNumber value
    'County
266:     sCurCounty = ConvertCode(pTLFeat, g_pFldnames.TLCountyFN, Me.cmbCounty.Text)
267:     sCurCounty = FormatOMMapNum(sCurCounty, "county")
    
    'Town
270:     sCurTown = Me.cmbTown.Text  'ConvertCode(pTLFeat, g_pFldnames.TLTownFN, Me.cmbTown.Text)
271:     sCurTown = FormatOMMapNum(sCurTown, "town")

    'TownPart
274:     sCurTownPart = Me.cmbTownPart.Text 'ConvertCode(pTLFeat, g_pFldnames.TLTownPartFN, Me.cmbTownPart.Text)
275:     sCurTownPart = FormatOMMapNum(sCurTownPart, "townpart")

    'TownDir
278:     sCurTownDir = Me.cmbTownDir.Text
279:     sCurTownDir = FormatOMMapNum(sCurTownDir, "towndir")

    'Range
282:     sCurRange = Me.cmbRange.Text
283:     sCurRange = FormatOMMapNum(sCurRange, "range")

    'RangePart
286:     sCurRangePart = Me.cmbRangePart.Text
287:     sCurRangePart = FormatOMMapNum(sCurRangePart, "rangepart")

    'RangeDir
290:     sCurRangeDir = Me.cmbRangeDir.Text
291:     sCurRangeDir = FormatOMMapNum(sCurRangeDir, "rangedir")

    'Section
294:     sCurSection = Me.cmbSection.Text
295:     sCurSection = FormatOMMapNum(sCurSection, "section")
 
    'Qtr
298:     sCurQtr = Me.cmbQtr.Text
299:     sCurQtr = FormatOMMapNum(sCurQtr, "qtr")

    'QtrQtr
302:     sCurQtrQtr = Me.cmbQtrQtr.Text
303:     sCurQtrQtr = FormatOMMapNum(sCurQtrQtr, "qtrqtr")

    'MapSuffixType
306:     sCurSuffixType = Me.cmbSufftype.Text
307:     sCurSuffixType = FormatOMMapNum(sCurSuffixType, "suffixtype")
    
    'MapSuffixNum
310:     sCurSuffixNum = Me.cmbSufftype.Text
311:     sCurSuffixNum = FormatOMMapNum(sCurSuffixNum, "suffixnum")

    'Anomaly
314:     sCurAnomaly = Me.txtAnomaly.Text
315:     sCurAnomaly = FormatOMMapNum(sCurAnomaly, "anomaly")
    
    'Concatenate everything and compare to existing ORMAPMapNumber
    'ORMAPMapNumber
319:     sCurOMMapNum = sCurCounty & sCurTown & sCurTownPart & sCurTownDir & _
        sCurRange & sCurRangePart & sCurRangeDir & sCurSection & sCurQtr & _
        sCurQtrQtr & sCurAnomaly & sCurSuffixType & sCurSuffixNum
        
323:     Me.txtORMAPMapNum.Text = sCurOMMapNum
324:     m_pMIFeat.Value(m_lOMMapNumFld) = sCurOMMapNum
    
326:     Set pTLFeat = Nothing
327:     Set pFeatCur = Nothing
328:     m_pMIFeat.Store
329:     pWSEdit.StopEditOperation

  Exit Sub
ErrorHandler:
  HandleError True, "cmdAssign_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmdHelp_Click()
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
340:     sFilePath = app.Path & "\" & "MapIndex_help.rtf"
341:     If modUtils.FileExists(sFilePath) Then
342:         If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
343:             Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
344:         End If
345:     Else
346:         MsgBox "No help available"
347:     End If
End Sub

Private Sub cmdQuit_Click()
  On Error GoTo ErrorHandler

    'Prompt for save if necessary
    
355:     Unload frmMapIndex

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
376:     lFld = pFClass.Fields.FindField(pFldName)
377:     If lFld > -1 Then
378:       LocateFields = lFld
379:     Else
380:         MsgBox "Unable to locate " & g_pFldnames.MICountyFN & " field in " & _
        g_pFldnames.FCMapIndex & " feature class"
382:         m_bContinue = False
383:     End If


  Exit Function
ErrorHandler:
  HandleError False, "FindField " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function


Private Sub txtAnomaly_Change()
393:     If Not IsNumeric(txtAnomaly.Text) Then txtAnomaly.Text = ""
End Sub

Private Sub txtSuffNum_Change()
397:     If Not IsNumeric(txtSuffNum.Text) Then txtSuffNum.Text = ""
End Sub

Public Function InitForm() As Boolean
  'Populate the drop down comboboxes with domain values
  'set defaults if a new feature
  'Select current values if an existing feature
  'Get a reference to the MXDocument
405:   Set m_pMxDoc = modUtils.GetMXDocRef

    Dim bOpenForm As Boolean
408:     Me.Refresh
    Dim response As Variant
410:     Set m_pMIFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
411:     If m_pMIFlayer Is Nothing Then
412:         MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
414:     End If
415:     Set m_pMIFclass = m_pMIFlayer.FeatureClass
    'Get the MapIndex feature layer and fclass
417:     Set m_pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
418:     If m_pTaxlotFlayer Is Nothing Then
419:         response = MsgBox("Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot & ".  " & _
        "Load " & g_pFldnames.FCMapIndex & " automatically?", vbYesNo)
422:         If response <> vbYes Then
423:             InitForm = False
            Exit Function
425:         Else
426:             modUtils.LoadFCIntoMap g_pFldnames.FCTaxlot, m_pMIFclass
427:             Set m_pTaxlotFlayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
428:         End If
429:     End If
430:     Set m_pTaxlotFClass = m_pTaxlotFlayer.FeatureClass

    'Get fields needed to populate the form
433:     Set m_pMIFields = m_pMIFclass.Fields
434:     m_bContinue = True
435:     m_lOMMapNumFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIORMAPMapNumberFN)
436:     m_lReliabFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIReliabFN)
437:     m_lScaleFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIMapScaleFN)
438:     m_lMapNumFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIMapNumberFN)
439:     m_lPageFld = modUtils.LocateFields(m_pMIFclass, g_pFldnames.MIPageFN)
440:     If Not m_bContinue Then
441:         InitForm = False
        Exit Function 'If any fields not found
443:     End If
    
    'Get the selected feature and its attributes
    Dim sExistOMMapNum As String
    Dim sExistVal As String
    Dim pFeatCur As IFeatureCursor
449:     Set pFeatCur = modUtils.GetSelectedFeatures(m_pMIFlayer)
450:     If pFeatCur Is Nothing Then
451:         InitForm = False
        Exit Function
453:     End If
454:     Set m_pMIFeat = pFeatCur.NextFeature
    
    'Get a Taxlot feature, so its domains can be referenced
    Dim pTLFeatCur As IFeatureCursor
458:     Set pTLFeatCur = m_pTaxlotFlayer.Search(Nothing, True)
    Dim pTLFeat As IFeature
460:     Set pTLFeat = pTLFeatCur.NextFeature
    
    'Populate the form with domain values
    'ORMAPMapNumber
464:     sExistOMMapNum = ReadValue(m_pMIFeat, g_pFldnames.MIORMAPMapNumberFN)
    'Verify that the number is the right length.  If not, load default values
    'into the fields below
467:     If Not Len(sExistOMMapNum) = 24 Then
468:         Me.txtORMAPMapNum.Text = ""
469:         m_bSuccess = AddCodesToCmb(g_pFldnames.MIReliabFN, m_pMIFields, Me.cmbReliability, "", True)
470:         m_bSuccess = AddCodesToCmb(g_pFldnames.MIMapScaleFN, m_pMIFields, Me.cmbScale, "", True)
471:         Me.txtPage.Text = ""
        'Convert default county to description
        Dim sDefCntyDesc As String
474:         sDefCntyDesc = modUtils.ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLCountyFN, CLng(g_pFldnames.DefCounty))
475:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLCountyFN, m_pTaxlotFClass.Fields, Me.cmbCounty, sDefCntyDesc, True)
476:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownFN, m_pTaxlotFClass.Fields, Me.cmbTown, "", True)
477:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownPartFN, m_pTaxlotFClass.Fields, Me.cmbTownPart, g_pFldnames.DefTownPart, True)
478:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownDirFN, m_pTaxlotFClass.Fields, Me.cmbTownDir, g_pFldnames.DefTownDir, True)
479:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeFN, m_pTaxlotFClass.Fields, Me.cmbRange, "", True)
480:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangePartFN, m_pTaxlotFClass.Fields, Me.cmbRangePart, g_pFldnames.DefRangePart, True)
481:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeDirFN, m_pTaxlotFClass.Fields, Me.cmbRangeDir, g_pFldnames.DefRangeDir, True)
482:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLSectNumberFN, m_pTaxlotFClass.Fields, Me.cmbSection, "", True)
483:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtr, g_pFldnames.DefQtr, True)
484:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtrQtr, g_pFldnames.DefQtrQtr, True)
485:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLSufTypeFN, m_pTaxlotFClass.Fields, Me.cmbSufftype, g_pFldnames.DefSuffType, True)
486:         Me.txtSuffNum.Text = g_pFldnames.DefSuffNum
487:         txtAnomaly.Text = ""
488:     Else
489:         Me.txtORMAPMapNum.Text = sExistOMMapNum
        'm_bSuccess = AddCodesToCmb(g_pFldnames.MICountyFN, m_pMIFields, Me.cmbCounty, sExistVal)
491:         sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIMapNumberFN)
492:         Me.txtMapNum.Text = sExistVal
        'Reliability
494:         sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIReliabFN)
495:         m_bSuccess = AddCodesToCmb(g_pFldnames.MIReliabFN, m_pMIFields, Me.cmbReliability, sExistVal, True)
        'Scale
497:         sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIMapScaleFN)
498:         m_bSuccess = AddCodesToCmb(g_pFldnames.MIMapScaleFN, m_pMIFields, Me.cmbScale, sExistVal, True)
        'Page
500:         sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIPageFN)
501:         Me.txtPage.Text = sExistVal
        'County
503:         sExistVal = ParseOMMapNum(sExistOMMapNum, "county")
504:         sExistVal = ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLCountyFN, Int(sExistVal))
505:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLCountyFN, m_pTaxlotFClass.Fields, Me.cmbCounty, sExistVal, True)
        'Town
507:         sExistVal = ParseOMMapNum(sExistOMMapNum, "town")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLTownFN)
509:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownFN, m_pTaxlotFClass.Fields, Me.cmbTown, sExistVal, True)
        'TownPart
511:         sExistVal = ParseOMMapNum(sExistOMMapNum, "townpart")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLTownPartFN)
513:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownPartFN, m_pTaxlotFClass.Fields, Me.cmbTownPart, sExistVal, True)
        'TownDir
515:         sExistVal = ParseOMMapNum(sExistOMMapNum, "towndir")
516:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownDirFN, m_pTaxlotFClass.Fields, Me.cmbTownDir, sExistVal, True)
        'Range
518:         sExistVal = ParseOMMapNum(sExistOMMapNum, "range")
519:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeFN, m_pTaxlotFClass.Fields, Me.cmbRange, sExistVal, True)
        'RangePart
521:         sExistVal = ParseOMMapNum(sExistOMMapNum, "rangepart")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLRangePartFN)
523:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangePartFN, m_pTaxlotFClass.Fields, Me.cmbRangePart, sExistVal, True)
        'RangeDir
525:         sExistVal = ParseOMMapNum(sExistOMMapNum, "rangedir")
526:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeDirFN, m_pTaxlotFClass.Fields, Me.cmbRangeDir, sExistVal, True)
        'Section
528:         sExistVal = ParseOMMapNum(sExistOMMapNum, "section")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLSectNumberFN)
530:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLSectNumberFN, m_pTaxlotFClass.Fields, Me.cmbSection, sExistVal, True)
        'Qtr
532:         sExistVal = ParseOMMapNum(sExistOMMapNum, "qtr")
533:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtr, sExistVal, True)
        'QtrQtr
535:         sExistVal = ParseOMMapNum(sExistOMMapNum, "qtrqtr")
536:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtrQtr, sExistVal, True)
        'MapSuffixType
538:         sExistVal = ParseOMMapNum(sExistOMMapNum, "suffixtype")
539:         m_bSuccess = AddCodesToCmb(g_pFldnames.TLSufTypeFN, m_pTaxlotFClass.Fields, Me.cmbSufftype, sExistVal, True)
        'MapSuffixNum
541:         sExistVal = ParseOMMapNum(sExistOMMapNum, "suffixnum")
542:         Me.txtSuffNum.Text = sExistVal
        'Anomaly
544:         sExistVal = ParseOMMapNum(sExistOMMapNum, "anomaly")
545:         txtAnomaly.Text = sExistVal
546:     End If
547:     m_bPossiblyChanged = False
548:     InitForm = True
End Function
