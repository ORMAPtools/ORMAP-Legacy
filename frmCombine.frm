VERSION 5.00
Begin VB.Form frmCombine 
   Caption         =   "Taxlot Combine"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNewTaxlot 
      Height          =   375
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   4
      Top             =   360
      Width           =   2175
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
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
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
      TabIndex        =   1
      Top             =   960
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
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "New Taxlot:"
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
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmCombine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FORM USED TO COMBINE SELECTED TAXLOTS

Private m_pEditor As IEditor
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "C:\active\ModelingWorkshop_01-05-05\CustomCode\ormap\frmCombine.frm"



Private Sub cmdApply_Click()
  On Error GoTo ErrorHandler
    'Code that combines taxlots
    Dim pMXDoc As IMxDocument
    Dim pMap As IMap
16:     Set pMXDoc = g_pApp.Document
17:     Set pMap = pMXDoc.FocusMap
    'Validate new taxlot number entered and make sure it doesn't exist
19:     If Not IsNumeric(Me.txtNewTaxlot.Text) Or Not (Len(Me.txtNewTaxlot.Text) = 5) Then
20:       MsgBox "Invalid Start Value.  Please enter a 5-digit number", vbOKOnly, "Error"
21:       Me.txtNewTaxlot.SetFocus
      Exit Sub
23:     End If

    'Taxlots already selected and taxlot number known
    Dim pFeatcls As IFeatureClass
    Dim pWorkspaceEdit As IWorkspaceEdit
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
30:     Set pFeatureLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
31:     Set pFeatcls = pFeatureLayer.FeatureClass
32:     Set pDataset = pFeatureLayer.FeatureClass
    If pDataset Is Nothing Then Exit Sub
34:     Set pWorkspaceEdit = pDataset.Workspace
35:     If pWorkspaceEdit.IsBeingEdited Then 'Check if being edited
        Dim pFeatCur As IFeatureCursor
37:         Set pFeatCur = modUtils.GetSelectedFeatures(pFeatureLayer) 'Make sure more than one selected
38:         If Not pFeatCur Is Nothing Then
            'Combine taxlots
            ' code to merge the features, evaluate the merge rules and assign values to fields appropriatly
            
            ' start edit operation
43:             m_pEditor.StartOperation
            
            ' create a new feature to be the merge feature
            Dim pCurFeature As IFeature
            Dim pNewFeature As IFeature
            Dim lCount As Long
49:             Set pNewFeature = pFeatcls.CreateFeature
            
            ' create the new geometry.
            Dim pGeom As IGeometry
            Dim pTmpGeom As IGeometry
            Dim pOutputGeometry As IGeometry
            Dim pTopoOperator As ITopologicalOperator
              
            ' initialize the default values for the new feature
            Dim pOutRSType As IRowSubtypes
59:             Set pOutRSType = pNewFeature
60:             If lSCode <> 0 Then
61:               pOutRSType.SubtypeCode = lSCode
62:             End If
63:             pOutRSType.InitDefaultValues
            
            ' get the first feature
66:             Set pCurFeature = pFeatCur.NextFeature
            Dim pFlds As IFields
68:             Set pFlds = pFeatcls.Fields
            
            Dim pArea As IArea
71:             Set pArea = pCurFeature.Shape
            'Now that we have a feature,
            'Verify that within this map index, this taxlot number is unique
            'If not unique, prompt user to enter a new value
75:             If Not modUtils.ValidateTaxlotNum(frmCombine.txtNewTaxlot.Text, pArea.Centroid) Then
76:                 MsgBox "The current Taxlot value (" & frmTaxlotAssignment.txtTaxlotNum.Text & _
                ") is not unique withing this MapIndex.  Please enter a new number"
78:                 m_pEditor.AbortOperation
                Exit Sub
80:             End If
    
82:             lCount = 1
83:             Do
              ' get the geometry
85:               Set pGeom = pCurFeature.ShapeCopy
86:               If lCount = 1 Then ' if its the first feature
87:                 Set pTmpGeom = pGeom
88:               Else ' merge the geometry of the features
89:                 Set pTopoOperator = pTmpGeom
90:                 Set pOutputGeometry = pTopoOperator.Union(pGeom)
91:                 Set pTmpGeom = pOutputGeometry
92:               End If
                  
              ' now go through each field, if it has a domain associated with it, then
              ' evaluate the merge policy...
              Dim pFld As IField
              Dim pDomain As IDomain
              Dim pSubtypes As ISubtypes
99:               Set pSubtypes = pFeatcls
              Dim i As Long
101:               For i = 0 To pFlds.FieldCount - 1
102:                 Set pFld = pFlds.Field(i)
103:                 Set pDomain = pSubtypes.Domain(lSCode, pFld.Name)
104:                 If Not pDomain Is Nothing Then
                  Select Case pDomain.MergePolicy
                    Case esriMPTSumValues 'Sum values
107:                       If lCount = 1 Then
108:                         pNewFeature.Value(i) = pCurFeature.Value(i)
109:                       Else
110:                         pNewFeature.Value(i) = pNewFeature.Value(i) + pCurFeature.Value(i)
111:                       End If
                    Case esriMPTAreaWeighted 'Area/length weighted average
113:                       If lCount = 1 Then
114:                         pNewFeature.Value(i) = pCurFeature.Value(i) * (getGeomVal(pCurFeature) / lGTotalVal)
115:                       Else
116:                         pNewFeature.Value(i) = pNewFeature.Value(i) + (pCurFeature.Value(i) * (getGeomVal(pCurFeature) / lGTotalVal))
117:                       End If
                    Case Else 'If no merge policy, just take one of the existing values
119:                         pNewFeature.Value(i) = pCurFeature.Value(i)
120:                     End Select 'do not need a case for default value as it is set above
121:                 Else 'If not a domain, copy the existing value
122:                     If pNewFeature.Fields.Field(i).Editable Then 'Don't attempt to copy objectid or other non-editable field
123:                         pNewFeature.Value(i) = pCurFeature.Value(i)
124:                     End If
125:                 End If
126:               Next i
127:               pCurFeature.Delete ' delete the feature
              
129:               Set pCurFeature = pFeatCur.NextFeature
130:               lCount = lCount + 1
131:             Loop Until pCurFeature Is Nothing
            
133:             Set pNewFeature.Shape = pOutputGeometry
            
            'Set taxlot number
            Dim lTLTaxlotFld As Long
137:             lTLTaxlotFld = modUtils.LocateFields(pFeatureLayer.FeatureClass, g_pFldnames.TLTaxlotFN)
138:             pNewFeature.Value(lTLTaxlotFld) = Me.txtNewTaxlot.Text
            
140:             pNewFeature.Store
            
            ' refresh features
            Dim pRefresh As IInvalidArea
144:             Set pRefresh = New InvalidArea
145:             Set pRefresh.Display = m_pEditor.Display
146:             pRefresh.Add pNewFeature
147:             pRefresh.Invalidate esriAllScreenCaches

            ' select new feature
150:             pMap.ClearSelection
151:             pMap.SelectFeature pFeatureLayer, pNewFeature
            
            'Find the Reference Lines feature class to insert any deleted lines
            Dim pWorkspace As IWorkspace
            Dim pFWorkspace As IFeatureWorkspace
            Dim pRLFclass As IFeatureClass
157:             Set pWorkspace = pDataset.Workspace
158:             Set pFWorkspace = pWorkspace
159:             Set pRLFclass = pFWorkspace.OpenFeatureClass(g_pFldnames.FCReferenceLines)
160:             If pRLFclass Is Nothing Then
                'If feature class not present, don't move lines
162:                 MsgBox "Unable to locate Reference Lines feature class", vbCritical
                Exit Sub
164:             End If
            'Move historical taxlot lines to linetype 33
            Dim pTLLinesLayer As IFeatureLayer
            Dim pTLLinesFC As IFeatureClass
            Dim lLineTypeFld As Long
169:             Set pTLLinesLayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlotLines)
170:             If Not pTLLinesLayer Is Nothing Then
171:                 Set pTLLinesFC = pTLLinesLayer.FeatureClass
172:                 lLineTypeFld = modUtils.LocateFields(pRLFclass, g_pFldnames.TLLinesLineTypeFN)
                Dim pLineFCur As IFeatureCursor
                Dim pMergedGeom As IGeometry
175:                 Set pMergedGeom = pNewFeature.Shape
176:                 Set pLineFCur = modUtils.SpatialQueryForEdit(pTLLinesFC, pMergedGeom, esriSpatialRelContains)
177:                 If Not pLineFCur Is Nothing Then
                    Dim pLineFeat As IFeature
                    Dim pNewLineFeat As IFeature
180:                     Set pLineFeat = pLineFCur.NextFeature
181:                     Do While Not pLineFeat Is Nothing
182:                         Set pNewLineFeat = pRLFclass.CreateFeature
183:                         Set pNewLineFeat.Shape = pLineFeat.ShapeCopy
184:                         pNewLineFeat.Value(lLineTypeFld) = 33
185:                         pNewLineFeat.Store
186:                         pLineFCur.DeleteFeature
                        'pLineFeat.Value(lLineTypeFld) = 33
                        'pLineFCur.UpdateFeature pLineFeat
189:                         Set pLineFeat = pLineFCur.NextFeature
190:                     Loop
191:                 End If
192:             End If
            ' finish edit operation
194:             m_pEditor.StopOperation ("Features merged")
195:         End If
196:     End If

198:     Unload Me
  Exit Sub
ErrorHandler:
  HandleError True, "cmdApply_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdHelp_Click()
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
208:     sFilePath = app.Path & "\" & "Combine_help.rtf"
209:     If modUtils.FileExists(sFilePath) Then
210:         If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
211:             Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
212:         End If
213:     Else
214:         MsgBox "No help available"
215:     End If
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
220:     Set m_pApp = New AppRef
221:     Set m_pMxDoc = m_pApp.Document
    'Set a reference to the Editor
    Dim pUID As New UID
224:     pUID = "esriEditor.editor"
225:     Set m_pEditor = g_pApp.FindExtensionByCLSID(pUID)

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Public Function getGeomVal(pFeature As IFeature) As Double
  On Error GoTo ErrorHandler

  ' helper function to get the area/length/perimeter of a feature
  Dim pFC As IFeatureClass
237:   Set pFC = pFeature.Class
  Dim pvFlds As IFields
239:   Set pvFlds = pFC.Fields
  
241:   If pFC.ShapeType = esriGeometryMultipoint Or pFC.ShapeType = esriGeometryMultipoint Or pFC.ShapeType = esriGeometryNull Then
242:     getGeomVal = 0
243:   ElseIf pFC.ShapeType = esriGeometryPolygon Then
244:     getGeomVal = pFeature.Value(pvFlds.FindField(pFC.AreaField.Name))
245:   Else
246:     getGeomVal = pFeature.Value(pvFlds.FindField(pFC.LengthField.Name))
247:   End If

  Exit Function
ErrorHandler:
  HandleError True, "getGeomVal " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

