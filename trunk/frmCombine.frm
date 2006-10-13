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
'
' File name:            frmCombine
'
' Initial Author:       Type your name here
'
' Date Created:     10/11/2006
'
' Description: FORM USED TO COMBINE SELECTED TAXLOTS
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
Private m_pEditor As IEditor
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument
'------------------------------
'Private Constants and Enums
'------------------------------
' Variables used by the Error handler function - DO NOT REMOVE
Private Const c_sModuleFileName As String = "frmCombine.frm"
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
'Description:   Combines taxlot polygons
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:    None
'Outputs:       What variables are changed in this routine?
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006      Initial creation
'***************************************************************************
Private Sub cmdApply_Click()
  On Error GoTo ErrorHandler
    'Code that combines taxlots
    Dim pMXDoc As IMxDocument
    Dim pMap As IMap
102:     Set pMXDoc = g_pApp.Document
103:     Set pMap = pMXDoc.FocusMap
    'Validate new taxlot number entered and make sure it doesn't exist
105:     If Not IsNumeric(Me.txtNewTaxlot.Text) Or Not (Len(Me.txtNewTaxlot.Text) = 5) Then
106:       MsgBox "Invalid Start Value.  Please enter a 5-digit number", vbOKOnly, "Error"
107:       Me.txtNewTaxlot.SetFocus
108:       GoTo Process_Exit
109:     End If

    'Taxlots already selected and taxlot number known
    Dim pFeatcls As IFeatureClass
    Dim pWorkspaceEdit As IWorkspaceEdit
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
116:     Set pFeatureLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
117:     Set pFeatcls = pFeatureLayer.FeatureClass
118:     Set pDataset = pFeatureLayer.FeatureClass
119:     If pDataset Is Nothing Then GoTo Process_Exit
120:     Set pWorkspaceEdit = pDataset.Workspace
121:     If pWorkspaceEdit.IsBeingEdited Then 'Check if being edited
        Dim pFeatCur As IFeatureCursor
123:         Set pFeatCur = modUtils.GetSelectedFeatures(pFeatureLayer) 'Make sure more than one selected
124:         If Not pFeatCur Is Nothing Then
            'Combine taxlots
            ' code to merge the features, evaluate the merge rules and assign values to fields appropriatly
            
            ' start edit operation
129:             m_pEditor.StartOperation
            
            ' create a new feature to be the merge feature
            Dim pCurFeature As IFeature
            Dim pNewFeature As IFeature
            Dim lCount As Long
135:             Set pNewFeature = pFeatcls.CreateFeature
            
            ' create the new geometry.
            Dim pGeom As IGeometry
            Dim pTmpGeom As IGeometry
            Dim pOutputGeometry As IGeometry
            Dim pTopoOperator As ITopologicalOperator
              
            ' initialize the default values for the new feature
            Dim pOutRSType As IRowSubtypes
145:             Set pOutRSType = pNewFeature
146:             If lSCode <> 0 Then
147:               pOutRSType.SubtypeCode = lSCode
148:             End If
149:             pOutRSType.InitDefaultValues
            
            ' get the first feature
152:             Set pCurFeature = pFeatCur.NextFeature
            Dim pFlds As IFields
154:             Set pFlds = pFeatcls.Fields
            
            Dim pArea As IArea
157:             Set pArea = pCurFeature.Shape
            'Now that we have a feature,
            'Verify that within this map index, this taxlot number is unique
            'If not unique, prompt user to enter a new value
161:             If Not modUtils.ValidateTaxlotNum(frmCombine.txtNewTaxlot.Text, pArea.Centroid) Then
162:                 MsgBox "The current Taxlot value (" & frmTaxlotAssignment.txtTaxlotNum.Text & _
                ") is not unique withing this MapIndex.  Please enter a new number"
164:                 m_pEditor.AbortOperation
165:                 GoTo Process_Exit
166:             End If
    
168:             lCount = 1
169:             Do
              ' get the geometry
171:               Set pGeom = pCurFeature.ShapeCopy
172:               If lCount = 1 Then ' if its the first feature
173:                 Set pTmpGeom = pGeom
174:               Else ' merge the geometry of the features
175:                 Set pTopoOperator = pTmpGeom
176:                 Set pOutputGeometry = pTopoOperator.Union(pGeom)
177:                 Set pTmpGeom = pOutputGeometry
178:               End If
                  
              ' now go through each field, if it has a domain associated with it, then
              ' evaluate the merge policy...
              Dim pFld As IField
              Dim pDomain As IDomain
              Dim pSubtypes As ISubtypes
185:               Set pSubtypes = pFeatcls
              Dim i As Long
187:               For i = 0 To pFlds.FieldCount - 1
188:                 Set pFld = pFlds.Field(i)
189:                 Set pDomain = pSubtypes.Domain(lSCode, pFld.Name)
190:                 If Not pDomain Is Nothing Then
                  Select Case pDomain.MergePolicy
                    Case esriMPTSumValues 'Sum values
193:                       If lCount = 1 Then
194:                         pNewFeature.Value(i) = pCurFeature.Value(i)
195:                       Else
196:                         pNewFeature.Value(i) = pNewFeature.Value(i) + pCurFeature.Value(i)
197:                       End If
                    Case esriMPTAreaWeighted 'Area/length weighted average
199:                       If lCount = 1 Then
200:                         pNewFeature.Value(i) = pCurFeature.Value(i) * (getGeomVal(pCurFeature) / lGTotalVal)
201:                       Else
202:                         pNewFeature.Value(i) = pNewFeature.Value(i) + (pCurFeature.Value(i) * (getGeomVal(pCurFeature) / lGTotalVal))
203:                       End If
                    Case Else 'If no merge policy, just take one of the existing values
205:                         pNewFeature.Value(i) = pCurFeature.Value(i)
206:                     End Select 'do not need a case for default value as it is set above
207:                 Else 'If not a domain, copy the existing value
208:                     If pNewFeature.Fields.Field(i).Editable Then 'Don't attempt to copy objectid or other non-editable field
209:                         pNewFeature.Value(i) = pCurFeature.Value(i)
210:                     End If
211:                 End If
212:               Next i
213:               pCurFeature.Delete ' delete the feature
              
215:               Set pCurFeature = pFeatCur.NextFeature
216:               lCount = lCount + 1
217:             Loop Until pCurFeature Is Nothing
            
219:             Set pNewFeature.Shape = pOutputGeometry
            
            'Set taxlot number
            Dim lTLTaxlotFld As Long
223:             lTLTaxlotFld = modUtils.LocateFields(pFeatureLayer.FeatureClass, g_pFldnames.TLTaxlotFN)
224:             pNewFeature.Value(lTLTaxlotFld) = Me.txtNewTaxlot.Text
            
226:             pNewFeature.Store
            
            ' refresh features
            Dim pRefresh As IInvalidArea
230:             Set pRefresh = New InvalidArea
231:             Set pRefresh.Display = m_pEditor.Display
232:             pRefresh.Add pNewFeature
233:             pRefresh.Invalidate esriAllScreenCaches

            ' select new feature
236:             pMap.ClearSelection
237:             pMap.SelectFeature pFeatureLayer, pNewFeature
            
            'Find the Reference Lines feature class to insert any deleted lines
            Dim pWorkspace As IWorkspace
            Dim pFWorkspace As IFeatureWorkspace
            Dim pRLFclass As IFeatureClass
243:             Set pWorkspace = pDataset.Workspace
244:             Set pFWorkspace = pWorkspace
245:             Set pRLFclass = pFWorkspace.OpenFeatureClass(g_pFldnames.FCReferenceLines)
246:             If pRLFclass Is Nothing Then
                'If feature class not present, don't move lines
248:                 MsgBox "Unable to locate Reference Lines feature class", vbCritical
249:                 GoTo Process_Exit
250:             End If
            'Move historical taxlot lines to linetype 33
            Dim pTLLinesLayer As IFeatureLayer
            Dim pTLLinesFC As IFeatureClass
            Dim lLineTypeFld As Long
255:             Set pTLLinesLayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlotLines)
256:             If Not pTLLinesLayer Is Nothing Then
257:                 Set pTLLinesFC = pTLLinesLayer.FeatureClass
258:                 lLineTypeFld = modUtils.LocateFields(pRLFclass, g_pFldnames.TLLinesLineTypeFN)
                Dim pLineFCur As IFeatureCursor
                Dim pMergedGeom As IGeometry
261:                 Set pMergedGeom = pNewFeature.Shape
262:                 Set pLineFCur = modUtils.SpatialQueryForEdit(pTLLinesFC, pMergedGeom, esriSpatialRelContains)
263:                 If Not pLineFCur Is Nothing Then
                    Dim pLineFeat As IFeature
                    Dim pNewLineFeat As IFeature
266:                     Set pLineFeat = pLineFCur.NextFeature
267:                     Do While Not pLineFeat Is Nothing
268:                         Set pNewLineFeat = pRLFclass.CreateFeature
269:                         Set pNewLineFeat.Shape = pLineFeat.ShapeCopy
270:                         pNewLineFeat.Value(lLineTypeFld) = 33
271:                         pNewLineFeat.Store
272:                         pLineFCur.DeleteFeature
                        'pLineFeat.Value(lLineTypeFld) = 33
                        'pLineFCur.UpdateFeature pLineFeat
275:                         Set pLineFeat = pLineFCur.NextFeature
276:                     Loop
277:                 End If
278:             End If
            ' finish edit operation
280:             m_pEditor.StopOperation ("Features merged")
281:         End If
282:     End If

284:     Unload Me
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "cmdApply_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  cmdHelp_Click
'Initial Author:
'Subsequent Author:     James Moore
'Created:
'Description:   Type the description of the function here.
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
Private Sub cmdHelp_Click()
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
314:     sFilePath = app.Path & "\" & "Combine_help.rtf"
315:     If modUtils.FileExists(sFilePath) Then
316:     Debug.Assert True 'need a different method to open rtf files
317:         If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
318:             Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
319:         End If
320:     Else
321:         MsgBox "No help available"
322:     End If
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
327:     Set m_pApp = New AppRef
328:     Set m_pMxDoc = m_pApp.Document
    'Set a reference to the Editor
    Dim pUID As New UID
331:     pUID = "esriEditor.editor"
332:     Set m_pEditor = g_pApp.FindExtensionByCLSID(pUID)

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  getGeomVal
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:   helper function to get the area/length/perimeter of a feature
'Called From:   cmb_Apply_Click()
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:       The area or length or perimeter of the feature or zero if not a valid feature type
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'JWM            10/11/2006  Replaced if statement with select case to improve readability
'                           because the if statement was checking for multipoints twice
'***************************************************************************
Public Function getGeomVal(ByRef pFeature As IFeature) As Double
  On Error GoTo ErrorHandler

  Dim pFC As IFeatureClass
365:   Set pFC = pFeature.Class
  Dim pvFlds As IFields
367:   Set pvFlds = pFC.Fields
  
'++ START JWM 10/11/2006 us
Select Case pFC.ShapeType
    Case esriGeometryMultipoint, esriGeometryNull
372:         getGeomVal = 0
    Case esriGeometryPolygon
374:         getGeomVal = pFeature.Value(pvFlds.FindField(pFC.AreaField.Name))
    Case Else
376:         getGeomVal = pFeature.Value(pvFlds.FindField(pFC.LengthField.Name))
377: End Select

'  If pFC.ShapeType = esriGeometryMultipoint Or pFC.ShapeType = esriGeometryMultipoint Or pFC.ShapeType = esriGeometryNull Then
'    getGeomVal = 0
'  ElseIf pFC.ShapeType = esriGeometryPolygon Then
'    getGeomVal = pFeature.Value(pvFlds.FindField(pFC.AreaField.Name))
'  Else
'    getGeomVal = pFeature.Value(pvFlds.FindField(pFC.LengthField.Name))
'  End If
'++ END JWM 10/11/2006

  Exit Function
ErrorHandler:
  HandleError True, "getGeomVal " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

