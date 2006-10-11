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
    Set pMXDoc = g_pApp.Document
    Set pMap = pMXDoc.FocusMap
    'Validate new taxlot number entered and make sure it doesn't exist
    If Not IsNumeric(Me.txtNewTaxlot.Text) Or Not (Len(Me.txtNewTaxlot.Text) = 5) Then
      MsgBox "Invalid Start Value.  Please enter a 5-digit number", vbOKOnly, "Error"
      Me.txtNewTaxlot.SetFocus
      GoTo Process_Exit
    End If

    'Taxlots already selected and taxlot number known
    Dim pFeatcls As IFeatureClass
    Dim pWorkspaceEdit As IWorkspaceEdit
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Set pFeatureLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    Set pFeatcls = pFeatureLayer.FeatureClass
    Set pDataset = pFeatureLayer.FeatureClass
    If pDataset Is Nothing Then GoTo Process_Exit
    Set pWorkspaceEdit = pDataset.Workspace
    If pWorkspaceEdit.IsBeingEdited Then 'Check if being edited
        Dim pFeatCur As IFeatureCursor
        Set pFeatCur = modUtils.GetSelectedFeatures(pFeatureLayer) 'Make sure more than one selected
        If Not pFeatCur Is Nothing Then
            'Combine taxlots
            ' code to merge the features, evaluate the merge rules and assign values to fields appropriatly
            
            ' start edit operation
            m_pEditor.StartOperation
            
            ' create a new feature to be the merge feature
            Dim pCurFeature As IFeature
            Dim pNewFeature As IFeature
            Dim lCount As Long
            Set pNewFeature = pFeatcls.CreateFeature
            
            ' create the new geometry.
            Dim pGeom As IGeometry
            Dim pTmpGeom As IGeometry
            Dim pOutputGeometry As IGeometry
            Dim pTopoOperator As ITopologicalOperator
              
            ' initialize the default values for the new feature
            Dim pOutRSType As IRowSubtypes
            Set pOutRSType = pNewFeature
            If lSCode <> 0 Then
              pOutRSType.SubtypeCode = lSCode
            End If
            pOutRSType.InitDefaultValues
            
            ' get the first feature
            Set pCurFeature = pFeatCur.NextFeature
            Dim pFlds As IFields
            Set pFlds = pFeatcls.Fields
            
            Dim pArea As IArea
            Set pArea = pCurFeature.Shape
            'Now that we have a feature,
            'Verify that within this map index, this taxlot number is unique
            'If not unique, prompt user to enter a new value
            If Not modUtils.ValidateTaxlotNum(frmCombine.txtNewTaxlot.Text, pArea.Centroid) Then
                MsgBox "The current Taxlot value (" & frmTaxlotAssignment.txtTaxlotNum.Text & _
                ") is not unique withing this MapIndex.  Please enter a new number"
                m_pEditor.AbortOperation
                GoTo Process_Exit
            End If
    
            lCount = 1
            Do
              ' get the geometry
              Set pGeom = pCurFeature.ShapeCopy
              If lCount = 1 Then ' if its the first feature
                Set pTmpGeom = pGeom
              Else ' merge the geometry of the features
                Set pTopoOperator = pTmpGeom
                Set pOutputGeometry = pTopoOperator.Union(pGeom)
                Set pTmpGeom = pOutputGeometry
              End If
                  
              ' now go through each field, if it has a domain associated with it, then
              ' evaluate the merge policy...
              Dim pFld As IField
              Dim pDomain As IDomain
              Dim pSubtypes As ISubtypes
              Set pSubtypes = pFeatcls
              Dim i As Long
              For i = 0 To pFlds.FieldCount - 1
                Set pFld = pFlds.Field(i)
                Set pDomain = pSubtypes.Domain(lSCode, pFld.Name)
                If Not pDomain Is Nothing Then
                  Select Case pDomain.MergePolicy
                    Case esriMPTSumValues 'Sum values
                      If lCount = 1 Then
                        pNewFeature.Value(i) = pCurFeature.Value(i)
                      Else
                        pNewFeature.Value(i) = pNewFeature.Value(i) + pCurFeature.Value(i)
                      End If
                    Case esriMPTAreaWeighted 'Area/length weighted average
                      If lCount = 1 Then
                        pNewFeature.Value(i) = pCurFeature.Value(i) * (getGeomVal(pCurFeature) / lGTotalVal)
                      Else
                        pNewFeature.Value(i) = pNewFeature.Value(i) + (pCurFeature.Value(i) * (getGeomVal(pCurFeature) / lGTotalVal))
                      End If
                    Case Else 'If no merge policy, just take one of the existing values
                        pNewFeature.Value(i) = pCurFeature.Value(i)
                    End Select 'do not need a case for default value as it is set above
                Else 'If not a domain, copy the existing value
                    If pNewFeature.Fields.Field(i).Editable Then 'Don't attempt to copy objectid or other non-editable field
                        pNewFeature.Value(i) = pCurFeature.Value(i)
                    End If
                End If
              Next i
              pCurFeature.Delete ' delete the feature
              
              Set pCurFeature = pFeatCur.NextFeature
              lCount = lCount + 1
            Loop Until pCurFeature Is Nothing
            
            Set pNewFeature.Shape = pOutputGeometry
            
            'Set taxlot number
            Dim lTLTaxlotFld As Long
            lTLTaxlotFld = modUtils.LocateFields(pFeatureLayer.FeatureClass, g_pFldnames.TLTaxlotFN)
            pNewFeature.Value(lTLTaxlotFld) = Me.txtNewTaxlot.Text
            
            pNewFeature.Store
            
            ' refresh features
            Dim pRefresh As IInvalidArea
            Set pRefresh = New InvalidArea
            Set pRefresh.Display = m_pEditor.Display
            pRefresh.Add pNewFeature
            pRefresh.Invalidate esriAllScreenCaches

            ' select new feature
            pMap.ClearSelection
            pMap.SelectFeature pFeatureLayer, pNewFeature
            
            'Find the Reference Lines feature class to insert any deleted lines
            Dim pWorkspace As IWorkspace
            Dim pFWorkspace As IFeatureWorkspace
            Dim pRLFclass As IFeatureClass
            Set pWorkspace = pDataset.Workspace
            Set pFWorkspace = pWorkspace
            Set pRLFclass = pFWorkspace.OpenFeatureClass(g_pFldnames.FCReferenceLines)
            If pRLFclass Is Nothing Then
                'If feature class not present, don't move lines
                MsgBox "Unable to locate Reference Lines feature class", vbCritical
                GoTo Process_Exit
            End If
            'Move historical taxlot lines to linetype 33
            Dim pTLLinesLayer As IFeatureLayer
            Dim pTLLinesFC As IFeatureClass
            Dim lLineTypeFld As Long
            Set pTLLinesLayer = modUtils.FindFeatureLayerByDS(g_pFldnames.FCTaxlotLines)
            If Not pTLLinesLayer Is Nothing Then
                Set pTLLinesFC = pTLLinesLayer.FeatureClass
                lLineTypeFld = modUtils.LocateFields(pRLFclass, g_pFldnames.TLLinesLineTypeFN)
                Dim pLineFCur As IFeatureCursor
                Dim pMergedGeom As IGeometry
                Set pMergedGeom = pNewFeature.Shape
                Set pLineFCur = modUtils.SpatialQueryForEdit(pTLLinesFC, pMergedGeom, esriSpatialRelContains)
                If Not pLineFCur Is Nothing Then
                    Dim pLineFeat As IFeature
                    Dim pNewLineFeat As IFeature
                    Set pLineFeat = pLineFCur.NextFeature
                    Do While Not pLineFeat Is Nothing
                        Set pNewLineFeat = pRLFclass.CreateFeature
                        Set pNewLineFeat.Shape = pLineFeat.ShapeCopy
                        pNewLineFeat.Value(lLineTypeFld) = 33
                        pNewLineFeat.Store
                        pLineFCur.DeleteFeature
                        'pLineFeat.Value(lLineTypeFld) = 33
                        'pLineFCur.UpdateFeature pLineFeat
                        Set pLineFeat = pLineFCur.NextFeature
                    Loop
                End If
            End If
            ' finish edit operation
            m_pEditor.StopOperation ("Features merged")
        End If
    End If

    Unload Me
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
    sFilePath = app.Path & "\" & "Combine_help.rtf"
    If modUtils.FileExists(sFilePath) Then
    Debug.Assert True 'need a different method to open rtf files
        If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
            Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
        End If
    Else
        MsgBox "No help available"
    End If
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
    Set m_pApp = New AppRef
    Set m_pMxDoc = m_pApp.Document
    'Set a reference to the Editor
    Dim pUID As New UID
    pUID = "esriEditor.editor"
    Set m_pEditor = g_pApp.FindExtensionByCLSID(pUID)

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
  Set pFC = pFeature.Class
  Dim pvFlds As IFields
  Set pvFlds = pFC.Fields
  
'++ START JWM 10/11/2006 us
Select Case pFC.ShapeType
    Case esriGeometryMultipoint, esriGeometryNull
        getGeomVal = 0
    Case esriGeometryPolygon
        getGeomVal = pFeature.Value(pvFlds.FindField(pFC.AreaField.Name))
    Case Else
        getGeomVal = pFeature.Value(pvFlds.FindField(pFC.LengthField.Name))
End Select

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

