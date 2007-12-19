VERSION 5.00
Begin VB.Form frmCombine 
   Caption         =   "Taxlot Combine"
   ClientHeight    =   1215
   ClientLeft      =   4770
   ClientTop       =   4575
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNewTaxlot 
      Height          =   315
      Left            =   1170
      MaxLength       =   5
      TabIndex        =   3
      Top             =   150
      Width           =   1725
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   690
      Width           =   800
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   690
      Width           =   800
   End
   Begin VB.Label Label1 
      Caption         =   "New Taxlot:"
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmCombine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Copyright (C) 2006  opet developers opet-developers@lists.sourceforge.net
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details located in AppSpecs.bas file.
'
'    You should have received a copy of the GNU General Public License along
'    with this program; if not, write to the Free Software Foundation, Inc.,
'    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
' Keyword expansion for source code control
' Tag for this file : $Name$
' SCC Revision number: $Revision$
' Date of last change: $Date$
'
'
' File name:            frmCombine
'
' Initial Author:       <<Unknown>>
'
' Date Created:         10/11/2006
'
' Description: FORM USED TO COMBINE SELECTED TAXLOTS.
'       Form used by the Combine Taxlot tools for its user interface
'       Portions of this code may have come from the clsMergeRules.cls located in ArcGIS _
'       Developer help
'
' Entry points:
'       Form Object
'
' Dependencies:
'       File References:
'           esriArcMapUI
'           esriCarto
'           esriEditor
'           esriFramework
'           esriGeoDatabase
'           esriGeometry
'           esriSystem
'       File Dependencies
'           basGlobals
'           basUtilities
'
' Issues:
'       None known at this time (2/6/2007 JWalton)
'
' Method:
'       None
'
' Updates:
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)

Option Explicit
'******************************
' Private Definitions
'------------------------------
' Private Variables
'------------------------------
Private m_pEditor As esriEditor.IEditor
'++ START JWalton 2/6/2007
    ' Removed declaration for m_pMxDoc as it is no longer used
    ' Removed declaration for m_pApp in favor of g_pApp
'++ END JWalton 2/6/2007
Private ml_SubtypeCode As Long
Private m_pEnumFeature As IEnumFeature
Private m_lGTotalVal As Double
'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "frmCombine.frm"

'***************************************************************************
'Name:  cmdApply_Click
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Description:   Combines taxlot polygons
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Initial creation
'James Moore    10-30-2006  Some of this code was copied from a developer sample and not fully fleshed out.
'James Moore    01/11/2007  I have fleshed out the code mentioned at beginning of this file and implemented it
'***************************************************************************

Private Sub cmdApply_Click()
On Error GoTo ErrorHandler
    '++ START JWalton 2/6/2007 Centralized Variable Declarations
    Dim pMXDoc As esriArcMapUI.IMxDocument
    Dim pMap As esriCarto.IMap
    Dim pDataset As esriGeoDatabase.IDataset
    Dim pDomain As esriGeoDatabase.IDomain
    Dim pCurFeature As esriGeoDatabase.IFeature
    Dim pLineFeat As esriGeoDatabase.IFeature
    Dim pNewFeature As esriGeoDatabase.IFeature
    Dim pNewLineFeat As esriGeoDatabase.IFeature
    Dim pFWorkspace As esriGeoDatabase.IFeatureWorkspace
    Dim pFeatcls As esriGeoDatabase.IFeatureClass
    Dim pRLFclass As esriGeoDatabase.IFeatureClass
    Dim pTLLinesFC As esriGeoDatabase.IFeatureClass
    Dim pFeatCur As esriGeoDatabase.IFeatureCursor
    Dim pLineFCur As esriGeoDatabase.IFeatureCursor
    Dim pFeatureLayer As esriCarto.IFeatureLayer
    Dim pTLLinesLayer As esriCarto.IFeatureLayer
    Dim pFld As esriGeoDatabase.IField
    Dim pFlds As esriGeoDatabase.IFields
    Dim pRefresh As esriGeoDatabase.IInvalidArea
    Dim pOutRSType As esriGeoDatabase.IRowSubtypes
    Dim pSubtypes As esriGeoDatabase.ISubtypes
    Dim pWorkspace As esriGeoDatabase.IWorkspace
    Dim pWorkspaceEdit As esriGeoDatabase.IWorkspaceEdit
    Dim pArea As esriGeometry.IArea
    Dim pMergedGeom As esriGeometry.IGeometry
    Dim pGeom As esriGeometry.IGeometry
    Dim pOutputGeometry As esriGeometry.IGeometry
    Dim pTmpGeom As esriGeometry.IGeometry
    Dim pTopoOperator As esriGeometry.ITopologicalOperator
    Dim i As Long
    Dim lCount As Long
    Dim lDefaultSubType As Long
    Dim lLineTypeFld As Long
    Dim lTLTaxlotFld As Long
    '++ END JWalton 2/6/2007
    
    ' Initialize objects
    Set pMXDoc = g_pApp.Document
    Set pMap = pMXDoc.FocusMap
    
    'Validate new taxlot number entered and make sure it doesn't exist
    If (Len(Me.txtNewTaxlot.Text) = 0 Or (Len(Me.txtNewTaxlot.Text) > ORMAP_TAXLOT_FIELD_LENGTH)) Or _
       Not IsNumeric(Me.txtNewTaxlot.Text) Then
        MsgBox "Invalid Start Value.  Please enter a 5-digit number", vbOKOnly, "Error"
        Me.txtNewTaxlot.SetFocus
        GoTo Process_Exit
    End If

    'Taxlots already selected and taxlot number known
    Set pFeatureLayer = basUtilities.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    Set pFeatcls = pFeatureLayer.FeatureClass
    Set pDataset = pFeatureLayer.FeatureClass
    If pDataset Is Nothing Then GoTo Process_Exit
    Set pWorkspaceEdit = pDataset.Workspace
    If pWorkspaceEdit.IsBeingEdited Then 'Check if being edited
        Set pFeatCur = basUtilities.GetSelectedFeatures(pFeatureLayer) 'Make sure more than one selected
        If Not pFeatCur Is Nothing Then
            ' Combine taxlots
            ' Code to merge the features, evaluate the merge rules and assign values to fields appropriatly
            
            ' Start edit operation
            m_pEditor.StartOperation
            
            ' create a new feature to be the merge feature
            Set pNewFeature = pFeatcls.CreateFeature
              
            '++ START JWalton 2/14/2007 Extract the default subtype from the feature's class
            ' Initialize the default values for the new feature
            Set pSubtypes = pNewFeature.Class
            lDefaultSubType = pSubtypes.DefaultSubtypeCode
            '++ END JWalton 2/14/2007
            
            Set pOutRSType = pNewFeature
            
'++ START merge policy revisited JWM 01/11/2007 I have removed previous my previous code and
'have implemented the code from the developer sample clsMergeRules.cls as best as I am able
            If ml_SubtypeCode <> 0 Then
              pOutRSType.SubtypeCode = ml_SubtypeCode
            End If
            pOutRSType.InitDefaultValues
'++ END merge policy revisted JWM 01/11/2007

            ' get the first feature
            Set pFeatCur = basUtilities.GetSelectedFeatures(pFeatureLayer)
            Set pCurFeature = pFeatCur.NextFeature
            Set pFlds = pFeatcls.Fields
            
            Set pArea = pCurFeature.Shape
            ' Now that we have a feature,
            ' Verify that within this map index, this taxlot number is unique
            ' If not unique, prompt user to enter a new value
            If Not basUtilities.ValidateTaxlotNum(Me.txtNewTaxlot.Text, pArea.Centroid) Then
                MsgBox "The current Taxlot value (" & Me.txtNewTaxlot.Text & _
                ") is not unique within this MapIndex.  Please enter a new number"
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
                Set pSubtypes = pFeatcls
                For i = 0 To pFlds.FieldCount - 1
                    Set pFld = pFlds.Field(i)
                    Set pDomain = pSubtypes.Domain(ml_SubtypeCode, pFld.Name)
                    If Not pDomain Is Nothing Then
                      Select Case pDomain.MergePolicy
                            Case esriGeoDatabase.esriMergePolicyType.esriMPTSumValues 'Sum values
                                If lCount = 1 Then
                                    pNewFeature.Value(i) = pCurFeature.Value(i)
                                Else
                                    pNewFeature.Value(i) = pNewFeature.Value(i) + pCurFeature.Value(i)
                                End If
                            Case esriGeoDatabase.esriMergePolicyType.esriMPTAreaWeighted 'Area/length weighted average
                                If lCount = 1 Then
                                    pNewFeature.Value(i) = pCurFeature.Value(i) * (GetGeomVal(pCurFeature) / m_lGTotalVal)
                                Else
                                    pNewFeature.Value(i) = pNewFeature.Value(i) + (pCurFeature.Value(i) * (GetGeomVal(pCurFeature) / m_lGTotalVal))
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
            lTLTaxlotFld = basUtilities.LocateFields(pFeatureLayer.FeatureClass, g_pFldnames.TLTaxlotFN)
            pNewFeature.Value(lTLTaxlotFld) = Me.txtNewTaxlot.Text
            
            pNewFeature.Store
            
            ' refresh features
            Set pRefresh = New esriCarto.InvalidArea
            Set pRefresh.Display = m_pEditor.Display
            pRefresh.Add pNewFeature
            pRefresh.Invalidate esriDisplay.esriScreenCache.esriAllScreenCaches

            ' select new feature
            pMap.ClearSelection
            pMap.SelectFeature pFeatureLayer, pNewFeature
            
            'Find the Reference Lines feature class to insert any deleted lines
            Set pWorkspace = pDataset.Workspace
            Set pFWorkspace = pWorkspace
            Set pRLFclass = pFWorkspace.OpenFeatureClass(g_pFldnames.FCReferenceLines)
            If pRLFclass Is Nothing Then
                'If feature class not present, don't move lines
                MsgBox "Unable to locate Reference Lines feature class", vbCritical
                GoTo Process_Exit
            End If
            'Move historical taxlot lines to linetype 33
            Set pTLLinesLayer = basUtilities.FindFeatureLayerByDS(g_pFldnames.FCTaxlotLines)
            If Not pTLLinesLayer Is Nothing Then
                Set pTLLinesFC = pTLLinesLayer.FeatureClass
                lLineTypeFld = basUtilities.LocateFields(pRLFclass, g_pFldnames.TLLinesLineTypeFN)
                Set pMergedGeom = pNewFeature.Shape
                Set pLineFCur = basUtilities.SpatialQueryForEdit(pTLLinesFC, pMergedGeom, esriSpatialRelContains)
                If Not pLineFCur Is Nothing Then
                    Set pLineFeat = pLineFCur.NextFeature
                    Do While Not pLineFeat Is Nothing
                        Set pNewLineFeat = pRLFclass.CreateFeature
                        Set pNewLineFeat.Shape = pLineFeat.ShapeCopy
                        pNewLineFeat.Value(lLineTypeFld) = 33
                        pNewLineFeat.Store
                        pLineFCur.DeleteFeature
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
  HandleError True, _
              "cmdApply_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
              Err.Number, _
              Err.Source, _
              Err.Description, _
              4
    Stop
    Resume
End Sub

'++ START JWalton 1/29/2007
    ' Removed cmdCancel_Click() Routine - No longer necessary
'++ END JWalton 1/29/2007

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
'JWM            10/16/2006      using new method to open help file
'***************************************************************************

Private Sub cmdHelp_Click()
    '++ START JWalton 2/6/2007 Centralized Variable Declarations
    ' Variable declarations
    Dim sFilePath As String
    '++ END JWalton 2/6/2007
     
    ' Opens a custom help file if it exists
    sFilePath = app.Path & "\" & "Combine_help.rtf"
    If basUtilities.FileExists(sFilePath) Then
        '++ START JWM 10/16/2006 using new method to open help file
        basUtilities.gsb_StartDoc Me.hwnd, sFilePath
        '++ START/END JWM 10/16/2006
      Else
        MsgBox "No help file available in current directory", vbOKOnly + vbInformation
    End If
End Sub

'***************************************************************************
'Name:                  Form_Load
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Event Handler
'Called From:   System when form is loaded
'Description:   OnLoad even handler
'Methods:       Registers the status of the form with the class collection
'               g_pForms.
'Inputs:        None
'Parameters:    None
'Outputs:       m_pEditor as IEditor
'               g_pForms as clsFormsCatalog
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'**************************************************************************

Private Sub Form_Load()
On Error GoTo ErrorHandler
    '++ START JWalton 2/6/2007 Centralized Variable Declarations
    Dim pUID As New esriSystem.UID
    '++ END JWalton 2/6/2007
    
    'Set a reference to the Editor
    pUID = "esriEditor.editor"
    Set m_pEditor = g_pApp.FindExtensionByCLSID(pUID)
    
    '++ START JWalton 1/29/2007
    ' Sets the form status to open
    g_pForms.SetFormStatus Me.Name, True
    '++ END JWalton 1/29/2007

'++ START merge policy  JWM 01/11/2007
    If m_pEditor.SelectionCount > 1 Then
        Dim pChkFeature As IFeature
        Dim pRowSubtypes As IRowSubtypes
        
        Set m_pEnumFeature = m_pEditor.EditSelection
        Set pChkFeature = m_pEnumFeature.Next
        Set pRowSubtypes = pChkFeature
        ml_SubtypeCode = pRowSubtypes.SubtypeCode
        Do
            Set pRowSubtypes = pChkFeature
            If pRowSubtypes.SubtypeCode = ml_SubtypeCode Then
                m_lGTotalVal = m_lGTotalVal + GetGeomVal(pChkFeature)
            End If
            Set pChkFeature = m_pEnumFeature.Next
        Loop Until pChkFeature Is Nothing
        Set pRowSubtypes = Nothing
        Set pChkFeature = Nothing
    End If
'++ END JWM 01/11/2007
  Exit Sub
ErrorHandler:
  HandleError True, _
              "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
              Err.Number, _
              Err.Source, _
              Err.Description, _
              4
End Sub

'***************************************************************************
'Name:                  Query_Unload
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Event Handler
'Called From:   System when the form is unloaded
'Description:   OnQueryOnload event handler
'Methods:       Registers the status of the form with the class collection
'               g_pForms.
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

'++ START JWalton 1/29/2007
Private Sub Form_QueryUnload( _
  Cancel As Integer, _
  UnloadMode As Integer)
    ' Sets the form status to not open
    g_pForms.SetFormStatus Me.Name, False
End Sub
'++ END JWalton 1/29/2007

'***************************************************************************
'Name:  GetGeomVal
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Helper function to get the area/length/perimeter of a feature
'Called From:   cmb_Apply_Click()
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:    None
'Outputs:       What variables are changed in this routine?
'Returns:       The area or length or perimeter of the feature or zero if
'               not a valid feature type
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'JWM            10/11/2006  Replaced if statement with select case to improve
'                           readability also the if statement was checking
'                           for multipoints types twice
'***************************************************************************

Public Function GetGeomVal( _
  ByRef pFeature As esriGeoDatabase.IFeature) As Double
On Error GoTo ErrorHandler
    '++ START JWalton 2/6/2007 Centralized Variable Declarations
    Dim pFC As esriGeoDatabase.IFeatureClass
    Dim pvFlds As esriGeoDatabase.IFields
    '++ END JWalton 2/6/2007
    
    ' Initialize objects
    Set pFC = pFeature.Class
    Set pvFlds = pFC.Fields
    
    '++ START JWM 10/11/2006
    Select Case pFC.ShapeType
      Case esriGeometryMultipoint, esriGeometryNull
        GetGeomVal = 0
      Case esriGeometryPolygon
        GetGeomVal = pFeature.Value(pvFlds.FindField(pFC.AreaField.Name))
      Case Else
        GetGeomVal = pFeature.Value(pvFlds.FindField(pFC.LengthField.Name))
    End Select
    '++ END JWM 10/11/2006

  Exit Function
ErrorHandler:
  HandleError True, _
              "getGeomVal " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
              Err.Number, _
              Err.Source, _
              Err.Description, _
              4
End Function
