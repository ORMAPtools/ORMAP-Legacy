VERSION 5.00
Begin VB.Form frmLocate 
   Caption         =   "Locate"
   ClientHeight    =   2235
   ClientLeft      =   3270
   ClientTop       =   4605
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   4335
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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
         Weight          =   400
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
         Weight          =   400
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
  Set pQueryFilter = New QueryFilter
  
  sMsg = "Please enter a Map Number and optionally, a Taxlot number"
  
'++ START JWM 10/11/2006 trim and then test for length
  sTLNum = Trim$(frmLocate.txtTaxlot)
  sMNum = Trim$(frmLocate.cmbMapNumber.Text)
  
If Len(sTLNum) = 0 Then 'Just Query MapIndex
'  If sTLNum = "" Then
    If Len(sMNum) = 0 Then 'If both empty
'    If sMNum = "" Then
        MsgBox sMsg, vbOKOnly
        GoTo Process_Exit
    Else 'MapNumber entered with no Taxlot
        Set pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        If pMIFlayer Is Nothing Then
              MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
              "This process requires a feature class called " & g_pFldnames.FCMapIndex
              GoTo Process_Exit
        End If
        Set pMIFclass = pMIFlayer.FeatureClass
        pQueryFilter.whereClause = "[" & g_pFldnames.MIMapNumberFN & "] = '" & frmLocate.cmbMapNumber.Text & "'"
        Set pFeatureCursor = pMIFclass.Search(pQueryFilter, False)
    End If
  ElseIf Len(sTLNum) > 0 And Len(sMNum) > 0 Then 'Both values entered
        pQueryFilter.whereClause = "[" & g_pFldnames.TLMapNumberFN & "] = '" & frmLocate.cmbMapNumber.Text & "' and [" & g_pFldnames.TLTaxlotFN & "]= '" & frmLocate.txtTaxlot & "'"
        Set pFeatureCursor = m_pTaxlotFClass.Search(pQueryFilter, False)
  Else 'Only a taxlot entered
        MsgBox sMsg, vbOKOnly
        GoTo Process_Exit
    End If
  If pFeatureCursor Is Nothing Then GoTo Process_Exit
  Set pFeature = pFeatureCursor.NextFeature

  If pFeature Is Nothing Then
    If Len(sTLNum) = 0 Then
        MsgBox "Map Index could not be found.", vbInformation, "Try Again"
    Else
        MsgBox "Taxlot could not be found.", vbInformation, "Try Again"
    End If
    '++ END JWM 10/11/2006
    frmLocate.txtTaxlot = ""
    frmLocate.txtTaxlot.SetFocus
    GoTo Process_Exit
  Else
    'Zoom to selected feature
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pFeature.Shape.Envelope
    
    ZoomToExtent pEnvelope, m_pMxDoc

  End If
  
  Unload Me
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "cmdApply_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

    Unload Me

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdHelp_Click()
    Dim sFilePath As String
    sFilePath = app.Path & "\" & "Locate_help.rtf"
    If FileExists(sFilePath) Then
'++ START JWM 10/16/2006 using new method to open help file
        gsb_StartDoc Me.hwnd, sFilePath
'++ START/END JWM 10/16/2006
    Else
        MsgBox "No help file available in current directory.", vbOKOnly + vbInformation
    End If
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
    Set pApp = GetAppRef ' AppRef
    Set m_pMxDoc = pApp.Document
    'Get the MapIndex feature layer and fclass
    Set m_pTaxlotFlayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If m_pTaxlotFlayer Is Nothing Then
        MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        GoTo Proc_Exit
    End If
    Set m_pTaxlotFClass = m_pTaxlotFlayer.FeatureClass
    Set m_pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If m_pMIFlayer Is Nothing Then
        MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
       GoTo Proc_Exit
    End If
    Set m_pMIFclass = m_pMIFlayer.FeatureClass
    'Get fields needed to populate the form
    Set m_pMIFields = m_pMIFclass.Fields

    m_lMapNumFld = LocateFields(m_pMIFclass, g_pFldnames.MIMapNumberFN)
  Dim pQueryDef As IQueryDef
  Dim pRow As IRow
  Dim pCursor As ICursor
  Dim pFeatureWorkspace As IFeatureWorkspace
  Dim pDataset As IDataset
  
  Set pDataset = m_pMIFlayer 'm_pTaxlotFlayer
  
  Dim sFieldName As String
  sFieldName = g_pFldnames.TLMapNumberFN
  
  Set pFeatureWorkspace = pDataset.Workspace
  Set pQueryDef = pFeatureWorkspace.CreateQueryDef
  With pQueryDef
    .Tables = pDataset.Name ' Fully qualified table name
         'Problems with some values -- prevents the form from loading
    .SubFields = "DISTINCT(" & sFieldName & ")"
    Set pCursor = .Evaluate
  End With
  
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    If Not IsNull(pRow.Value(0)) Then
        frmLocate.cmbMapNumber.AddItem pRow.Value(0) ' Note only one field in the cursor
    End If
    Set pRow = pCursor.NextRow
  Loop
  
Proc_Exit:
Exit Sub
Err_Handler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
