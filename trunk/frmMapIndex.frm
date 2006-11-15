VERSION 5.00
Begin VB.Form frmMapIndex 
   Caption         =   "Map Index"
   ClientHeight    =   3720
   ClientLeft      =   1545
   ClientTop       =   4605
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   9555
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

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbCounty_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbQtr_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbQtr_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbQtrQtr_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbQtrQtr_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbRange_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbRange_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbRangeDir_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbRangeDir_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbRangePart_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbRangePart_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbReliability_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbReliability_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbScale_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbScale_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbSection_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbSection_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbSuffType_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbSuffType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbTown_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbTown_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub
Private Sub cmbTownDir_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmbTownDir_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmbTownPart_Click()
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

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
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Initial creation of this comment section
'***************************************************************************
Private Sub cmdAssign_Click()
  On Error GoTo ErrorHandler
    
    Dim bValuesPresent As Boolean
    bValuesPresent = True
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            Dim pTxtBox As TextBox
            Set pTxtBox = ctl
            If Len(pTxtBox.Text) = 0 Then
                If StrComp(pTxtBox.Name, "txtORMAPMapNum", vbTextCompare) <> 0 And StrComp(pTxtBox.Name, "txtPage", vbTextCompare) <> 0 Then
'                If Not pTxtBox.Name = "txtORMAPMapNum" And Not pTxtBox.Name = "txtPage" Then
                    bValuesPresent = False
                End If
            End If
        ElseIf TypeOf ctl Is ComboBox Then
            Dim pCmb As ComboBox
            Set pCmb = ctl
            If Len(pCmb.Text) = 0 Then
               bValuesPresent = False
            End If
        ElseIf TypeOf ctl Is ListBox Then
            'this comment is here to poke us in the eye and ask WHY is this here
            MsgBox "listbox"
        End If
    Next ctl
    
    If Not bValuesPresent Then
       MsgBox "All fields must be filled in before assigning", vbOKOnly
       GoTo Process_Exit
    End If

    'Otherwise, save values
    
    'Save values if necessary
    If Not m_bPossiblyChanged Then GoTo Process_Exit
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
    
    Set pDataset = m_pMIFclass
    Set pWSEdit = pDataset.Workspace
    pWSEdit.StartEditOperation
    
    'Get a Taxlot feature, so its domains can be referenced
    Dim pFeatCur As IFeatureCursor
    Set pFeatCur = m_pTaxlotFlayer.Search(Nothing, True)
    Dim pTLFeat As IFeature
    Set pTLFeat = pFeatCur.NextFeature
    'This functionality was originally set up to work with the MapIndex feature
    'currently being edited.  The db design changed, but the structure of the code
    'has not been changed.  To obtain the domains necessary to display and save
    'values in MapIndex, taxlots are used.  This requires that at least one taxlot
    'feature exist.
    If pTLFeat Is Nothing Then
        MsgBox "No taxlot features present.  Please create at least one taxlot", vbOKOnly
        GoTo Process_Exit
    End If
    
    'MapNumber
    sCurMapNum = Me.txtMapNum.Text
    m_pMIFeat.Value(m_lMapNumFld) = sCurMapNum

    'Reliability
    sCurReliabil = ConvertCode(m_pMIFeat, g_pFldnames.MIReliabFN, Me.cmbReliability)
    m_pMIFeat.Value(m_lReliabFld) = CInt(sCurReliabil)
    
    'Scale
    sCurScale = ConvertCode(m_pMIFeat, g_pFldnames.MIMapScaleFN, Me.cmbScale.Text)
    m_pMIFeat.Value(m_lScaleFld) = CLng(sCurScale)

    'Page
    sCurPage = Me.txtPage.Text
    If IsNumeric(sCurPage) Then
        Dim lCurPage As Long
        lCurPage = CLng(sCurPage)
        m_pMIFeat.Value(m_lPageFld) = lCurPage
    Else
        'If null, enter a null value
        Dim nullVal As Variant
        m_pMIFeat.Value(m_lPageFld) = nullVal
    End If

    'Compile values below into the OrMAPMapNumber value
    'County
    sCurCounty = ConvertCode(pTLFeat, g_pFldnames.TLCountyFN, Me.cmbCounty.Text)
    sCurCounty = FormatOMMapNum(sCurCounty, "county")
    
    'Town
    sCurTown = Me.cmbTown.Text  'ConvertCode(pTLFeat, g_pFldnames.TLTownFN, Me.cmbTown.Text)
    sCurTown = FormatOMMapNum(sCurTown, "town")

    'TownPart
    sCurTownPart = Me.cmbTownPart.Text 'ConvertCode(pTLFeat, g_pFldnames.TLTownPartFN, Me.cmbTownPart.Text)
    sCurTownPart = FormatOMMapNum(sCurTownPart, "townpart")
    'If Len(sCurTownPart) = 3 Then
        'sCurTownPart = Replace(sCurTownPart, ".", "")
        'sCurTownPart = Left(sCurTownPart, 1) & "." & Right(sCurTownPart, 2)
    'End If
    'TownDir
    sCurTownDir = Me.cmbTownDir.Text
    sCurTownDir = FormatOMMapNum(sCurTownDir, "towndir")

    'Range
    sCurRange = Me.cmbRange.Text
    sCurRange = FormatOMMapNum(sCurRange, "range")

    'RangePart
    sCurRangePart = Me.cmbRangePart.Text
    sCurRangePart = FormatOMMapNum(sCurRangePart, "rangepart")
    'If Len(sCurRangePart) = 3 Then
    'sCurRangePart = Replace(sCurRangePart, ".", "")
        'sCurRangePart = Left(sCurRangePart, 1) & "." & Right(sCurRangePart, 2)
    'End If
    'RangeDir
    sCurRangeDir = Me.cmbRangeDir.Text
    sCurRangeDir = FormatOMMapNum(sCurRangeDir, "rangedir")

    'Section
    sCurSection = Me.cmbSection.Text
    sCurSection = FormatOMMapNum(sCurSection, "section")
 
    'Qtr
    sCurQtr = Me.cmbQtr.Text
    sCurQtr = FormatOMMapNum(sCurQtr, "qtr")

    'QtrQtr
    sCurQtrQtr = Me.cmbQtrQtr.Text
    sCurQtrQtr = FormatOMMapNum(sCurQtrQtr, "qtrqtr")

    'MapSuffixType
    sCurSuffixType = ConvertCode(pTLFeat, g_pFldnames.TLSufTypeFN, cmbSufftype.Text)
    sCurSuffixType = FormatOMMapNum(sCurSuffixType, "suffixtype")
    
    'MapSuffixNum
    sCurSuffixNum = Me.txtSuffNum.Text
    sCurSuffixNum = FormatOMMapNum(sCurSuffixNum, "suffixnum")

    'Anomaly
    sCurAnomaly = Me.txtAnomaly.Text
    sCurAnomaly = FormatOMMapNum(sCurAnomaly, "anomaly")
    
    'Concatenate everything and compare to existing ORMAPMapNumber
    'ORMAPMapNumber
    sCurOMMapNum = sCurCounty & sCurTown & sCurTownPart & sCurTownDir & _
        sCurRange & sCurRangePart & sCurRangeDir & sCurSection & sCurQtr & _
        sCurQtrQtr & sCurAnomaly & sCurSuffixType & sCurSuffixNum
        
    Me.txtORMAPMapNum.Text = sCurOMMapNum


    m_pMIFeat.Value(m_lOMMapNumFld) = sCurOMMapNum
    
    Set pTLFeat = Nothing
    Set pFeatCur = Nothing
    m_pMIFeat.Store
    pWSEdit.StopEditOperation
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "cmdAssign_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmdHelp_Click()
    Dim sFilePath As String
    sFilePath = app.Path & "\" & "MapIndex_help.rtf"
    If FileExists(sFilePath) Then
'++ START JWM 10/16/2006 using new method to open help file
        gsb_StartDoc Me.hwnd, sFilePath
'++ START/END JWM 10/16/2006
    Else
        MsgBox "No help file available in current directory.", vbOKOnly + vbInformation
    End If
End Sub

Private Sub cmdQuit_Click()
  On Error GoTo ErrorHandler

    'Prompt for save if necessary
    
    Unload frmMapIndex

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

Private Function ffn_l_LocateFields(ByRef pFldName As String, ByRef pFClass As IFeatureClass) As Long
  On Error GoTo ErrorHandler
'++ This function appears to be reproduced in modutils with a different parameter list JWM 10/30/2006

    'Generic function to locate a field and warn user if it can't be found
    Dim lFld As Long
    lFld = pFClass.Fields.FindField(pFldName)
    If lFld > -1 Then
      ffn_l_LocateFields = lFld
    Else
        MsgBox "Unable to locate " & g_pFldnames.MICountyFN & " field in " & _
        g_pFldnames.FCMapIndex & " feature class"
        m_bContinue = False
    End If


  Exit Function
ErrorHandler:
  HandleError False, "FindField " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function


Private Sub txtAnomaly_Change()
    'This is now able to contain non numeric values
    'If Not IsNumeric(txtAnomaly.Text) Then txtAnomaly.Text = ""
  On Error GoTo ErrorHandler

    m_bPossiblyChanged = True

  Exit Sub
ErrorHandler:
  HandleError True, "txtAnomaly_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub

Private Sub txtMapNum_Change()
    m_bPossiblyChanged = True
End Sub

Private Sub txtSuffNum_Change()
    If Not IsNumeric(txtSuffNum.Text) Then txtSuffNum.Text = ""
    
    m_bPossiblyChanged = True
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
  Set m_pMxDoc = GetMXDocRef

    Dim bOpenForm As Boolean
    Me.Refresh
    Dim response As Variant
    Set m_pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If m_pMIFlayer Is Nothing Then
        MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
    End If
    Set m_pMIFclass = m_pMIFlayer.FeatureClass
    'Get the MapIndex feature layer and fclass
    Set m_pTaxlotFlayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If m_pTaxlotFlayer Is Nothing Then
        response = MsgBox("Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot & ".  " & _
        "Load " & g_pFldnames.FCMapIndex & " automatically?", vbYesNo)
        If response <> vbYes Then
            InitForm = False
            GoTo Process_Exit
        Else
            LoadFCIntoMap g_pFldnames.FCTaxlot, m_pMIFclass
            Set m_pTaxlotFlayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
        End If
    End If
    Set m_pTaxlotFClass = m_pTaxlotFlayer.FeatureClass

    'Get fields needed to populate the form
    Set m_pMIFields = m_pMIFclass.Fields
    m_bContinue = True
    m_lOMMapNumFld = LocateFields(m_pMIFclass, g_pFldnames.MIORMAPMapNumberFN)
    m_lReliabFld = LocateFields(m_pMIFclass, g_pFldnames.MIReliabFN)
    m_lScaleFld = LocateFields(m_pMIFclass, g_pFldnames.MIMapScaleFN)
    m_lMapNumFld = LocateFields(m_pMIFclass, g_pFldnames.MIMapNumberFN)
    m_lPageFld = LocateFields(m_pMIFclass, g_pFldnames.MIPageFN)
    If Not m_bContinue Then
        InitForm = False
        GoTo Process_Exit 'If any fields not found
    End If
    
    'Get the selected feature and its attributes
    Dim sExistOMMapNum As String
    Dim sExistVal As String
    Dim pFeatCur As IFeatureCursor
    Set pFeatCur = modUtils.GetSelectedFeatures(m_pMIFlayer)
    If pFeatCur Is Nothing Then
        InitForm = False
        GoTo Process_Exit
    End If
    Set m_pMIFeat = pFeatCur.NextFeature
    
    'Get a Taxlot feature, so its domains can be referenced
    Dim pTLFeatCur As IFeatureCursor
    Set pTLFeatCur = m_pTaxlotFlayer.Search(Nothing, True)
    Dim pTLFeat As IFeature
    Set pTLFeat = pTLFeatCur.NextFeature
    
    'Populate the form with domain values
    'ORMAPMapNumber
    sExistOMMapNum = ReadValue(m_pMIFeat, g_pFldnames.MIORMAPMapNumberFN)
    'Verify that the number is the right length.  If not, load default values
    'into the fields below
    If Not Len(sExistOMMapNum) = ORMAP_MAPNUM_FIELD_LENGTH Then
        Me.txtORMAPMapNum.Text = ""
        m_bSuccess = AddCodesToCmb(g_pFldnames.MIReliabFN, m_pMIFields, Me.cmbReliability, "", True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.MIMapScaleFN, m_pMIFields, Me.cmbScale, "", True)
        Me.txtPage.Text = ""
        'Convert default county to description
        Dim sDefCntyDesc As String
        sDefCntyDesc = modUtils.ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLCountyFN, CLng(g_pFldnames.DefCounty))
        
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLCountyFN, m_pTaxlotFClass.Fields, Me.cmbCounty, sDefCntyDesc, True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownFN, m_pTaxlotFClass.Fields, Me.cmbTown, "", True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownPartFN, m_pTaxlotFClass.Fields, Me.cmbTownPart, modUtils.ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLTownPartFN, CDbl(g_pFldnames.DefTownPart)), True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownDirFN, m_pTaxlotFClass.Fields, Me.cmbTownDir, g_pFldnames.DefTownDir, True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeFN, m_pTaxlotFClass.Fields, Me.cmbRange, "", True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangePartFN, m_pTaxlotFClass.Fields, Me.cmbRangePart, modUtils.ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLRangePartFN, CDbl(g_pFldnames.DefRangePart)), True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeDirFN, m_pTaxlotFClass.Fields, Me.cmbRangeDir, g_pFldnames.DefRangeDir, True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLSectNumberFN, m_pTaxlotFClass.Fields, Me.cmbSection, "", True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtr, g_pFldnames.DefQtr, True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtrQtr, g_pFldnames.DefQtrQtr, True)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLSufTypeFN, m_pTaxlotFClass.Fields, Me.cmbSufftype, modUtils.ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.MIMapSuffNumFN, g_pFldnames.DefSuffType))
        
        Me.txtSuffNum.Text = g_pFldnames.DefSuffNum
        Me.txtAnomaly.Text = g_pFldnames.DefAnomaly
        'txtAnomaly.Text = ""
    Else
        Me.txtORMAPMapNum.Text = sExistOMMapNum
        'm_bSuccess = AddCodesToCmb(g_pFldnames.MICountyFN, m_pMIFields, Me.cmbCounty, sExistVal)
        sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIMapNumberFN)
        Me.txtMapNum.Text = sExistVal
        'Reliability
        sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIReliabFN)
        m_bSuccess = AddCodesToCmb(g_pFldnames.MIReliabFN, m_pMIFields, Me.cmbReliability, sExistVal, True)
        'Scale
        sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIMapScaleFN)
        m_bSuccess = AddCodesToCmb(g_pFldnames.MIMapScaleFN, m_pMIFields, Me.cmbScale, sExistVal, True)
        'Page
        sExistVal = ReadValue(m_pMIFeat, g_pFldnames.MIPageFN)
        Me.txtPage.Text = sExistVal
        'County
        sExistVal = ParseOMMapNum(sExistOMMapNum, "county")
        sExistVal = ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLCountyFN, Int(sExistVal))
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLCountyFN, m_pTaxlotFClass.Fields, Me.cmbCounty, sExistVal, True)
        'Town
        sExistVal = ParseOMMapNum(sExistOMMapNum, "town")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLTownFN)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownFN, m_pTaxlotFClass.Fields, Me.cmbTown, sExistVal, True)
        'TownPart
        sExistVal = ParseOMMapNum(sExistOMMapNum, "townpart")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLTownPartFN)
        'If Len(sExistVal) = 3 Then sExistVal = Left(sExistVal, 1) & "." & Right(sExistVal, 2)
        If Len(sExistVal) = 3 And Left(sExistVal, 1) = "." Then sExistVal = "0" & sExistVal
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownPartFN, m_pTaxlotFClass.Fields, Me.cmbTownPart, sExistVal, True)
        'TownDir
        sExistVal = ParseOMMapNum(sExistOMMapNum, "towndir")
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLTownDirFN, m_pTaxlotFClass.Fields, Me.cmbTownDir, sExistVal, True)
        'Range
        sExistVal = ParseOMMapNum(sExistOMMapNum, "range")
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeFN, m_pTaxlotFClass.Fields, Me.cmbRange, sExistVal, True)
        'RangePart
        sExistVal = ParseOMMapNum(sExistOMMapNum, "rangepart")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLRangePartFN)
        'If Len(sExistVal) = 3 Then sExistVal = Left(sExistVal, 1) & "." & Right(sExistVal, 2)
        If Len(sExistVal) = 3 And Left(sExistVal, 1) = "." Then sExistVal = "0" & sExistVal
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangePartFN, m_pTaxlotFClass.Fields, Me.cmbRangePart, sExistVal, True)
        'RangeDir
        sExistVal = ParseOMMapNum(sExistOMMapNum, "rangedir")
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLRangeDirFN, m_pTaxlotFClass.Fields, Me.cmbRangeDir, sExistVal, True)
        'Section
        sExistVal = ParseOMMapNum(sExistOMMapNum, "section")
        'sExistVal = ReadValue(pTLFeat, g_pFldnames.TLSectNumberFN)
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLSectNumberFN, m_pTaxlotFClass.Fields, Me.cmbSection, sExistVal, True)
        'Qtr
        sExistVal = ParseOMMapNum(sExistOMMapNum, "qtr")
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtr, sExistVal, True)
        'QtrQtr
        sExistVal = ParseOMMapNum(sExistOMMapNum, "qtrqtr")
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLQtrQtrFN, m_pTaxlotFClass.Fields, Me.cmbQtrQtr, sExistVal, True)
        'MapSuffixType
        sExistVal = ConvertToDescription(m_pTaxlotFClass.Fields, g_pFldnames.TLSufTypeFN, ParseOMMapNum(sExistOMMapNum, "suffixtype"))
        m_bSuccess = AddCodesToCmb(g_pFldnames.TLSufTypeFN, m_pTaxlotFClass.Fields, Me.cmbSufftype, sExistVal, True)
        'MapSuffixNum
        sExistVal = ParseOMMapNum(sExistOMMapNum, "suffixnum")
        Me.txtSuffNum.Text = sExistVal
        'Anomaly
        sExistVal = ParseOMMapNum(sExistOMMapNum, "anomaly")
        txtAnomaly.Text = sExistVal
    End If
    m_bPossiblyChanged = False
    InitForm = True
Process_Exit:
Exit Function

End Function
