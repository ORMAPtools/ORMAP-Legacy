VERSION 5.00
Begin VB.Form frmMapIndex 
   AutoRedraw      =   -1  'True
   Caption         =   "Map Index"
   ClientHeight    =   4965
   ClientLeft      =   1545
   ClientTop       =   4605
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Map Information"
      ClipControls    =   0   'False
      Height          =   3525
      Left            =   3690
      TabIndex        =   13
      Top             =   60
      Width           =   3855
      Begin VB.TextBox txtAnomaly 
         Height          =   285
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   41
         Top             =   3120
         Width           =   615
      End
      Begin VB.ComboBox cmbCounty 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   2025
      End
      Begin VB.ComboBox cmbReliability 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1980
         Width           =   2025
      End
      Begin VB.ComboBox cmbScale 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2370
         Width           =   2025
      End
      Begin VB.TextBox txtMapNum 
         Height          =   285
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   15
         Top             =   735
         Width           =   2025
      End
      Begin VB.TextBox txtPage 
         Height          =   285
         Left            =   1620
         TabIndex        =   20
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtSuffNum 
         Height          =   285
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   16
         Top             =   1095
         Width           =   2025
      End
      Begin VB.ComboBox cmbSufftype 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lblAnomaly 
         Caption         =   "Anomaly:"
         Height          =   225
         Left            =   900
         TabIndex        =   40
         Top             =   3150
         Width           =   675
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   3700
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   3695
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label lblCounty 
         Alignment       =   1  'Right Justify
         Caption         =   "County:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   960
         TabIndex        =   39
         Top             =   420
         Width           =   615
      End
      Begin VB.Label lblSufType 
         Alignment       =   1  'Right Justify
         Caption         =   "Map Suffix Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   38
         Top             =   1500
         Width           =   1395
      End
      Begin VB.Label lblReliability 
         Alignment       =   1  'Right Justify
         Caption         =   "Reliability:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   810
         TabIndex        =   37
         Top             =   2040
         Width           =   765
      End
      Begin VB.Label lblScale 
         Alignment       =   1  'Right Justify
         Caption         =   "Scale:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1080
         TabIndex        =   36
         Top             =   2430
         Width           =   495
      End
      Begin VB.Label lblSufNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Map Suffix Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   35
         Top             =   1155
         Width           =   1455
      End
      Begin VB.Label lblMapNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Map Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   570
         TabIndex        =   34
         Top             =   765
         Width           =   1005
      End
      Begin VB.Label lblPage 
         Alignment       =   1  'Right Justify
         Caption         =   "Page:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1110
         TabIndex        =   33
         Top             =   2790
         Width           =   465
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ORMAP Number"
      ClipControls    =   0   'False
      Height          =   735
      Left            =   3690
      TabIndex        =   21
      Top             =   3690
      Width           =   3855
      Begin VB.Label lblORMAPNumber 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   42
         Top             =   300
         Width           =   3555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Section"
      ClipControls    =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   3300
      Width           =   3405
      Begin VB.ComboBox cmbSection 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbQtr 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   735
         Width           =   855
      End
      Begin VB.ComboBox cmbQtrQtr 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label lblSection 
         Alignment       =   1  'Right Justify
         Caption         =   "Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   960
         TabIndex        =   32
         Top             =   390
         Width           =   645
      End
      Begin VB.Label lblQtr 
         Alignment       =   1  'Right Justify
         Caption         =   "Quarter:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   930
         TabIndex        =   31
         Top             =   780
         Width           =   675
      End
      Begin VB.Label lblQtrQtr 
         Alignment       =   1  'Right Justify
         Caption         =   "Quarter of Quarter:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   30
         Top             =   1170
         Width           =   1425
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Range"
      ClipControls    =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3405
      Begin VB.ComboBox cmbRange 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbRangePart 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1110
         Width           =   1575
      End
      Begin VB.ComboBox cmbRangeDir 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   900
         TabIndex        =   29
         Top             =   390
         Width           =   675
      End
      Begin VB.Label lblRangePart 
         Alignment       =   1  'Right Justify
         Caption         =   "Partial Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   28
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label lblRangeDir 
         Alignment       =   1  'Right Justify
         Caption         =   "Directional:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   750
         TabIndex        =   27
         Top             =   780
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Township"
      ClipControls    =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3405
      Begin VB.ComboBox cmbTown 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbTownPart 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1110
         Width           =   1575
      End
      Begin VB.ComboBox cmbTownDir 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label lblTown 
         Alignment       =   1  'Right Justify
         Caption         =   "Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   750
         TabIndex        =   26
         Top             =   390
         Width           =   795
      End
      Begin VB.Label lblTownPart 
         Alignment       =   1  'Right Justify
         Caption         =   "Partial Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   510
         TabIndex        =   25
         Top             =   1170
         Width           =   1035
      End
      Begin VB.Label lblTownDir 
         Alignment       =   1  'Right Justify
         Caption         =   "Directional:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   660
         TabIndex        =   24
         Top             =   780
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6810
      TabIndex        =   23
      Top             =   4530
      Width           =   800
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3690
      TabIndex        =   22
      Top             =   4530
      Width           =   800
   End
   Begin VB.CommandButton cmdEditSave 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5970
      TabIndex        =   4
      Top             =   4530
      Width           =   800
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
' Initial Author:       <<Unknown>>
'
' Date Created:         <<Unknown>>
'
' Description:
'       Form used to capture and validate attributes for MapIndex features. The attributes,
'       once validated are used to construct the ORMAP Number for the feature.
'
' Entry points:
'       Form Object
'       Methods
'           InitForm
'               Initialize the form to the current feature selection
'           Frame
'               The Frame object that ArcGIS uses to display the form

' Dependencies:
'       File References:
'           esriArcMapUI
'           esriCarto
'           esriGeoDatabase
'           esriSystem
'       File Dependencies
'           ORMAPNumber
'
' Issues:
'       None known at this time (2/6/2007 JWalton)
'
' Method:
'       This form is implemented as a modeless form that can sit on top of ArcMap
'       while allowing the user continued access to ArcMap.
'       This implementation is made possible through the Frame property of the
'       form that implements ESRI's Modeless Window interface.
'
' Updates:
'       10/11/2006 -- Added comment headers (JWM)
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)



Option Explicit
'******************************
' Private Definitions
'------------------------------
'------------------------------
' Private Variables
'------------------------------
'++ START JWalton 1/31/2007 Changed variable scope from Public to Private
Private m_pMxDoc As esriArcMapUI.IMxDocument
Private m_pMIFlayer As esriCarto.IFeatureLayer2
Private m_pTaxlotFLayer As esriCarto.IFeatureLayer2
'++ START JWalton 1/31/2007 Additional variable declarations
Private m_pApp As esriFramework.IApplication
'++ END JWalton 1/31/2007
Private m_pFrame As esriFramework.IModelessFrame
Private m_pMIFeat As esriGeoDatabase.IFeature
Private m_pMIFclass As esriGeoDatabase.IFeatureClass
Private m_pTaxlotFClass As esriGeoDatabase.IFeatureClass
Private m_pMIFields As esriGeoDatabase.IFields2
'++ START JWalton 1/31/2007 Additional variable declarations
Private m_pMapIndexFields As MapIndexFieldMap
Private WithEvents m_pORMAPNumber As ORMAPNumber
Attribute m_pORMAPNumber.VB_VarHelpID = -1
Private m_pTaxlotFields As TaxlotFieldMap
'++ END JWalton 1/31/2007
Private m_bSuccess As Boolean
Private m_bPossiblyChanged As Boolean
'++ START JWalton 1/31/2007 Additional variable declarations
Private m_blnEditState As Boolean
'++ END JWalton 1/31/2007
Private m_lMapNumFld As Long
Private m_lOMMapNumFld As Long
Private m_lPageFld As Long
Private m_lReliabFld As Long
Private m_lScaleFld As Long
Private m_ParentHWND As Long
'++ END JWalton 1/31/2007

'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "frmMapIndex.frm"

'------------------------------
' Private Types
'------------------------------
Private Type MapIndexFieldMap
    MapNumber As Long
    MapScale As Long
    ORMAPNumber As Long
    Page As Long
    Reliability As Long
End Type

Private Type TaxlotFieldMap
    Anomaly As Long
    County As Long
    MapAcres As Long
    MapNumber As Long
    MapTaxlotNumber As Long
    MapTaxlot As Long
    OrmapTaxlotNumber As Long
    OrmapMapNumber As Long
    PartialRangeCode As Long
    PartialTownshipCode As Long
    Quarter As Long
    QuarterQuarter As Long
    Range As Long
    RangeDirectional As Long
    Section As Long
    SpecialInterest As Long
    SuffixNumber As Long
    SuffixType As Long
    Taxlot As Long
    Township As Long
    TownshipDirectional As Long
End Type


'----------------------------------------------------------------------------
'Name:                  cmdAssign_Click                                     '
'Initial Author:        <<Unknown>>                                         '
'Subsequent Author:     JWalton                                             '
'Created:               <<Unknown>>                                         '
'Purpose:       Assign attributes to the current MapIndex polygon           '
'Description:   Verify that all fields have values entered.  Don't allow    '
'               changes tobe applied without all values present             '
'Methods:       Describe any complex details.                               '
'Inputs:        <<None>>                                                    '
'Parameters:    <<None>>                                                    '
'Outputs:       m_pMIFeat                                                   '
'Returns:       <<N/A>>                                                     '
'Errors:        This routine raises no known errors.                        '
'Assumptions:   m_pMIFeat Is Not Nothing                                    '
'               m_pORMAPNumber Is Not Nothing                               '
'Updates:                                                                   '
'       Type any updates here.                                              '
'Developer:     Date:       Comments:                                       '
'----------     ------      ---------                                       '
'James Moore    10/11/2006  Initial creation of this comment section        '
'JWalton        2/1/2006    Rewritten due to ORMAP Class Object             '
'----------------------------------------------------------------------------

Private Sub cmdEditSave_Click()
On Error GoTo ErrorHandler
'++ START JWalton 1/31/2007 Centralized variable declarations
    ' Variable declarations
    Dim pWSEdit As esriGeoDatabase.IWorkspaceEdit
    Dim pDataset As esriGeoDatabase.IDataset
    Dim bValuesPresent As Boolean
    Dim sValue As String
    
    If m_blnEditState Then
        ' Validate controls
        bValuesPresent = True
        bValuesPresent = bValuesPresent And (Len(cmbReliability.Text) <> 0)
        bValuesPresent = bValuesPresent And (Len(cmbScale.Text) <> 0)
        bValuesPresent = bValuesPresent And (Len(txtMapNum.Text) <> 0)
        bValuesPresent = bValuesPresent And (Len(txtPage.Text) <> 0)
        bValuesPresent = bValuesPresent And m_pORMAPNumber.IsValidNumber()
        
        ' Deals with the case of errors
        If Not bValuesPresent Then
           MsgBox "All fields must be filled in before assigning", vbOKOnly
           GoTo Process_Exit
        End If
        
        ' Starts an edit operation
        Set pDataset = m_pMIFclass
        Set pWSEdit = pDataset.Workspace
        pWSEdit.StartEditOperation
        
        ' Update the form caption
        Me.Caption = "Map Index (Map Feature: " & m_pORMAPNumber.ORMAPNumber & ")"

        'MapNumber
        m_pMIFeat.Value(m_pMapIndexFields.MapNumber) = Me.txtMapNum.Text
    
        'Reliability
        sValue = basUtilities.ConvertCode(m_pMIFeat.Fields, g_pFldnames.MIReliabFN, Me.cmbReliability)
        If IsNumeric(sValue) Then
            m_pMIFeat.Value(m_pMapIndexFields.Reliability) = CInt(sValue)
          Else
            m_pMIFeat.Value(m_pMapIndexFields.Reliability) = Null
        End If
        
        'Scale
        sValue = basUtilities.ConvertCode(m_pMIFeat.Fields, g_pFldnames.MIMapScaleFN, Me.cmbScale.Text)
        If IsNumeric(sValue) Then
            m_pMIFeat.Value(m_pMapIndexFields.MapScale) = CLng(sValue)
          Else
            m_pMIFeat.Value(m_pMapIndexFields.MapScale) = Null
        End If
    
        'Page
        sValue = Me.txtPage.Text
        If IsNumeric(sValue) Then
            m_pMIFeat.Value(m_pMapIndexFields.Page) = CLng(sValue)
        Else
            m_pMIFeat.Value(m_pMapIndexFields.Page) = Null
        End If

        ' ORMAP Number
        m_pMIFeat.Value(m_pMapIndexFields.ORMAPNumber) = m_pORMAPNumber.ORMAPNumber
        
        ' Store the edited feature
        m_pMIFeat.Store
        
        ' Update all taxlot polygons that underlie this polygon
        UpdateTaxlots m_pMIFeat, m_pTaxlotFClass
        
        ' Finalize the update operation
        pWSEdit.StopEditOperation
    End If
    
    '++ START JWalton 1/31/2007
    ' Toggles form options after update
    m_blnEditState = Not m_blnEditState
    ToggleControls m_blnEditState
    
    ' Update the form caption
    Me.Frame.Caption = "Map Index (" & m_pORMAPNumber.ORMAPNumber & ")"
    '++ END JWalton 1/31/2007

Process_Exit:
  Exit Sub
  
ErrorHandler:
    MsgBox "Error # " & Err.Number & " (" & Err.Description & ")"
    ' Cancels the current edit operation
    pWSEdit.AbortEditOperation

    ' Handle the error
    HandleError True, "cmdAssign_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

'***************************************************************************
'Name:                  cmdHelp
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:           Help for Locate Tool
'Called From:       cmdHelp
'Description:       Event handler for cmdHelp command button control
'Methods:           None
'Inputs:            None
'Parameters:        None
'Outputs:           None
'Returns:           Nothing
'Errors:            This routine raises no known errors.
'Assumptions:       None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007        Initial creation
'***************************************************************************

Private Sub cmdHelp_Click()
    '++ START JWalton 2/6/2007 Centralized Variable Declarations
    ' Variable declartions
    Dim sFilePath As String
    '++ END JWalton 2/6/2007    Dim sFilePath As String
    
    sFilePath = app.Path & "\" & "MapIndex_help.rtf"
    If FileExists(sFilePath) Then
'++ START JWM 10/16/2006 using new method to open help file
        basUtilities.gsb_StartDoc Me.hwnd, sFilePath
'++ START/END JWM 10/16/2006
    Else
        MsgBox "No help file available in current directory.", vbOKOnly + vbInformation
    End If
End Sub

'***************************************************************************
'Name:                  cmdQuit
'Initial Author:        <<Unknown>>
'Subsequent Author:     JWalton
'Created:               <<Unknown>>
'Purpose:       Quit the form or cancel edits, depending on the edit mode
'Called From:   cmdQuit Control
'Description:   Either cancels current edits and reinitializes the form, or
'               closes the form based upon whether or not the user is
'               currently editing an entry
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Rewritten from Original
'***************************************************************************


Private Sub cmdQuit_Click()
On Error GoTo ErrorHandler

    '++ START JWalton 1/31/2007
    ' Reset the form or exits depending on the editing status
    If m_blnEditState Then
        Me.InitForm
        m_blnEditState = False
        ToggleControls m_blnEditState
      Else
        Unload Me
    End If
    '++ END JWalton 1/31/2007
    
    Exit Sub
ErrorHandler:
    HandleError True, _
                "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4, _
                m_ParentHWND
End Sub

'***************************************************************************
'Name:                  Form_Initialize
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Initialize the form position
'Called From:   Form Object
'Description:   Event Handler for the form initialize event
'Methods:       The window position is not, and cannot be, set through the
'               form left and top properties.
'               The window position is set using the IWindowPosition
'               interface of the Frame object exposed by ESRI.
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:       None
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub Form_Initialize()
    ' Variable declarations
    Dim pWindowPos As esriFramework.IWindowPosition
    
    ' Initialize objects
    Set pWindowPos = Me.Frame
    
    ' Get the initial position
    pWindowPos.Top = GetSetting("ArcGIS.ArcMap.ORMAP.Tools", "FrmMapIndex", "Top", 0)
    pWindowPos.Left = GetSetting("ArcGIS.ArcMap.ORMAP.Tools", "FrmMapIndex", "Left", 0)
    
    ' Makes the form the topmost form
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

'***************************************************************************
'Name:                  Form_Load
'Initial Author:        John Walton
'Subsequent Author:     <<Type your name here>>
'Created:               1/31/2007
'Purpose:       Initialize the form and all controls on the form, and
'               register the form
'Called From:   Form Object
'Description:   Event handler for Form Load event
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pMxDoc
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub Form_Load()
On Error GoTo ErrorHandler
    ' Variable declarations
    Dim iResponse As Variant
    
    ' Sets the form status to open
    g_pForms.SetFormStatus Me.Name, True

    'Get a reference to the MXDocument
    Set m_pMxDoc = g_pApp.Document

    ' Initialize the Map Index feature layer and feature class
    Set m_pMIFlayer = basUtilities.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If m_pMIFlayer Is Nothing Then
        MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex & _
        "Load " & g_pFldnames.FCMapIndex & "?"
        If iResponse <> vbYes Then
            Set g_pApp.CurrentTool = Nothing
            Exit Sub
        Else
            If basUtilities.LoadFCIntoMap(g_pFldnames.FCMapIndex) Then
                Set m_pMIFlayer = basUtilities.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
              Else
                Set g_pApp.CurrentTool = Nothing
                Exit Sub
            End If
        End If
    End If
    Set m_pMIFclass = m_pMIFlayer.FeatureClass
    
    ' Initialize the Taxlot feature layer and feature class
    Set m_pTaxlotFLayer = basUtilities.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If m_pTaxlotFLayer Is Nothing Then
        iResponse = MsgBox("Unable to locate Taxlot layer in Table of Contents.  " & _
                "This process requires a feature class called " & g_pFldnames.FCTaxlot & ".  " & _
                "Load " & g_pFldnames.FCMapIndex & "?", vbYesNo)
        If iResponse <> vbYes Then
            Set g_pApp.CurrentTool = Nothing
            Exit Sub
        Else
            If basUtilities.LoadFCIntoMap(g_pFldnames.FCTaxlot) Then
                Set m_pTaxlotFLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
              Else
                Set g_pApp.CurrentTool = Nothing
                Exit Sub
            End If
        End If
    End If
    Set m_pTaxlotFClass = m_pTaxlotFLayer.FeatureClass

    'Initialize the field positions of fields to be manipulated
    If Not InitializeFieldPositions(m_pMIFclass.Fields, m_pTaxlotFClass.Fields) Then
        MsgBox "Required fields are missing from the feature class MapIndex", vbOKOnly Or vbCritical
        Set g_pApp.CurrentTool = Nothing
        Exit Sub
    End If
    
    ' Initialize the lock state of controls on the form
    ToggleControls False
    
    ' Initialize the editing status of the form
    m_blnEditState = False

Process_Exit:
    Exit Sub

ErrorHandler:
    HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    GoTo Process_Exit
End Sub

'***************************************************************************
'Name:                  Form_Unload
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Save the current window position
'Called From:   Form Object
'Description:   Saves the current position of the frame that ESRI ArcMap
'               uses to draw the form
'Methods:       The window position does not, and cannot, come from the
'               form, since the form is modeless.
'               The window position must be obtained through the IWindowPos
'               interface of the Frame object.
'Inputs:        None
'Parameters:    None
'Outputs:       Registry Settings
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub Form_Unload( _
  Cancel As Integer)
    ' Variable declarations
    Dim pWindowPos As esriFramework.IWindowPosition
    
    ' Initialize objects
    Set pWindowPos = Me.Frame
    
    ' Saves the position of the form
    SaveSetting "ArcGIS.ArcMap.ORMAP.Tools", "FrmMapIndex", "Top", pWindowPos.Top
    SaveSetting "ArcGIS.ArcMap.ORMAP.Tools", "FrmMapIndex", "Left", pWindowPos.Left
End Sub

'***************************************************************************
'Name:                  Form_QueryUnload
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Unregister the form and save the form's position
'Called From:   Form Object
'Description:   Event handler for Form QueryUnload event
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub Form_QueryUnload( _
  Cancel As Integer, _
  UnloadMode As Integer)
    ' Sets the form status to not open
    g_pForms.SetFormStatus Me.Name, False
End Sub

'***************************************************************************
'Name:                  InitForm
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Called From:   Multiple Location
'Description:   Coordinate the initialization of the form for a new
'               selection.
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pMIFeat, m_pORMAPNumber
'Returns:       A boolean value that indicates the success of initializing
'               the form to the selected Map Index polygon.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    12-19-06    Implemented implied logic when getting field
'                           indexes
'JWalton        2/1/2007    Implemented initialization procedures for both
'                           empty and initialized feature selections.
'               2/1/2007    Implemented field position initialization in the
'                           Form_Load event.
'***************************************************************************

Public Function InitForm() As Boolean
    ' Variable declarations
    Dim pFeatCur As esriGeoDatabase.IFeatureCursor
    
    'Get the selected feature and its attributes
    Set pFeatCur = basUtilities.GetSelectedFeatures(m_pMIFlayer)
    If pFeatCur Is Nothing Then
        InitForm = False
        GoTo Process_Exit
    End If
    Set m_pMIFeat = pFeatCur.NextFeature
    
    ' Obtains and parses the ORMAP Map number
    Set m_pORMAPNumber = New ORMAPNumber
    m_pORMAPNumber.ParseNumber basUtilities.ReadValue(m_pMIFeat, g_pFldnames.MIORMAPMapNumberFN)
    
    ' Verify that the ORMAP Number, and load defaults if invalid
    If Not m_pORMAPNumber.IsValidNumber() Then
        InitForm = InitEmpty(m_pMIFclass.Fields, m_pTaxlotFClass.Fields)
        ToggleControls True
        m_blnEditState = True
    Else
        InitForm = InitWithFeature(m_pMIFeat, m_pMIFclass.Fields, m_pTaxlotFClass.Fields)
    End If
    
    ' Sets the current ORMAP Number in the caption of the form and the ORMAP textbox
    Frame.Caption = "Map Index (" & m_pORMAPNumber.ORMAPNumber & ")"
    lblORMAPNumber.Caption = m_pORMAPNumber.ORMAPNumber
        
    ' Refresh the form
    Me.Refresh
    
Process_Exit:
    Exit Function

Err_Handler:
    InitForm = False
End Function

'***************************************************************************
'Name:                  ToggleControls
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Lock or unlock controls on the form
'Called From:   Multiple Locations
'Description:   Sets the enabled state of all editing controls on the form
'               according the state passed into the procedure
'Methods:       None
'Inputs:        A boolean value representing the desired enabled state of
'               editing controls.
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

'++ START JWalton 1/31/2007
Private Sub ToggleControls( _
  ByVal blnState As Boolean)
    ' Variable declarations
    Dim ctl As Control
    
    ' Loops through controls disabling all input controls
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Or _
           TypeOf ctl Is ComboBox Then
            ctl.Enabled = blnState
        End If
    Next ctl
    
    ' Handles the command buttons
    If blnState Then
        cmdEditSave.Caption = "&Save"
        cmdQuit.Caption = "&Cancel"
      Else
        cmdEditSave.Caption = "&Edit"
        cmdQuit.Caption = "&Quit"
    End If
    
    ' Clean up and exit
    Set ctl = Nothing
End Sub

'***************************************************************************
'Name:                  InitEmpty
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Initialize the form for a new selection
'Called From:   InitForm
'Description:   Clear the form and set default value for fields in
'               accordance with a selection that has no values set
'Methods:       None
'Inputs:        pMapIndexFlds, pTaxlotFlds
'Parameters:    None
'Outputs:       None
'Returns:       A boolean value that indicates the success of initializing
'               the form
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************


Private Function InitEmpty( _
  ByRef pMapIndexFlds As esriGeoDatabase.IFields, _
  pTaxlotFlds As esriGeoDatabase.IFields) As Boolean
On Error GoTo Err_Handler
    ' Clears the ORMAP Number object
    Set m_pORMAPNumber = New ORMAPNumber

    ' Clear All Lists and Values
    ResetControls

    ' Initialize the new ORMAP Number
    With m_pORMAPNumber
        .County = CLng(g_pFldnames.DefCounty)
        .Township = ""
        .TownshipDirectional = g_pFldnames.DefTownDir
        .PartialTownshipCode = CDbl(g_pFldnames.DefTownPart)
        .Range = ""
        .RangeDirectional = g_pFldnames.DefRangeDir
        .PartialRangeCode = CDbl(g_pFldnames.DefRangePart)
        .Section = ""
        .Quarter = g_pFldnames.DefQtr
        .QuarterQuarter = g_pFldnames.DefQtrQtr
        .SuffixNumber = g_pFldnames.DefSuffNum
        .SuffixType = g_pFldnames.DefSuffType
        .Anomaly = g_pFldnames.DefAnomaly
    End With
    
    ' Sets the caption of the form
    Me.Caption = "Map Index (Map Feature: <Not Attributed>)"
        
    ' Reliability
    basUtilities.AddCodesToCmb g_pFldnames.MIReliabFN, _
                               pMapIndexFlds, _
                               Me.cmbReliability, _
                               "", _
                               True
    
    ' Scale
    basUtilities.AddCodesToCmb g_pFldnames.MIMapScaleFN, _
                               pMapIndexFlds, _
                               Me.cmbScale, _
                               "", _
                               True
    
    ' Counties
    basUtilities.AddCodesToCmb g_pFldnames.TLCountyFN, _
                               pTaxlotFlds, _
                               Me.cmbCounty, _
                               basUtilities.ConvertToDescription(pTaxlotFlds, _
                                                                 g_pFldnames.TLCountyFN, _
                                                                 CLng(m_pORMAPNumber.County)), _
                               True
                  
    ' Townships
    basUtilities.AddCodesToCmb g_pFldnames.TLTownFN, _
                               pTaxlotFlds, _
                               Me.cmbTown, _
                               m_pORMAPNumber.Township, _
                               True
                  
    ' Partial Township Codes
    basUtilities.AddCodesToCmb g_pFldnames.TLTownPartFN, _
                               pTaxlotFlds, _
                               Me.cmbTownPart, _
                               basUtilities.ConvertToDescription(pTaxlotFlds, _
                                                                 g_pFldnames.TLTownPartFN, _
                                                                 CDbl(m_pORMAPNumber.PartialTownshipCode)), _
                               True
                  
    ' Township Directionals
    basUtilities.AddCodesToCmb g_pFldnames.TLTownDirFN, _
                               pTaxlotFlds, _
                               Me.cmbTownDir, _
                               m_pORMAPNumber.TownshipDirectional, _
                               True
                  
    ' Ranges
    basUtilities.AddCodesToCmb g_pFldnames.TLRangeFN, _
                               pTaxlotFlds, _
                               Me.cmbRange, _
                               m_pORMAPNumber.Range, _
                               True
                               
    ' Partial Range Codes
    basUtilities.AddCodesToCmb g_pFldnames.TLRangePartFN, _
                               pTaxlotFlds, _
                               Me.cmbRangePart, _
                               basUtilities.ConvertToDescription(pTaxlotFlds, _
                                                                 g_pFldnames.TLRangePartFN, _
                                                                 CDbl(m_pORMAPNumber.PartialRangeCode)), _
                               True
                  
    ' Range Directionals
    basUtilities.AddCodesToCmb g_pFldnames.TLRangeDirFN, _
                               pTaxlotFlds, _
                               Me.cmbRangeDir, _
                               m_pORMAPNumber.RangeDirectional, _
                               True
                  
    ' Sections
    basUtilities.AddCodesToCmb g_pFldnames.TLSectNumberFN, _
                               pTaxlotFlds, _
                               Me.cmbSection, _
                               m_pORMAPNumber.Section, _
                               True
                  
    ' Quarter
    basUtilities.AddCodesToCmb g_pFldnames.TLQtrFN, _
                               pTaxlotFlds, _
                               Me.cmbQtr, _
                               m_pORMAPNumber.Quarter, _
                               True
                  
    ' QuarterQuarter
    basUtilities.AddCodesToCmb g_pFldnames.TLQtrQtrFN, _
                               pTaxlotFlds, _
                               Me.cmbQtrQtr, _
                               m_pORMAPNumber.QuarterQuarter, _
                               True
                  
    ' Suffix Type
    basUtilities.AddCodesToCmb g_pFldnames.TLSufTypeFN, _
                               pTaxlotFlds, _
                               Me.cmbSufftype, _
                               m_pORMAPNumber.SuffixType
    
    ' Anomaly, Page, and Suffix Number
    Me.txtAnomaly.Text = m_pORMAPNumber.Anomaly
    Me.txtPage.Text = "0"
    Me.txtSuffNum.Text = m_pORMAPNumber.SuffixNumber
    
    ' Returns the function's value
    InitEmpty = True
    Exit Function
    
Err_Handler:
    ' Return's the function's value
    InitEmpty = False
End Function

'***************************************************************************
'Name:                  InitWithFeature
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Initialize the form with a valid selection
'Called From:   InitForm
'Description:   Clear the form and sets values for fields in accordance with
'               a selection that has valid a valid ORMAP Number
'Methods:       None
'Inputs:        pFeature, pMapIndexFlds, pTaxlotFlds
'Parameters:    None
'Outputs:       None
'Returns:       A boolean value that indicates the success of initializing
'               the form
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Function InitWithFeature( _
  ByVal pFeature As esriGeoDatabase.IFeature, _
  ByVal pMapIndexFlds As esriGeoDatabase.IFields, _
  ByVal pTaxlotFlds As esriGeoDatabase.IFields) As Boolean
On Error GoTo Err_Handler
    ' Resets all controls
    ResetControls

    ' Sets the caption of the form
    Me.Caption = "Map Index (Map Feature: " & m_pORMAPNumber.ORMAPNumber & ")"
        
    ' Map Number
    Me.txtMapNum.Text = basUtilities.ReadValue(pFeature, g_pFldnames.MIMapNumberFN)
    
    ' Reliability
    basUtilities.AddCodesToCmb g_pFldnames.MIReliabFN, _
                               pMapIndexFlds, _
                               Me.cmbReliability, _
                               basUtilities.ReadValue(pFeature, g_pFldnames.MIReliabFN), _
                               True
    
    ' Scale
    basUtilities.AddCodesToCmb g_pFldnames.MIMapScaleFN, _
                               pMapIndexFlds, _
                               Me.cmbScale, _
                               basUtilities.ReadValue(pFeature, g_pFldnames.MIMapScaleFN), _
                               True
    
    ' Page
    Me.txtPage.Text = basUtilities.ReadValue(pFeature, g_pFldnames.MIPageFN)
    
    ' County
    basUtilities.AddCodesToCmb g_pFldnames.TLCountyFN, _
                               m_pTaxlotFClass.Fields, _
                               Me.cmbCounty, _
                               basUtilities.ConvertToDescription(pTaxlotFlds, _
                                                                 g_pFldnames.TLCountyFN, _
                                                                 Int(m_pORMAPNumber.County)), _
                               True
    
    ' Township
    basUtilities.AddCodesToCmb g_pFldnames.TLTownFN, _
                               pTaxlotFlds, _
                               Me.cmbTown, _
                               m_pORMAPNumber.Township, _
                               True
    
    ' Partial Township Code
    basUtilities.AddCodesToCmb g_pFldnames.TLTownPartFN, _
                               pTaxlotFlds, _
                               Me.cmbTownPart, _
                               "0" & m_pORMAPNumber.PartialTownshipCode, _
                               True
    
    ' Township Directional
    basUtilities.AddCodesToCmb g_pFldnames.TLTownDirFN, _
                              pTaxlotFlds, _
                              Me.cmbTownDir, _
                              m_pORMAPNumber.TownshipDirectional, _
                              True
    
    ' Range
    basUtilities.AddCodesToCmb g_pFldnames.TLRangeFN, _
                               pTaxlotFlds, _
                               Me.cmbRange, _
                               m_pORMAPNumber.Range, _
                               True
    
    ' Partial Range Code
    basUtilities.AddCodesToCmb g_pFldnames.TLRangePartFN, _
                               pTaxlotFlds, _
                               Me.cmbRangePart, _
                               "0" & m_pORMAPNumber.PartialRangeCode, _
                               True
    
    ' Range Directional
    basUtilities.AddCodesToCmb g_pFldnames.TLRangeDirFN, _
                               pTaxlotFlds, _
                               Me.cmbRangeDir, _
                               m_pORMAPNumber.RangeDirectional, _
                               True
    
    ' Section
    basUtilities.AddCodesToCmb g_pFldnames.TLSectNumberFN, _
                               pTaxlotFlds, _
                               Me.cmbSection, _
                               m_pORMAPNumber.Section, _
                               True
    
    ' Quarter
    basUtilities.AddCodesToCmb g_pFldnames.TLQtrFN, _
                               pTaxlotFlds, _
                               Me.cmbQtr, _
                               m_pORMAPNumber.Quarter, _
                               True
    
    ' QuarterQuarter
    basUtilities.AddCodesToCmb g_pFldnames.TLQtrQtrFN, _
                               pTaxlotFlds, _
                               Me.cmbQtrQtr, _
                               m_pORMAPNumber.QuarterQuarter, _
                               True
    
    ' Map Suffix Type
    basUtilities.AddCodesToCmb g_pFldnames.TLSufTypeFN, _
                               pTaxlotFlds, _
                               Me.cmbSufftype, _
                               basUtilities.ConvertToDescription(pTaxlotFlds, _
                                                                 g_pFldnames.TLSufTypeFN, _
                                                                 m_pORMAPNumber.SuffixType), _
                               True
    
    'Map Suffix Number
    Me.txtSuffNum.Text = m_pORMAPNumber.SuffixNumber
    
    'Anomaly
    txtAnomaly.Text = m_pORMAPNumber.Anomaly
    
    ' Returns the function's value
    InitWithFeature = True
    Exit Function
    
Err_Handler:
    ' Return's the function's value
    InitWithFeature = False
    Resume
End Function

'***************************************************************************
'Name:                  ResetControls
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Clear all edit controls
'Called From:   InitEmpty, InitWithFeature
'Description:   Clears all textbox and combobox controls on the form;
'               setting textboxes to a zero-length string, and resetting the
'               list in the combobox and its listindex property to -1.
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub ResetControls()
    ' Variable declarations
    Dim ctl As Control
    
    ' Loop through controls resetting all text and comboboxes
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
          ElseIf TypeOf ctl Is ComboBox Then
            ctl.Clear
            ctl.ListIndex = -1
        End If
    Next ctl
    
    ' Clean up
    Set ctl = Nothing
End Sub

'***************************************************************************
'Name:                  InitializeFieldPosition
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Initialize the field positions in the MapIndex and Taxlot
'               feature classes
'Called From:   Form_Load
'Description:   Initializes the field positions in the MapIndex and Taxlot
'               feature classes by loading the field indices into two
'               type structions -- one MapIndexFields and one TaxlotFields
'Methods:       None
'Inputs:        pMapIndexFields, pTaxlotClassFields
'Parameters:    None
'Outputs:       None
'Returns:       A boolean value indicating the success of the initialization
'               of the field position structures
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************


Private Function InitializeFieldPositions( _
  ByRef pMapIndexFields As esriGeoDatabase.IFields, _
  ByRef pTaxlotClassFields As esriGeoDatabase.IFields) As Boolean
On Error GoTo Err_Handler
    ' Validate the ORMAP field
    m_pMapIndexFields.ORMAPNumber = pMapIndexFields.FindField(g_pFldnames.MIORMAPMapNumberFN)
    If m_pMapIndexFields.ORMAPNumber = -1 Then
        InitializeFieldPositions = False
        Exit Function
    End If
    
    ' Validate the Reliability field
    m_pMapIndexFields.Reliability = pMapIndexFields.FindField(g_pFldnames.MIReliabFN)
    If m_pMapIndexFields.Reliability = -1 Then
        InitializeFieldPositions = False
        Exit Function
    End If
    
    ' Validate the Scale field
    m_pMapIndexFields.MapScale = pMapIndexFields.FindField(g_pFldnames.MIMapScaleFN)
    If m_pMapIndexFields.MapScale = -1 Then
        InitializeFieldPositions = False
        Exit Function
    End If
    
    ' Validate the Map Number field
    m_pMapIndexFields.MapNumber = pMapIndexFields.FindField(g_pFldnames.MIMapNumberFN)
    If m_pMapIndexFields.MapNumber = -1 Then
        InitializeFieldPositions = False
        Exit Function
    End If
    
    ' Validate the Page field
    m_pMapIndexFields.Page = pMapIndexFields.FindField(g_pFldnames.MIPageFN)
    If m_pMapIndexFields.Page = -1 Then
        InitializeFieldPositions = False
        Exit Function
    End If
    
    ' Returns the function's value
    InitializeFieldPositions = True
    
    ' What happens after here is not absolutely critical
    On Error Resume Next
    
    ' Catalog all of the taxlot fields for potential use
    With m_pTaxlotFields
        .Taxlot = pTaxlotClassFields.FindField(g_pFldnames.TLTaxlotFN)
        .Anomaly = pTaxlotClassFields.FindField(g_pFldnames.TLAnomalyFN)
        .County = pTaxlotClassFields.FindField(g_pFldnames.TLCountyFN)
        .OrmapMapNumber = pTaxlotClassFields.FindField(g_pFldnames.TLOrmapMapNumberFN)
        .OrmapTaxlotNumber = pTaxlotClassFields.FindField(g_pFldnames.TLOrmapTaxlotFN)
        .MapTaxlotNumber = pTaxlotClassFields.FindField(g_pFldnames.TLMapTaxlotFN)
        .PartialRangeCode = pTaxlotClassFields.FindField(g_pFldnames.TLRangePartFN)
        .PartialTownshipCode = pTaxlotClassFields.FindField(g_pFldnames.TLTownPartFN)
        .Quarter = pTaxlotClassFields.FindField(g_pFldnames.TLQtrFN)
        .QuarterQuarter = pTaxlotClassFields.FindField(g_pFldnames.TLQtrQtrFN)
        .Range = pTaxlotClassFields.FindField(g_pFldnames.TLRangeFN)
        .RangeDirectional = pTaxlotClassFields.FindField(g_pFldnames.TLRangeDirFN)
        .Section = pTaxlotClassFields.FindField(g_pFldnames.TLSectNumberFN)
        .SuffixNumber = pTaxlotClassFields.FindField(g_pFldnames.TLSufNumFN)
        .SuffixType = pTaxlotClassFields.FindField(g_pFldnames.TLSufTypeFN)
        .Township = pTaxlotClassFields.FindField(g_pFldnames.TLTownFN)
        .TownshipDirectional = pTaxlotClassFields.FindField(g_pFldnames.TLTownDirFN)
        .SpecialInterest = pTaxlotClassFields.FindField(g_pFldnames.TLSpecInterestFN)
    End With
    Exit Function
    
Err_Handler:
    InitializeFieldPositions = False
End Function

'***************************************************************************
'Name:                  Frame
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Create, if necessary, and return an ESRI ArcGIS Modeless
'               Frame
'Called From:   Multiple Location
'Description:   Creates and returns the Modeless Frame that is necessary to
'               display this form as modeless.
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pFrame
'Returns:       A frame through its IModelessFrame interface
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

'++ START JWalton 1/31/2007
Public Function Frame() As esriFramework.IModelessFrame
    If m_pFrame Is Nothing Then
        Set m_pFrame = New esriFramework.ModelessFrame
        m_pFrame.Create Me
        m_pFrame.Caption = Me.Caption
    End If
    Set Frame = m_pFrame
End Function
'++ END JWalton 1/31/2007

'***************************************************************************
'Name:                  cmbCounty_Validate
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbCounty Control
'Description:   Sets the county value in the ORMAPNumber class, determines
'               if the dataset is valid
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbCounty_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.County = cmbCounty.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  cmbTown_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbTown Control
'Description:   Sets the township value in the ORMAPNumber class, determines
'               if the dataset is valid
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbTown_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.Township = cmbTown.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  cmbTownPart_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbCounty Control
'Description:   Sets the partial township code value in the ORMAPNumber
'               class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbTownPart_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.PartialTownshipCode = cmbTownPart.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  cmbTownDir_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbTownDir Control
'Description:   Sets the township directional value in the ORMAPNumber
'               class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbTownDir_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.TownshipDirectional = cmbTownDir.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  cmbRange_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbRange Control
'Description:   Sets the range value in the ORMAPNumber class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbRange_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.Range = cmbRange.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  cmbRangePart_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbRangePart Control
'Description:   Sets the partial range code value in the ORMAPNumber class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbRangePart_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.PartialRangeCode = cmbRangePart.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  cmbRangeDir_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbRangeDir Control
'Description:   Sets the range directional value in the ORMAPNumber class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbRangeDir_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.RangeDirectional = cmbRangeDir.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  cmbSection_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbSection Control
'Description:   Sets the section in the ORMAPNumber class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbSection_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.Section = cmbSection.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  cmbQtr_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbSection Control
'Description:   Sets the quarter in the ORMAPNumber class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbQtr_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.Quarter = cmbQtr.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  cmbQtrQty_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbSection Control
'Description:   Sets the quarter/quarter in the ORMAPNumber class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbQtrQtr_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.QuarterQuarter = cmbQtrQtr.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  cmbSuffType_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbSection Control
'Description:   Sets the suffix type in the ORMAPNumber class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub cmbSufftype_Click()
    ' Saves the new text to the class
    m_pORMAPNumber.SuffixType = basUtilities.ConvertCode(m_pTaxlotFClass.Fields, _
                                                         g_pFldnames.TLSufTypeFN, _
                                                         cmbSufftype.Text)
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub


'***************************************************************************
'Name:                  cmbSuffNumValidate
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbSection Control
'Description:   Sets the suffix number in the ORMAPNumber class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub txtSuffNum_Change()
    ' Saves the new text to the class
    m_pORMAPNumber.SuffixNumber = txtSuffNum.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  txtAnomalyValidate
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/8/2007
'Purpose:       Determine if the current dataset is valid
'Called From:   cmbSection Control
'Description:   Sets the anomaly in the ORMAPNumber class
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/8/2007    Initial creation
'***************************************************************************

Private Sub txtAnomaly_Change()
    ' Saves the new text to the class
    m_pORMAPNumber.Anomaly = txtAnomaly.Text
    
    ' Enables or disables the save command button
    If m_blnEditState Then
        cmdEditSave.Enabled = m_pORMAPNumber.IsValidNumber
      Else
        cmdEditSave.Enabled = True
    End If
End Sub

'***************************************************************************
'Name:                  m_pORMAPNumber_OnChange
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/8/2007
'Purpose:       Handle the OnChange event of the ORMAP Number
'Called From:   m_pORMAPNumber Object
'Description:   Resets the ORMAP Number textbox
'Methods:       None
'Inputs:        Cancel
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/8/2007    Initial creation
'***************************************************************************

Private Sub m_pORMAPNumber_OnChange(ByVal sNewNumber As String)
    lblORMAPNumber.Caption = sNewNumber
End Sub

'***************************************************************************
'Name:                  UpdateTaxlots
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Update taxlots underlying a changed MapIndex polygon
'Called From:   cmbSection Control
'Description:   Updates any taxlots that underlie a MapIndex polygon when
'               changes to the MapIndex polygon are saved
'Methods:       None
'Inputs:        pFeature, pTaxlots
'Parameters:    None
'Outputs:       Cancel
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Function UpdateTaxlots( _
  ByVal pFeature As esriGeoDatabase.IFeature, _
  ByRef pTaxlots As esriGeoDatabase.IFeatureClass) As Boolean
On Error Resume Next
    ' Variable declarations
    Dim pTaxlotFeature As esriGeoDatabase.IFeature
    Dim pFeatSel As esriGeoDatabase.IFeatureCursor
    Dim pSpatQry As esriGeoDatabase.ISpatialFilter
    Dim pArea As esriGeometry.IArea
    Dim pStatBar As esriSystem.IStatusBar
    Dim lCount As Long
    Dim sSpecialInt As String
    Dim sTaxlot As String

    ' Updates the status bar for the current operation
    Set pStatBar = g_pApp.StatusBar
    pStatBar.Message(esriFramework.esriStatusBarPanes.esriStatusMain) = "Updating underlyling taxlot features..."
    
    ' Finds any taxlots that are underneath the map index polygon
    Set pSpatQry = New esriGeoDatabase.SpatialFilter
    Set pSpatQry.Geometry = pFeature.ShapeCopy
    pSpatQry.SpatialRel = esriSpatialRelContains
    Set pFeatSel = m_pTaxlotFClass.Update(pSpatQry, False)
    
    ' Loops through the selected features
    Set pTaxlotFeature = pFeatSel.NextFeature
    Do While Not pTaxlotFeature Is Nothing
        lCount = lCount + 1
        ' Gets the formatted taxlot value
        If Not IsNull(pTaxlotFeature.Value(m_pTaxlotFields.Taxlot)) Then
            sTaxlot = pTaxlotFeature.Value(m_pTaxlotFields.Taxlot)
            sTaxlot = String(5 - Len(sTaxlot), "0") & sTaxlot
          Else
            sTaxlot = "00000"
        End If
        
        ' Gets the formatted special interest value
        If Not IsNull(pTaxlotFeature.Value(m_pTaxlotFields.SpecialInterest)) Then
            sSpecialInt = pTaxlotFeature.Value(m_pTaxlotFields.SpecialInterest)
            sSpecialInt = String(5 - Len(sSpecialInt), "0") & sSpecialInt
          Else
            sSpecialInt = "00000"
        End If
        
        '@@ START NIS(LCOG) 11/19/2007
        ' Gets the map number value
        Dim sMapNumber As String
        If Not IsNull(pFeature.Value(m_pMapIndexFields.MapNumber)) Then
            sMapNumber = pFeature.Value(m_pMapIndexFields.MapNumber)
        Else
            sMapNumber = ""
        End If
        '@@ END NIS(LCOG) 11/19/2007
        
        ' Copy new attributes to the taxlot table
        '@@ START NIS(LCOG) 11/19/2007
        '@@ DESCR: Add special code for Lane County (see comment below).
        Dim sMapTaxlotID As String
        sMapTaxlotID = m_pORMAPNumber.ORMAPNumber & sTaxlot
        Dim sTLMapTaxlot As String
        Dim iCountyCode As Integer
        'iCountyCode = CInt(Left$(sMapTaxlotID, 2))
        iCountyCode = CInt(g_pFldnames.DefCounty)
        Select Case iCountyCode
        Case 1 To 19, 21 To 36
            sTLMapTaxlot = basUtilities.gfn_s_CreateMapTaxlotValue(sMapTaxlotID, g_pFldnames.MapTaxlotFormatString)
        Case 20
            ' 1.  Lane County uses a 2-digit numeric identifier for ranges.
            '     Special handling is required for east ranges, where 02E is
            '     stored as 25, 03E as 35, etc.
            ' 2.  ORMAP standards (OCDES (pg 13); Taxmap Data Model (pg 11)) assert that
            '     this field should be equal to MAPNUMBER + TAXLOT. In this case, MAPNUMBER
            '     is already in the right format, thus removing the need for the
            '     gfn_s_CreateMapTaxlotValue function. Also, in this case, TAXLOT is padded
            '     on the left with zeros to make it always a 5-digit number (see comment
            '     above).
            ' Trim the map number to only the left 8 characters (no spaces)
            sTLMapTaxlot = Trim$(Left$(sMapNumber, 8)) & sTaxlot
        End Select
        With pTaxlotFeature
            .Value(m_pTaxlotFields.County) = m_pORMAPNumber.County
            .Value(m_pTaxlotFields.Township) = m_pORMAPNumber.Township
            .Value(m_pTaxlotFields.PartialTownshipCode) = m_pORMAPNumber.PartialTownshipCode
            .Value(m_pTaxlotFields.TownshipDirectional) = m_pORMAPNumber.TownshipDirectional
            .Value(m_pTaxlotFields.Range) = m_pORMAPNumber.Range
            .Value(m_pTaxlotFields.PartialRangeCode) = m_pORMAPNumber.PartialRangeCode
            .Value(m_pTaxlotFields.RangeDirectional) = m_pORMAPNumber.RangeDirectional
            .Value(m_pTaxlotFields.Section) = CInt(m_pORMAPNumber.Section)
            .Value(m_pTaxlotFields.Quarter) = m_pORMAPNumber.Quarter
            .Value(m_pTaxlotFields.QuarterQuarter) = m_pORMAPNumber.QuarterQuarter
            .Value(m_pTaxlotFields.SuffixType) = m_pORMAPNumber.SuffixType
            .Value(m_pTaxlotFields.SuffixNumber) = m_pORMAPNumber.SuffixNumber
            .Value(m_pTaxlotFields.Anomaly) = m_pORMAPNumber.Anomaly
            .Value(m_pTaxlotFields.MapNumber) = pFeature.Value(m_pMapIndexFields.MapNumber)
            .Value(m_pTaxlotFields.OrmapMapNumber) = m_pORMAPNumber.ORMAPNumber
            .Value(m_pTaxlotFields.Taxlot) = CLng(sTaxlot)
            .Value(m_pTaxlotFields.SpecialInterest) = sSpecialInt
            '.Value(m_pTaxlotFields.MapTaxlotNumber) = basUtilities.gfn_s_CreateMapTaxlotValue(m_pORMAPNumber.ORMAPNumber & sTaxlot, _
            '                                                                                  g_pFldnames.MapTaxlotFormatString)
            .Value(m_pTaxlotFields.MapTaxlotNumber) = sTLMapTaxlot
            .Value(m_pTaxlotFields.OrmapTaxlotNumber) = m_pORMAPNumber.OrmapTaxlotNumber & sTaxlot
            .Store
        End With
        '@@ END NIS(LCOG) 11/19/2007
        
        ' Get the next feature
        Set pTaxlotFeature = pFeatSel.NextFeature
    Loop
    
    ' Reset the status bar
    pStatBar.Message(esriFramework.esriStatusBarPanes.esriStatusMain) = ""
    
    ' Cleans up
    Set pTaxlotFeature = Nothing
    Set pFeatSel = Nothing
    Set pSpatQry = Nothing
End Function
