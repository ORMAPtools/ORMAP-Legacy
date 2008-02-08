VERSION 5.00
Begin VB.Form frmArrows 
   Caption         =   "Add Arrows"
   ClientHeight    =   3195
   ClientLeft      =   930
   ClientTop       =   2355
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ComboBox cmbArrow 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Text            =   "100 - Anno Arrow"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2400
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
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdDimension 
      Caption         =   "Dimension Arrow"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdArrow 
      Caption         =   "Add Arrow"
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
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   45
      X2              =   2300
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   45
      X2              =   2270
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblCurrentTool 
      Caption         =   "none"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmArrows"
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
'    GNU General Public License for more details located in AppSpecs.bas file
' Keyword expansion for source code control
' Tag for this file : $Name$
' SCC Revision number: $Revision: 188 $
' Date of last change: $Date: 2008-02-07 16:40:28 -0800 (Thu, 07 Feb 2008) $
'
'
'
' File name:        frmArrows
'
' Initial Author:   <<Unknown>>
'
' Date Created:     10/11/2006
'
' Description:
'       Form used to generate hooks and arrows
'
' Entry points:
'       Form Object
'       Properties
'           Arrows
'               The type of arrow that is currently active
'
' Dependencies:
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
'       10/11/06 -- Jim Moore added this header comment section
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)

Option Explicit
'******************************
' Private Definitions
'------------------------------
' Private Variables
'------------------------------
Private m_sArrowType As String

Public Property Get ArrowType() As String
    ArrowType = m_sArrowType
End Property

Private Sub cmdArrow_Click()
    m_sArrowType = "Arrow"
    Me.Caption = "Arrows (Arrow)"
    Me.Hide
End Sub

Private Sub cmdDimension_Click()
    m_sArrowType = "Dimension"
    Me.Caption = "Arrows (Dimension)"
    Me.Hide
End Sub

Private Sub cmdHelp_Click()
    '++ START JWalton 2/7/2006 Centralized Variable Declarations
    ' Variable declarations
    Dim sFilePath As String
    '++ END JWalton
    
    
    sFilePath = app.Path & "\" & "Arrows_help.rtf"
    If basUtilities.FileExists(sFilePath) Then
'++ START JWM 10/16/2006 using new method to open help file
            gsb_StartDoc Me.hwnd, sFilePath
'++ START/END JWM 10/16/2006
    Else
        MsgBox "No help file available in current directory.", vbOKOnly + vbInformation
    End If
End Sub

'++ START JWalton 1/29/2007
Private Sub Form_Load()
    ' Sets the form status to open
    g_pForms.SetFormStatus Me.Name, True
    
    'START Laura Gordon 05/23/07, add additional arrow types
    'Load arrow combo box
    cmbArrow.AddItem "100 - Anno Arrow"
    cmbArrow.AddItem "101 - Hooks"
    cmbArrow.AddItem "102 - Radius Line"
    cmbArrow.AddItem "120 - Station Reference"
    cmbArrow.AddItem "125 - River Arrow"
    cmbArrow.AddItem "134 - Bearing/Distance Arrow"
    cmbArrow.AddItem "136 - Reference Notes"
    cmbArrow.AddItem "137 - Taxlot Arrow"
    cmbArrow.AddItem "141 - Subdivision Arrow"
    cmbArrow.AddItem "147 - DLC Arrow"
    cmbArrow.AddItem "154 - Code Arrow"
    cmbArrow.AddItem "162 - See Map Arrow"
    'END Laura Gordon
    
End Sub

Private Sub Form_QueryUnload( _
  Cancel As Integer, _
  UnloadMode As Integer)
    ' Sets the form status to not open
    g_pForms.SetFormStatus Me.Name, False
End Sub
'++ END JWalton 1/29/2007

'++ START Added by Laura Gordon, 02/20/2007
'***************************************************************************
'Name:                  cmdQuit
'Initial Author:        Laura Gordon
'Subsequent Author:     <<Unknown>>
'Created:               02/20/2007
'Purpose:       Quit the form
'Called From:   cmdQuit Control
'Description:   Unloads the form (note added a quit button to the form - match other forms)
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
'***************************************************************************

Private Sub cmdQuit_Click()
  
    Unload Me

End Sub
'++ END Added by Laura Gordon, 02/20/2007
