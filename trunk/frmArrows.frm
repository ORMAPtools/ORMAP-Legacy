VERSION 5.00
Begin VB.Form frmArrows 
   Caption         =   "Add"
   ClientHeight    =   2475
   ClientLeft      =   930
   ClientTop       =   2355
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdHook 
      Caption         =   "Hook"
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
      TabIndex        =   2
      Top             =   120
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
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdArrow 
      Caption         =   "Arrow"
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
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblCurrentTool 
      Caption         =   "none"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   360
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
' SCC Revision number: $Revision$
' Date of last change: $Date$
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

Private Sub cmdHook_Click()
    m_sArrowType = "Hook"
    Me.Caption = "Arrows (Hook)"
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
End Sub

Private Sub Form_QueryUnload( _
  Cancel As Integer, _
  UnloadMode As Integer)
    ' Sets the form status to not open
    g_pForms.SetFormStatus Me.Name, False
End Sub
'++ END JWalton 1/29/2007
