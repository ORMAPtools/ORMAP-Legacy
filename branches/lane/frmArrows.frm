VERSION 5.00
Begin VB.Form frmArrows 
   Caption         =   "Add"
   ClientHeight    =   3135
   ClientLeft      =   930
   ClientTop       =   2355
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   2925
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
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
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
      TabIndex        =   5
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
' File name:            frmArrows
'
' Initial Author:
'
' Date Created:     10/11/2006
'
' Description:  FORM USED TO GENERATE HOOKS AND ARROWS
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
'               Jim Moore added this header comment section on 10/11/06

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
'jwm dead variable Private m_pApp As IApplication
'jwm dead variable Private m_pMxDoc As IMxDocument
'------------------------------
'Private Constants and Enums
'------------------------------

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------

Private Sub cmdArrow_Click()
    Me.lblCurrentTool.Caption = "arrow"
    Me.Hide
End Sub

Private Sub cmdDimension_Click()
    Me.lblCurrentTool.Caption = "dimension"
    Me.Hide
End Sub

Private Sub cmdHelp_Click()
    Dim sFilePath As String
    sFilePath = app.Path & "\" & "Arrows_help.rtf"
    If modUtils.FileExists(sFilePath) Then
'++ START JWM 10/16/2006 using new method to open help file
            gsb_StartDoc Me.hwnd, sFilePath
'++ START/END JWM 10/16/2006
    Else
        MsgBox "No help file available in current directory.", vbOKOnly + vbInformation
    End If
End Sub

Private Sub cmdHook_Click()
    Me.lblCurrentTool.Caption = "hook"
    Me.Hide
End Sub

Private Sub cmdQuit_Click()
    Unload frmArrows
End Sub
