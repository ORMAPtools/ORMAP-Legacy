VERSION 5.00
Begin VB.Form frmDimensionArrowSizes 
   Caption         =   "Dimension Arrows"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   2670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Set to default values of 1.75 and 1.35"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Set dimension arrow variables to enterd values."
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtRatioLine 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "1.75"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtRatioCurve 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "1.35"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblRatioLine 
      Caption         =   "Ratio from the line "
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblRatioCurve 
      Caption         =   "Ratio of the curve"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmDimensionArrowSizes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'++ START Laura Gordon 03/08/2007, form to allow user to control shape of dimension arrows
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
' SCC Revision number: $Revision: 94 $
' Date of last change: $Date: 2007-03-8 09:21:17 -0800 (Tue, 3 Mar 2007) $
'
'
'
' File name:        frmDimensionArrowSizes
'
' Initial Author:   <<Unknown>>
'
' Date Created:     3/08/2007
'
' Description:
'       Form used to set dimension arrow shapes
'
' Entry points:
'       Form Object
'
'
' Dependencies:
'       File Dependencies
'
'
' Issues:
'       None known at this time (3/8/2007 JWalton)
'
' Method:
'       None
'
' Updates:
'

Option Explicit

'******************************
' Private Definitions
'------------------------------
' Private Variables
'------------------------------
Private m_sRatioLine As Double
Private m_sRatioCurve As Double

Public Property Get RatioLine() As Double
    RatioLine = m_sRatioLine
End Property

Public Property Get RatioCurve() As Double
    RatioCurve = m_sRatioCurve
End Property

Private Sub cmdApply_Click()
    'Check to be sure value is numeric, if so set variable to user input
    If IsNumeric(txtRatioLine.Text) Then
        m_sRatioLine = txtRatioLine.Text
    Else
        Call MsgBox("Line ratio text box must be a numeric value, (ie 1.75).", vbCritical, "Invalid Entry")
        txtRatioLine.Text = 1.75
        Exit Sub
    End If
    
    'Check to be sure value is numeric, if so set variable to user input
    If IsNumeric(txtRatioCurve.Text) Then
       m_sRatioCurve = txtRatioCurve.Text
    Else
        Call MsgBox("Curve ratio text box must be a numeric value, (ie 1.35).", vbCritical, "Invalid Entry")
        txtRatioCurve.Text = 1.35
        Exit Sub
    End If
    
    'close the form
    Unload Me
End Sub

Private Sub cmdReset_Click()
    txtRatioLine.Text = 1.75
    txtRatioCurve.Text = 1.35
End Sub

Public Sub Form_Load()
    If m_sRatioLine > 0 And m_sRatioCurve > 0 Then
        txtRatioLine.Text = m_sRatioLine
        txtRatioCurve.Text = m_sRatioCurve
    Else
        txtRatioLine.Text = 1.75
        txtRatioCurve.Text = 1.35
    End If
End Sub
'++ END Laura Gordon
