VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRowChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' SCC Revision number: $Revision: 18 $
' Date of last change: $Date: 2006-11-14 13:15:34 -0800 (Tue, 14 Nov 2006) $
'
'
' File name:            clsRowChanged
'
' Initial Author:       Type your name here
'
' Date Created:
'
' Description:
'       Short description of the file's overall purpose.
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
'++  JWM 10/11/2006 added this header to the file

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
Private m_bRowChanged As Boolean 'During the compare and save process, set a flag so the app
' knows whether to store the row when done comparing
'------------------------------
'Private Constants and Enums
'------------------------------

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------

Private Sub Class_Initialize()

  m_bRowChanged = False
  
End Sub

Public Property Let RowChanged(bRowChanged As Boolean)
  m_bRowChanged = bRowChanged
End Property

Public Property Get RowChanged() As Boolean
  RowChanged = m_bRowChanged
End Property