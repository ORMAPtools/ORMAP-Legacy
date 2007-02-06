VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTaxLotEditingToolbar"
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
' SCC Revision number: $Revision$
' Date of last change: $Date$
'
'
' File name:            clsTaxLotEditingToolbar
'
' Initial Author:       Type your name here
'
' Date Created:     10/11/2006
'
' Description: CLASS FOR THE TAXLOTEDITING TOOLBAR
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
'++  JWM 10/11/2006 added this file header

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

'------------------------------
'Private Constants and Enums
'------------------------------

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------


Implements IToolBarDef
 
Private Property Get IToolBarDef_ItemCount() As Long

    IToolBarDef_ItemCount = 5
    
End Property
 
Private Sub IToolBarDef_GetItemInfo(ByVal pos As Long, ByVal itemDef As esriSystemUI.IItemDef)
    
  Select Case pos
    Case 0
      'itemDef.Group = True
      itemDef.ID = "TaxlotEditing.cmdTaxlotAssignment"
    Case 1
      itemDef.ID = "TaxlotEditing.cmdLocate"
    Case 2
      itemDef.ID = "TaxlotEditing.cmdTaxlotCombine"
    Case 3
      itemDef.ID = "TaxlotEditing.cmdArrows"
    Case 4
      itemDef.ID = "TaxlotEditing.cmdMapIndex"
  End Select
    
End Sub
 
Private Property Get IToolBarDef_Name() As String
    
    IToolBarDef_Name = "ORMAP Taxlot Editor"
    
End Property
 
Private Property Get IToolBarDef_Caption() As String

    IToolBarDef_Caption = "ORMAP Taxlot Editor"
    
End Property
