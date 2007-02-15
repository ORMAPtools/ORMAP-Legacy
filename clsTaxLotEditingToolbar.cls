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
' Initial Author:       <<Unknown>>
'
' Date Created:         10/11/2006
'
' Description:
'       Programmatic description of the ORMAP Taxlot Editor Toolbar for ArcMap
'
'
' Entry points:
'       <<None>>
'
' Dependencies:
'       This class creates a toolbar, ORMAP Taxlot Editor, in ArcMap that exposes the
'       following classes to user control as tools:
'           cmdArrows
'           cmdAutoUpdate
'           cmdLocate
'           cmdMapIndex
'           cmdTaxlotAssignment
'           cmdTaxlotCombine
'
' Issues:
'       None are known at this time (2/8/2007 JWalton)
'
' Method:
'       Standard implementation of IToolbarDef
'
' Updates:
'       10/11/2006 -- Added this file header (JWM)
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)


Option Explicit

' Interface implementations
Implements esriSystemUI.IToolBarDef
 
Private Property Get IToolBarDef_ItemCount() As Long
    ' START JWalton 2/5/2007 Changed count from 5 to 6
    IToolBarDef_ItemCount = 6
    ' END JWalton 2/5/2007
End Property
 
Private Sub IToolBarDef_GetItemInfo( _
  ByVal pos As Long, _
  ByVal itemDef As esriSystemUI.IItemDef)
    ' START JWalton 2/5/2007 Changed order, added auto update, and added groups
    Select Case pos
      Case 0
        itemDef.Group = True
        itemDef.ID = "TaxlotEditing.cmdAutoUpdate"
      Case 1
        itemDef.Group = True
        itemDef.ID = "TaxlotEditing.cmdTaxlotAssignment"
      Case 2
        itemDef.ID = "TaxlotEditing.cmdLocate"
      Case 3
        itemDef.Group = True
        itemDef.ID = "TaxlotEditing.cmdMapIndex"
      Case 4
        itemDef.ID = "TaxlotEditing.cmdTaxlotCombine"
      Case 5
        itemDef.ID = "TaxlotEditing.cmdArrows"
    End Select
    ' END JWalton 2/5/2007
End Sub
 
Private Property Get IToolBarDef_Name() As String
    ' Name to be shown in the Customize...Toolbars dialog list
    IToolBarDef_Name = "ORMAP Taxlot Editor"
End Property
 
Private Property Get IToolBarDef_Caption() As String
    ' Name to be shown in the caption area of the toolbar
    IToolBarDef_Caption = "ORMAP Taxlot Editor"
End Property
