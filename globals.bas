Attribute VB_Name = "basGlobals"
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
' File name:            basGlobals
'
' Initial Author:       <<Unknown>>
'
' Date Created:         <<Unknown>>
'
' Description:
'       Global Variable Declaratoins
'
' Entry points:
'       Constants
'           ORMAP_TAXLOT_FIELD_LENGTH
'           ORMAP_MAPNUM_FIELD_LENGTH
'           ORMAP_TAXLOT_FIELD_LENGTH
'       Variables
'           g_pApp
'               DLL-Wide Application reference to ArcMap
'           g_pFldNames
'               DLL-Wide reference to common field names
'           g_pForms
'               Catalog for all DLL Forms and their current status
'           g_bDLLEnabled
'               DLL-Wide Enabled flag
'
' Dependencies:
'       File References
'           esriFramework
'       File Dependencies
'           clsFieldNames
'           clsFormsCatalog

' Issues:
'       None are known at this time (2/6/2007 JWalton)
'
' Method:
'       None
'
' Updates:
'       10/11/2007 -- Added this file header (JWM)
'       <<Unknown>> --Added the constants for field lengths (JWM)
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)

Option Explicit
'******************************
' Global/Public Definitions
'------------------------------
' Public API Declarations
'------------------------------
'++ START JWalton 2/6/2007
    ' Removed Win32API Function Sleep to basWin32API
'++ END JWalton 2/6/2007

'------------------------------
' Public Enums and Constants
'------------------------------
Public Const ORMAP_MAPNUM_FIELD_LENGTH As Integer = 24
Public Const ORMAP_TAXLOT_FIELD_LENGTH As Integer = 5
Public Const ORMAP_ORTAXLOT_FIELD_LENGTH = ORMAP_MAPNUM_FIELD_LENGTH + ORMAP_TAXLOT_FIELD_LENGTH
'------------------------------
' Public variables
'------------------------------
Public g_pApp As esriFramework.IApplication
Public g_pFldnames As New clsFieldNames
'++ START JWalton 2/6/2007 Additional variable declarations
Public g_pForms As New clsFormsCatalog
Public g_bDLLEnabled As Boolean
'++ END JWalton 2/6/2007
