Attribute VB_Name = "globals"
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
' SCC Revision number: $Revision: 25 $
' Date of last change: $Date: 2006-11-14 13:37:40 -0800 (Tue, 14 Nov 2006) $
'
'
' File name:            globals
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
'JWM 10/11/2006 added this file header
'jwm added the constants for field lengths
Option Explicit
'******************************
' Global/Public Definitions
'------------------------------
' Public API Declarations
'------------------------------
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'------------------------------
' Public Enums and Constants
'------------------------------
Public Const ORMAP_TAXLOT_FIELD_LENGTH As Integer = 5
Public Const ORMAP_MAPNUM_FIELD_LENGTH As Integer = 24
Public Const ORMAP_ORTAXLOT_FIELD_LENGTH = ORMAP_MAPNUM_FIELD_LENGTH + ORMAP_TAXLOT_FIELD_LENGTH
'------------------------------
' Public variables
'------------------------------
Public g_pApp As esriFramework.IApplication
Public g_pFldnames As New clsFieldNames
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
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms
'------------------------------
'Private Constants and Enums
'------------------------------
'
'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------


