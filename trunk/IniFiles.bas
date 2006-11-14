Attribute VB_Name = "IniFiles"
Attribute VB_Description = "Set of function to read and extract values from INI files"
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
' File name:            IniFiles
'
' Initial Author:       Type your name here
'
' Date Created:     10/11/2006
'
' Description: MODULE THAT PROCESSES INI FILE
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
'   JWM 10/11/2006 Added this file header

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
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'++ Removed api declaration to write to INI file JWM 10/09/2006
'------------------------------
' Private Variables
'------------------------------
Private mIniFile As String

'------------------------------
'Private Constants and Enums
'------------------------------

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------

Public Function TestIniFile(FilePath As String) As Boolean
  mIniFile = FilePath
  TestIniFile = Len(Dir$(FilePath)) > 0
End Function

Public Function ReadIniFile(ByVal strSection As String, ByVal strKey As String) As String
  
  Dim strBuffer As String
  Dim intPos As Integer
  strBuffer = Space$(255)
  If GetPrivateProfileString(strSection, strKey, "", strBuffer, 255, mIniFile) > 0 Then
    ReadIniFile = RTrim$(StripTerminator(strBuffer))
  Else
    ReadIniFile = ""
  End If
  strBuffer = ""
End Function

Private Function StripTerminator(ByVal strString As String) As String
  'function to strip out chr$(0) from the ReadIniFile function
  Dim intZeroPos As Integer
  intZeroPos = InStr(strString, vbNullChar)
  If intZeroPos > 0 Then
    StripTerminator = Left$(strString, intZeroPos - 1)
  Else
    StripTerminator = strString
  End If
End Function





