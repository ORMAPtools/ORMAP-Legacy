Attribute VB_Name = "basIniFiles"
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
' SCC Revision number: $Revision: 77 $
' Date of last change: $Date: 2007-02-15 10:24:03 -0800 (Thu, 15 Feb 2007) $
'
'
' File name:            IniFiles
'
' Initial Author:       <<Unknown>>
'
' Date Created:         10/11/2006
'
' Description:
'       Common initialization file routines
'
' Entry points:
'       Methods
'           TestIniFile
'               Returns the state of existence of a given file
'           ReadIniFile
'               Returns the value of a key in a section of an
'               initialization file
' Dependencies:
'       File Dependencies
'           basWin32API
'
' Issues:
'       None are known at this time (2/7/2007 JWalton)

' Method:
'       The ReadIniFile is dependent on the variable m_sIniFile. The
'       function TestIniFile must be called first, or a file must be
'       specified with the ReadIniFile call in order for the function to
'       succeed.
'
' Updates:
'       10/11/2006 -- Added this file header (JWM)
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)

Option Explicit
'******************************
' Private Definitions
'------------------------------
' Private API declarations
'------------------------------
'++ START JWalton 2/7/2007
    ' Removed Win32API function GetPrivateProfileString to basWin32API
'++ END JWalton 2/7/2007
'++ Removed api declaration to write to INI file JWM 10/09/2006
'------------------------------
' Private Variables
'------------------------------
Private m_sIniFile As String

'***************************************************************************
'Name:                  TestIniFile
'Initial Author:        James Moore
'Subsequent Author:     <<Type your name here>>
'Created:               1/24/2001
'Purpose:       Determine if an initialization file is valid
'Called From:   Multiple Locations
'Description:   Checks for a given file one disk, and returns a boolean
'               value indicating the existence.
'Methods:       None
'Inputs:        FilePath
'Parameters:    None
'Outputs:       None
'Returns:       A boolean the represents the existence of the file
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    12-18-06    This new code does not rely on outside functions
'***************************************************************************

Public Function TestIniFile( _
  FilePath As String) As Boolean
    ' Saves the file for later use
    m_sIniFile = FilePath
    
    ' Determines if the file exists
    TestIniFile = Len(Dir$(FilePath)) > 0
End Function

'***************************************************************************
'Name:                  TestIniFile
'Initial Author:        James Moore
'Subsequent Author:     <<Type your name here>>
'Created:               1/24/2001
'Purpose:       Retrieves a value from an ini file corresponding to the
'               section and key name passed.
'Called From:   Multiple Locations
'Description:   File the reads settings in an initialization file.
'Methods:       Call the API with the parameters passed.
'               The lResult value is the length of the string in sReturn,
'               not including the terminating null. If a default value
'               was passed, and the section or key name are not in the file,
'               that value is returned. If no default value was passed (""),
'               then lResult will = 0 if not found.
'Inputs:        FilePath
'Parameters:    None
'Outputs:       None
'Returns:       A string the represents the value of the setting in the
'               initialization file.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    12/18/06    This new code does not rely on outside functions
'John Walton    2/7/2007    Added optional argument for a file
'***************************************************************************

Public Function ReadIniFile( _
  ByVal strSection As String, _
  ByVal strKey As String, _
  Optional ByVal strFile As String = "") As String
    ' Variable declarations
    Dim lResult As Long
    Dim lSize As Long
    Dim sReturn As String
    
    '++ START JWalton 2/7/2007
    If Len(strFile) > 0 Then
        m_sIniFile = strFile
    End If
    '++ END JWalton 2/7/2007
    
    ' Pad a string large enough to hold the data
     sReturn = String$(256, vbNullChar)
     lSize = Len(sReturn)
     
     ' Look up the value in the specified file
     lResult = GetPrivateProfileString(strSection, strKey, "", sReturn, lSize, m_sIniFile)
     
     ' Translate the result to the function
     If lResult Then
         ReadIniFile = Left$(sReturn, lResult)
     Else
         ReadIniFile = ""
     End If
End Function
