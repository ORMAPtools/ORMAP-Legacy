VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryInstall"
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
' SCC Revision number: $Revision: 88 $
' Date of last change: $Date: 2007-02-21 09:27:42 -0800 (Wed, 21 Feb 2007) $
'
'
' File name:            CategoryIDs
'
' Initial Author:       JWalton
'
' Date Created:         4/11/2007
'
' Description: Definitions of system GUIDs for ArcGIS component categories.
'
' Entry points:
'       Multiple
'
' Dependencies:
'       File References
'           esriSystem
'
' Issues:
'       <<None>>
'
' Method:
'       Allows use of the CatInstall.exe utility located in the ArcGIS\Bin folder
'       to register the DLL with the operating, and components with ArcGIS.
'
' Updates:
'       4/11/2007 -- Initial implementation. (JWalton)


Option Explicit
'******************************
' Private Definitions
'------------------------------
'------------------------------
' Private Variables
'------------------------------
Private m_strErrorMessage As String

'***************************************************************************
'Name:                  ErrorMessage
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Used by CatInstall.exe
'Called From:   Called by CatInstall.exe
'Description:   Returns any errors to CatInstall.exe
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Public Property Get ErrorMessage() As String
  ErrorMessage = m_strErrorMessage
End Property

'***************************************************************************
'Name:                  ErrorMessage
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Used by CatInstall.exe
'Called From:   Called by CatInstall.exe
'Description:   Installs/Uninstalls specified components according to
'               options specified with CatInstall.exe
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Public Function Install( _
  ByVal bInstall As Boolean, _
  Optional ByVal sCustomOptions As String) As Long
On Error GoTo InstallHandler
     ' Variable declarations
     Dim pCATIDs As CategoryIDs
     Dim pCCM As IComponentCategoryManager
     Dim strPathName As String
     Dim strAppName As String

    ' Initialize objects
     Set pCCM = New ComponentCategoryManager

     ' Reset member error message variable
    m_strErrorMessage = ""

    ' Retrieve the current path of the DLL
    strAppName = app.EXEName
    strPathName = app.Path & "\" & strAppName & ".dll"

    ' The CategoryIDs object makes it easy to specify different
    ' categories to the Component Category Manager
    Set pCATIDs = New CategoryIDs

    ' Register/unregister this dll's exposed classes in the categories we want
    With pCATIDs
        ' Register commands
        pCCM.SetupObject strPathName, _
                         .MakeUID(strAppName & ".cmdAutoUpdate"), _
                         .MxCommands, _
                         bInstall
        pCCM.SetupObject strPathName, _
                         .MakeUID(strAppName & ".cmdTaxlotAssignment"), _
                         .MxCommands, _
                         bInstall
        pCCM.SetupObject strPathName, _
                         .MakeUID(strAppName & ".cmdLocate"), _
                         .MxCommands, _
                         bInstall
        pCCM.SetupObject strPathName, _
                         .MakeUID(strAppName & ".cmdMapIndex"), _
                         .MxCommands, _
                         bInstall
        pCCM.SetupObject strPathName, _
                         .MakeUID(strAppName & ".cmdTaxlotCombine"), _
                         .MxCommands, _
                         bInstall
        pCCM.SetupObject strPathName, _
                         .MakeUID(strAppName & ".cmdArrows"), _
                         .MxCommands, _
                         bInstall
        
        ' Register command bars
        pCCM.SetupObject strPathName, _
                         .MakeUID(strAppName & ".clsTaxlotEditingToolbar"), _
                         .MxCommandBars, _
                         bInstall
    End With

    Install = 0 'Success
    Exit Function
    
InstallHandler:
    ' Set member error message variable and return an error code
    m_strErrorMessage = Err.Description & " Line " & Erl
    Install = Err.Number
End Function
