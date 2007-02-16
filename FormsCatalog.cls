VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormsCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Copyright (C) 2006  opet developers opet-developers@lists.sourceforge.net
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version located in AppSpecs.bas file.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along
'    with this program; if not, write to the Free Software Foundation, Inc.,
'    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
' Keyword expansion for source code control
' Tag for this file : $Name$
' SCC Revision number: $Revision: 21 $
' Date of last change: $Date: 2006-11-14 13:24:03 -0800 (Tue, 14 Nov 2006) $
'
' File name:            FormsCatalot
'
' Initial Author:       Type your name here
'
' Date Created:         2/5/2007
'
' Description:
'       Creates, maintains, and coordinates a public registry of forms that are
'       available, open, and closed at any given time during the life of the
'       DLL.
'       This is used especially by command with modeless forms, or forms that are
'       hidden when not in use.  ArcGIS ArcMap will produce a fatal memory error
'       if such forms are not cleaned up before exiting.  The error doesn't seem
'       to do anything but scare the **** out of whoever is getting it.
'
' Entry points:
'       Methods
'           SetFormStatus
'               Register the form's status with the catalog
'           IsFormOpen
'               Returns the currently open status of the requested form
'
' Dependencies:
'       None

' Issues:
'       None are known at this time (2/8/2007 JWalton)
'
' Method:
'       The private constant Number_Of_Forms needs to be set for the proper
'       number of forms in the project
'       The array varForms in Class_Initialize needs to be changed to reflect
'       the names of all forms in the project.
'       The following code needs to be placed in each form for this class to work
'       consistently and properly
'       -------------------------------------------------------------------------
'       Private Sub Form_Load()
'           ' Sets the form status to open
'           g_pForms.SetFormStatus Me.Name, True
'       End Sub
'
'       Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'           ' Sets the form status to not open
'           g_pForms.SetFormStatus Me.Name, False
'       End Sub
'       -------------------------------------------------------------------------
'
' Updates:
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)


Option Explicit
'******************************
' Private Definitions
'------------------------------
' Private API declarations
'------------------------------

'------------------------------
' Private Variables
'------------------------------
Private m_Forms() As Boolean
Private m_FormsIndex() As String

'------------------------------
'Private Constants and Enums
'------------------------------
Private Const Number_Of_Forms = 5

'***************************************************************************
'Name:                  Class_Initialize
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Initialize module level variables
'Called From:   Class Object
'Description:   Initialize module level variables
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_Forms - Status of a given form
'               m_FormsIndex - Name of a given form
'Returns:       None
'Errors:        None
'Assumptions:   This routine raises no known errors.
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************
Private Sub Class_Initialize()
    ' Variable declarations
    Dim varFrm As Variant
    Dim intFrm As Integer
    Dim varForms As Variant
    
    ' Initialize arrays
    ReDim m_Forms(Number_Of_Forms - 1)
    ReDim m_FormsIndex(Number_Of_Forms - 1)

    ' Initialize array of form names
    varForms = Array("frmArrows", "frmCombine", "frmLocate", "frmMapIndex", "frmTaxlotAssignment")
    
    ' Loads and initialize form arrays
    intFrm = 0
    For Each varFrm In varForms
        m_Forms(intFrm) = False
        m_FormsIndex(intFrm) = CStr(varFrm)
        intFrm = intFrm + 1
    Next varFrm
End Sub

'***************************************************************************
'Name:                  SetFormStatus
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Register a form's status
'Called From:   frmArrows.Form_Load
'               frmArrows.Form_QueryUnload
'               frmCombine.Form_Load
'               frmCombine.Form_QueryUnload
'               frmLocate.Form_Load
'               frmLocate.Form_QueryUnload
'               frmMapIndex.Form_Load
'               frmMapIndex.Form_QueryUnload
'               frmTaxlotAssignment.Form_Load
'               frmTaxlotAssignment.Form_QueryUnload
'Description:   Given the name of a form, sForm, and the current status,
'               bStatus.
'               Register sForm with status bStatus
'Methods:       None
'Inputs:        sForm - Form to register
'               bStatus - Status to register
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Sub SetFormStatus(ByVal sForm As String, bStatus As Boolean)
On Error Resume Next
    ' Variable declarations
    Dim iForm As Integer
    
    ' Finds and set the status of a given form
    For iForm = 0 To UBound(m_Forms(), 1)
        If sForm = m_FormsIndex(iForm) Then
            m_Forms(iForm) = bStatus
            Exit Sub
        End If
    Next iForm
End Sub

'***************************************************************************
'Name:                  IsFormOpen
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determine current status of a form
'Called From:   cmdMapIndex.ICommand_OnClick
'               cmdMapIndex.m_pEditorEvents_OnSelectionChanged
'               cmdLocate.Class_Terminate
'               cmdTaxlotAssignment.Class_Terminate
'               cmdTaxlotAssignment.ITool_Cursor
'Description:   Given the name of a form, sForm.
'               Determines if sForm is currently open or not
'Methods:       None
'Inputs:        sForm - The name of the form to check
'Parameters:    None
'Outputs:       None
'Returns:       A boolean value representing the open status of strForm
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Function IsFormOpen(ByVal sForm As String) As Boolean
    ' Variable declarations
    Dim iForm As Integer
    
    ' Finds and returns the status of a given form
    For iForm = 0 To UBound(m_Forms(), 1)
        If sForm = m_FormsIndex(iForm) Then
            IsFormOpen = m_Forms(iForm)
            Exit Function
        End If
    Next iForm
End Function
