VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdLocate"
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
' SCC Revision number: $Revision: 19 $
' Date of last change: $Date: 2006-11-14 13:17:11 -0800 (Tue, 14 Nov 2006) $
'
' File name:            cmdLocate
'
' Initial Author:       Type your name here
'
' Date Created:
'
' Description: CLASS MODULE FOR THE LOCATE COMMAND OPENS frmLocate
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
'               None

Option Explicit
Implements ICommand
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
Private m_pApp As esriFramework.IApplication
Private m_pDoc As esriArcMapUI.IMxDocument
Private m_pMap As IMap
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

 
Private Property Get ICommand_Enabled() As Boolean

    ICommand_Enabled = True
    
End Property
 
Private Property Get ICommand_Checked() As Boolean

    ICommand_Checked = False
    
End Property
 
Private Property Get ICommand_Name() As String

     ICommand_Name = "Locate"
    
End Property
 
Private Property Get ICommand_Caption() As String

     ICommand_Caption = "Locate"
    
End Property
 
Private Property Get ICommand_Tooltip() As String

     ICommand_Tooltip = "Locate Tool"
    
End Property
 
Private Property Get ICommand_Message() As String
    
    ' TODO: Add your implementation here
    ' ICommand_Message =
    
End Property
 
Private Property Get ICommand_HelpFile() As String
    
    ' TODO: Add your implementation here
    ' ICommand_HelpFile =
    
End Property
 
Private Property Get ICommand_HelpContextID() As Long
    
    ' TODO: Add your implementation here
    ' ICommand_HelpContextID =
    
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    
    ' TODO: Add your implementation here
    ' ICommand_Bitmap =
    
End Property
 
Private Property Get ICommand_Category() As String

     ICommand_Category = "ORMAP"
    
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
  Set m_pApp = hook
  If g_pApp Is Nothing Then
    Set g_pApp = hook
  End If
  Set m_pDoc = m_pApp.Document
  Set m_pMap = m_pDoc.FocusMap

End Sub
 
Private Sub ICommand_OnClick()
Dim pTaxlotFlayer As IFeatureLayer
Dim pMIFlayer As IFeatureLayer
    'Get the MapIndex and taxlot feature layers
    Set pTaxlotFlayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If pTaxlotFlayer Is Nothing Then
        MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        GoTo OnClick_Proc_Exit
    End If
    Set pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then
        MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
        GoTo OnClick_Proc_Exit
    End If
Set pTaxlotFlayer = Nothing
Set pMIFlayer = Nothing
    'Initialize the Locate form
    frmLocate.Show vbModal
    
OnClick_Proc_Exit:

Exit Sub

End Sub

