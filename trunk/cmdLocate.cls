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
' SCC Revision number: $Revision$
' Date of last change: $Date$
'
' File name:            cmdLocate
'
' Initial Author:       <<Unknown>>
'
' Date Created:         <<Unknown>>
'
' Description:
'       Programmatic description for Locate Assignment tool
'
' Entry points:
'       The tool that this class describes is exposed on the toolbar defined by
'       clsTaxlotEditingToolbar.
'
' Dependencies:
'       File References
'           esriArcMapUI
'           esriCarto
'           esriFramework
'           esriSystemUI
'       File Dependencies
'           frmLocate
'           Globals
'
' Issues:
'       None are known at this time (2/8/2007 JWalton)
'
' Method:
'       Describe any complex details that make sense on the file level.  This includes explanations
'       of complex algorithms, how different routines within the module interact, and a description
'       of a data structure used in the module.
'
' Updates:
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)

Option Explicit
'******************************
' Private Definitions
'------------------------------
' Private Variables
'------------------------------
Private m_pDoc As esriArcMapUI.IMxDocument
Private WithEvents m_pDocEvents As esriArcMapUI.DocumentEvents
Attribute m_pDocEvents.VB_VarHelpID = -1
Private m_pMap As esriCarto.IMap
Private WithEvents m_pViewEvents As esriCarto.Map
Attribute m_pViewEvents.VB_VarHelpID = -1
Private m_pFrmLocate As frmLocate
Private m_pToolIcon As IPictureDisp
'++ START JWalton 2/12/2007
'   Added Variable Declarations
'       m_pDocEvents, m_pViewEvents, m_pFrmLocate, m_pToolIcon
'++ END JWalton 2/12/2007

'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName = "cmdLocate"

'------------------------------
' Interface Implementations
'------------------------------
Implements esriSystemUI.ICommand

'***************************************************************************
'Name:                  Class_Initialize
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Initialize the toolbar icon
'Called From:   Class Object
'Description:   Initialize the toolbar icon
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pToolIcon - The IPictureDisp object that represents the
'                             bitmap.
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

'++ START JWalton 1/29/2007
Private Sub Class_Initialize()
    Set m_pToolIcon = LoadResPicture("Locate_Icon", vbResBitmap)
End Sub

'***************************************************************************
'Name:                  Class_Terminate
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Clean up class level variables
'Called From:   Class Object
'Description:   Insures that if the form that represents the user interface
'               if open, that it closes, and that the variable that
'               references it is nullified
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pFormLocate - Variable that refers to the user interface
'                               form
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

Private Sub Class_Terminate()
    If g_pForms.IsFormOpen("frmLocate") Then Unload m_pFrmLocate
    Set m_pFrmLocate = Nothing
End Sub
'++ END JWalton 1/29/2007
 
Private Property Get ICommand_Enabled() As Boolean
'++ START JWalton 1/29/2007
    ' Variable declarations
    Dim pTaxlots As esriCarto.IFeatureLayer
    Dim pMapIndex As esriCarto.IFeatureLayer
    Static lCounter As Long

    ' Initialize objects every fifth iteration
    If lCounter Mod 5 = 0 Then
        Set pTaxlots = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
        Set pMapIndex = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    End If

    ' Determines if the command is available
    ICommand_Enabled = ((Not pTaxlots Is Nothing) And (Not pMapIndex Is Nothing)) And g_bDLLEnabled
'++ END JWalton 1/29/2007
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
    ' Not Implemented
End Property
 
Private Property Get ICommand_HelpFile() As String
    ' Not Implemented
End Property
 
Private Property Get ICommand_HelpContextID() As Long
    ' Not Implemented
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    '++ START JWalton 2/5/2007 Implemention of ICommand_Bitmap
    ICommand_Bitmap = m_pToolIcon
    '++ END JWalton 2/5/2007
End Property
 
Private Property Get ICommand_Category() As String
     ICommand_Category = "ORMAP"
End Property

'***************************************************************************
'Name:                  ICommand_OnCreate
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Initialize the command
'Called From:   Class Object
'Description:   Verify the calling application and the set the enabled state
'               of the DLL itself according to whether or not the calling
'               application is ArcMap.
'Methods:       None
'Inputs:        hook - An ArcGIS application that is calling this DLL
'Parameters:    None
'Outputs:       g_pApp - Global reference to ArcMap
'               g_bDLLEnabled - Global flag that indicates enabled state of
'                               the DLL
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

Private Sub ICommand_OnCreate(ByVal hook As Object)
    '++ START JWalton 2/6/2007
    ' Saves a reference to the application if it is ArcMap
    If TypeOf hook Is esriArcMapUI.IMxApplication Then
        Set g_pApp = hook
        g_bDLLEnabled = True
      Else
        Set g_pApp = Nothing
        g_bDLLEnabled = False
        Exit Sub
    End If
    
    ' Removed object initialization for m_pDoc and m_pMap for ICommand_OnClick
    '++ END JWalton 2/6/2007
End Sub
 
'***************************************************************************
'Name:                  ICommand_OnClick
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Show the locate form
'Called From:   Class Object
'Description:   Verifies the existence of a Map Index and Taxlot feature
'               classes, and then shows the locate form.
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pDoc - The current document in ArcMap
'               m_pMap - The current map in the current document in ArcMap
'               m_pFrmLocate - The locate form
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

Private Sub ICommand_OnClick()
On Error GoTo Err_Handler
    ' Variable declarations
    Dim pTaxlotFLayer As esriCarto.IFeatureLayer
    Dim pMIFlayer As esriCarto.IFeatureLayer

    'Get the MapIndex and taxlot feature layers
    Set pTaxlotFLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If pTaxlotFLayer Is Nothing Then
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

    '++ START JWalton 2/6/2007
    ' Initialize objects
    Set m_pDoc = g_pApp.Document
    Set m_pMap = m_pDoc.FocusMap
    Set m_pDocEvents = m_pDoc
    Set m_pViewEvents = m_pMap
    '++ END JWalton 2/6/2007

    '++ START JWalton 1/29/2007
    'Initialize and show the Locate form
    If Not g_pForms.IsFormOpen("frmLocate") Then
        Set m_pFrmLocate = New frmLocate
    End If
    m_pFrmLocate.Frame.Visible = True
    '++ END JWalton 1/29/2007
    
    ' Clean up
    Set pTaxlotFLayer = Nothing
    Set pMIFlayer = Nothing
    
OnClick_Proc_Exit:
    Exit Sub
    
Err_Handler:
    HandleError True, _
                "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  m_pDocEvents_CloseDocument
'Initial Author:        John Walton
'Subsequent Author:     <<Type your name here>>
'Created:               2/12/2007
'Purpose:       Insure proper disablement of the control when a map document
'               is changed
'Called From:   m_pDocEvents
'Description:   Resets the current tool and all current class level
'               variables
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
'***************************************************************************

Private Sub m_pDocEvents_CloseDocument()
    Reset
End Sub

'***************************************************************************
'Name:                  m_pViewEvents_ItemAdded
'Initial Author:        John Walton
'Subsequent Author:     James Moore
'Created:               2/12/2007
'Purpose:       Insure proper disablement of the control when a map document
'               is changed
'Called From:   m_pDocEvents
'Description:   Resets the current tool and all current class level
'               variables
'Methods:       None
'Inputs:        Item - A valid map layer object
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine will raise an error if another unrelated feature class is added
'                   to the TOC and it does not have a Name property.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'jim moore      5-10-07     Added error handling to trap the error and exit if the Item variant does not have a name property. Tracker number 1710184
'***************************************************************************

Private Sub m_pViewEvents_ItemAdded(ByVal Item As Variant)
On Error GoTo ItemAdded_Error
    ' Variable declarations
    Dim sName As String
    
    ' Determines if this class is either the Map Index or Taxlot feature class
    If TypeOf Item Is esriCarto.IFeatureLayer Then
        sName = Item.Name
        If sName = g_pFldnames.FCMapIndex Or sName = g_pFldnames.FCTaxlot Then
            Reset
        End If
    End If
ItemAdded_Exit:
Exit Sub

ItemAdded_Error:
Select Case Err.Number
Case 424 'Object required
    Resume ItemAdded_Exit
Case Else
HandleError True, _
                "m_pViewEvents_ItemAdded " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Select
End Sub

'***************************************************************************
'Name:                  m_pViewEvents_ItemDeleted
'Initial Author:        John Walton
'Subsequent Author:     James Moore
'Created:               2/12/2007
'Purpose:       Insure proper disablement of the control when a map document
'               is changed
'Called From:   m_pDocEvents
'Description:   Resets the current tool and all current class level
'               variables
'Methods:       None
'Inputs:        Item - A valid map layer object
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine will raise an error if another unrelated feature class is
'                   added to the TOC and it does not have a Name property.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'jim moore      5-10-07     Added error handling to trap the error and exit if the Item variant does not have a name property. Tracker number 1710184
'***************************************************************************

Private Sub m_pViewEvents_ItemDeleted(ByVal Item As Variant)
On Error GoTo ItemDeleted_Error
    ' Variable declarations
    Dim sName As String
    
    ' Determines if this class is either the Map Index or Taxlot feature class
    If TypeOf Item Is esriCarto.IFeatureLayer Then
            sName = Item.Name
            If sName = g_pFldnames.FCMapIndex Or sName = g_pFldnames.FCTaxlot Then
                Reset
            End If
    End If
ItemDeleted_Exit:
Exit Sub

ItemDeleted_Error:
Select Case Err.Number
    Case 424 ' Object required
        Resume ItemDeleted_Exit
    Case Else
        HandleError True, _
                    "m_pViewEvents_ItemDeleted " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                    Err.Number, _
                    Err.Source, _
                    Err.Description, _
                    4
End Select
End Sub

'***************************************************************************
'Name:                  Reset
'Initial Author:        John Walton
'Subsequent Author:     <<Type your name here>>
'Created:               2/12/2007
'Purpose:       Common reset procedure for class
'Called From:
'Description:   Resets the current tool and all current class level
'               variables
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
'***************************************************************************

Private Sub Reset()
On Error Resume Next
    ' Hide the taxlot assignment form
    If g_pForms.IsFormOpen("frmLocate") Then m_pFrmLocate.Frame.Visible = False
    Unload m_pFrmLocate
    
    ' Reset all class level variables
    Set m_pViewEvents = Nothing
    Set m_pDocEvents = Nothing
    Set m_pMap = Nothing
    Set m_pDoc = Nothing
End Sub

