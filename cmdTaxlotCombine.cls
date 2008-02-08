VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdTaxlotCombine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' SCC Revision number: $Revision: 188 $
' Date of last change: $Date: 2008-02-07 16:40:28 -0800 (Thu, 07 Feb 2008) $
'
' File name:            cmdTaxlotCombine
'
' Initial Author:       <<Unknown>>
'
' Date Created:         10/11/2006
'
' Description:
'       Programmatic description for Taxlot Assignment tool
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
'           esriGeoDatabase
'           esriSystemUI
'       File Dependencies
'           frmCombine
'           basGlobals
'
' Issues:
'       None are known at this time (2/8/2007 JWalton)
'
' Method:
'       None
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
Private m_pMap As esriCarto.IMap
Private m_bEnabled As Boolean

'++ START Nick Seigal (LCOG) 12/20/2007 - Support for bitmap
Private m_pBitmap As IPictureDisp
'++ END Nick Seigal (LCOG) 12/20/2007

'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "cmdTaxlotCombine.cls"

'------------------------------
' Interface Implementations
'------------------------------
Implements esriSystemUI.ICommand

'++ START Nick Seigal (LCOG) 12/20/2007 - Added event handler, bitmap
'***************************************************************************
'Name:                  ClassInitialize
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Initialize class level variables
'Called From:   Class Object
'Description:   Loads the bitmap that describes the command on the toolbar
'               from the resource file.
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pBitmap - The toolbar icon that describes the tool
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************

Private Sub Class_Initialize()
  On Error GoTo ErrorHandler
    
    Set m_pBitmap = LoadResPicture("COMBINE_BITMAP", vbResBitmap)
    
    Exit Sub
ErrorHandler:
    HandleError True, _
                "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub
'++ END Nick Seigal (LCOG) 12/20/2007

Private Property Get ICommand_Enabled() As Boolean
On Error GoTo ErrorHandler
    '++ START JWalton 2/8/2007 Centralized Variable Declarations
    ' Variable declarations
    Dim pFeatureLayer As esriCarto.IFeatureLayer
    Dim pDataset As esriGeoDatabase.IDataset
    Dim pFeat As esriGeoDatabase.IFeature
    Dim pFeatCur As esriGeoDatabase.IFeatureCursor
    Dim pWorkspaceEdit As esriGeoDatabase.IWorkspaceEdit
    '++ END JWalton 2/8/2007

    '++ START JWalton 2/6/2007 Link status of command to enabled status of DLL
    ' Checks the status of the DLL before proceding
    If Not g_bDLLEnabled Then
        ICommand_Enabled = False
        Exit Property
    End If
    '++ END JWalton 2/6/2007

    'Enable if editing and at least two taxlots selected
    m_bEnabled = False
    '++ START JWalton 2/8/2007 m_pApp to g_pApp
    Set m_pDoc = g_pApp.Document
    '++ END JWalton 2/8/2007
    Set m_pMap = m_pDoc.FocusMap
    
    'This is called before the mxd is open, so need to check if any layers present
    If m_pMap.LayerCount > 0 Then
        Set pFeatureLayer = basUtilities.FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
        If pFeatureLayer Is Nothing Then
            m_bEnabled = False
            GoTo Process_Exit
        End If
        Set pDataset = pFeatureLayer.FeatureClass
        If pDataset Is Nothing Then GoTo Process_Exit
        Set pWorkspaceEdit = pDataset.Workspace
        If pWorkspaceEdit.IsBeingEdited Then 'Check if being edited
            Set pFeatCur = basUtilities.GetSelectedFeatures(pFeatureLayer) 'Make sure more than one selected
            If Not pFeatCur Is Nothing Then
                Set pFeat = pFeatCur.NextFeature 'first feature
                If Not pFeat Is Nothing Then
                    Set pFeat = pFeatCur.NextFeature 'second feature
                    If Not pFeat Is Nothing Then
                        m_bEnabled = True
                    End If
                End If
            End If
        End If
    End If
    
    'If conditions are met, enable the command
    If m_bEnabled = True Then
        ICommand_Enabled = True
    End If
    
Process_Exit:
  Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Property
 
Private Property Get ICommand_Checked() As Boolean
On Error GoTo ErrorHandler

     ICommand_Checked = False

    Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Property
 
Private Property Get ICommand_Name() As String
On Error GoTo ErrorHandler

     ICommand_Name = "Combine"

    Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Property
 
Private Property Get ICommand_Caption() As String
On Error GoTo ErrorHandler

     ICommand_Caption = "Combine Selected Taxlots"

    Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Property
 
Private Property Get ICommand_Tooltip() As String
On Error GoTo ErrorHandler

     ICommand_Tooltip = "Taxlot Consolidation Tool"

    Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Property
 
Private Property Get ICommand_Message() As String
    '++ START JWalton 2/6/2007
    ' Not Implemented
    '++ END JWalton 2/6/2007
End Property
 
Private Property Get ICommand_HelpFile() As String
    '++ START JWalton 2/6/2007
    ' Not Implemented
    '++ END JWalton 2/6/2007
End Property
 
Private Property Get ICommand_HelpContextID() As Long
    '++ START JWalton 2/6/2007
    ' Not Implemented
    '++ END JWalton 2/6/2007
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
On Error GoTo ErrorHandler
  
    '++ START Nick Seigal (LCOG) 12/20/2007 - Implemented ICommand_Bitmap
    ICommand_Bitmap = m_pBitmap
    '++ END Nick Seigal (LCOG) 12/20/2007
    
    Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Property
 
Private Property Get ICommand_Category() As String
On Error GoTo ErrorHandler

     ICommand_Category = "ORMAP"

    Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
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
'               Initialize the Editor if the calling application is ArcMap
'Methods:       None
'Inputs:        hook - An ArcGIS application that is calling this DLL
'Parameters:    None
'Outputs:       g_pApp - Global reference to ArcMap
'               g_bDLLEnabled -
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

Private Sub ICommand_OnCreate( _
  ByVal hook As Object)
On Error GoTo ErrorHandler
'++ START JWalton 1/26/2007 Only continue for ArcMap
    If TypeOf hook Is esriArcMapUI.IMxApplication Then
        If g_pApp Is Nothing Then
          Set g_pApp = hook
        End If
        g_bDLLEnabled = True
      Else
        g_bDLLEnabled = False
    End If
'++ END JWalton 1/26/2007

    Exit Sub
ErrorHandler:
    HandleError True, _
                "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub
 

'***************************************************************************
'Name:                  ICommand_OnClick
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Show the Combine Taxlots form
'Called From:   Class Object
'Description:   Shows the Combine Taxlots form modally
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
'
'***************************************************************************

Private Sub ICommand_OnClick()
On Error GoTo ErrorHandler
    
    frmCombine.Show vbModal

    Exit Sub
ErrorHandler:
    HandleError True, _
                "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub
