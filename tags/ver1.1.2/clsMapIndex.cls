VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdMapIndex"
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
' SCC Revision number: $Revision: 188 $
' Date of last change: $Date: 2008-02-07 16:40:28 -0800 (Thu, 07 Feb 2008) $
'
'
' File name:            cmdMapIndex
'
' Initial Author:       <<Unknown>>
'
' Date Created:         <<Unknown>>
'
' Description:
'       Programmatic description for Map Index tool
'
' Entry points:
'       The tool that this class describes is exposed on the toolbar defined by
'       clsTaxlotEditingToolbar.
'
' Dependencies:
'       File References
'           esriArcMapUI
'           esriCarto
'           esriEditor
'           esriFramework
'           esriGeoDatabase
'           esriSystem
'           esriSystemUI
'       File Dependencies
'           frmMapIndex
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
Private WithEvents m_pEditorEvents As esriEditor.Editor
Attribute m_pEditorEvents.VB_VarHelpID = -1
Private m_pEditor As esriEditor.IEditor2
Private m_pApp As esriFramework.IApplication
Private m_pFrmMapIndex As frmMapIndex
Private m_pBitmap As IPictureDisp
Private m_bEnabled As Boolean
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms
'++ START JWalton 2/12/2007
'   Added Variable Declarations
'       m_pDocEvents, m_pViewEvents, m_pFrmMapIndex
'       m_pBitmap
'++ END JWalton 2/12/2007

'------------------------------
'Private Constants and Enums
'------------------------------
Const c_sModuleFileName As String = "clsMapIndex.cls"

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
'Outputs:       m_pBitmap - The IPictureDisp object that represents the
'                             bitmap.
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

Private Sub Class_Initialize()
  On Error GoTo ErrorHandler
    '++ START JWalton 2/5/2007
    Set m_pBitmap = LoadResPicture("MAPINDEX_BITMAP", vbResBitmap)
    '++ END JWalton
    
    m_bEnabled = False

    Exit Sub
ErrorHandler:
    HandleError True, _
                "Class_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub


'***************************************************************************
'Name:                  Class_Terminate
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Clean up class level variables
'Called From:   Class Object
'Description:   Insures that the editor and its associated events are
'               reset so that they become inactive.
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pEditorEvents - The editor events handler that is
'                                 currently active
'               m_Editor - The local reference to the editor
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

Private Sub Class_Terminate()
  On Error GoTo ErrorHandler

    Set m_pEditorEvents = Nothing
    Set m_pEditor = Nothing
    '++ START JWalton 1/31/2007
        ' Removed clearing of public variables
    '++ END JWalton 1/31/2007
    Exit Sub
ErrorHandler:
    HandleError True, _
                "Class_Terminate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler
    '++ START JWalton 1/31/2007 Added DLL level enablement checking
    If Not g_bDLLEnabled Then
        ICommand_Enabled = False
      Else
        ICommand_Enabled = m_bEnabled And g_bDLLEnabled
    End If
    '++ END JWalton 1/31/2007

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
    
     ICommand_Name = "MapIndex"

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
    
    ICommand_Caption = "Map Index"

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
    
     ICommand_Tooltip = "Map Index Tool"

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
On Error GoTo ErrorHandler
    
     ICommand_Message = ""

    Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Property
 
Private Property Get ICommand_HelpFile() As String
On Error GoTo ErrorHandler
    
     ICommand_HelpFile = ""

    Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Property
 
Private Property Get ICommand_HelpContextID() As Long
    ' Not Implemented
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
On Error GoTo ErrorHandler
  
    '++ START JWalton 2/5/2007 Implemented ICommand_Bitmap
    ICommand_Bitmap = m_pBitmap
    '++ END JWalton 2/5/2007
    
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
'               If the calling application is ArcMap it creates a reference
'               to the Editor as well.
'Methods:       None
'Inputs:        hook - A ArcGIS Application
'Parameters:    None
'Outputs:       g_pApp - Global ArcMap reference
'               g_bDLLEnabled - Global flag that indicates enabled state of
'                               the DLL
'               m_pEditor - Local reference to the Editor
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************

Private Sub ICommand_OnCreate( _
  ByVal hook As Object)
On Error GoTo ErrorHandler
    
    ' Saves a reference to the application and checks to see that it is ArcMap
    If TypeOf hook Is IMxApplication Then
        Set g_pApp = hook
        g_bDLLEnabled = True
      Else
        Set g_pApp = Nothing
        g_bDLLEnabled = False
    End If
    Set m_pApp = hook
    
    '++ START JWalton 1/31/2007
        ' Removed initialization for m_pDoc and m_pMap to ICommand_OnClick
    '++ END JWalton 1/31/2007
    
    'Get a reference to the editor and its events interface
    Dim pId As New esriSystem.UID
    pId = "esriEditor.Editor"
    Set m_pEditor = m_pApp.FindExtensionByCLSID(pId)
    Set m_pEditorEvents = m_pEditor
      
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
'Purpose:       Initialize the command user interface
'Called From:   Class Object
'Description:   Retrieves references to the current document and map, and
'               then open the Map Index form.
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pDoc - Local reference to thencurrent document
'               m_pMap - Local reference to the current map in the current
'                        document
'               m_pFrmMapIndex - Local reference to the Map Index form
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

Private Sub ICommand_OnClick()
On Error GoTo ErrorHandler
    '++ START JWalton 1/31/2007 Variable initializations
    Set m_pDoc = m_pApp.Document
    Set m_pMap = m_pDoc.FocusMap
    Set m_pDocEvents = m_pDoc
    Set m_pViewEvents = m_pMap
    '++ END JWalton 1/31/2007
    
    '++ START JWalton 1/31/2007 Modal form to Modeless form
    If Not g_pForms.IsFormOpen("frmMapIndex") Then
        Set m_pFrmMapIndex = New frmMapIndex
    End If
    If m_pFrmMapIndex.InitForm Then
        m_pFrmMapIndex.Frame.Visible = True
    End If
    '++ END JWalton 1/31/2007

    Exit Sub
ErrorHandler:
    HandleError True, _
                "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  m_pEditorEvents_OnSelectionChanged
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Synchronize the selection with the Map Index form
'Called From:   m_pEditorEvents
'Description:   Determines if there is a single feature selected in the Map
'               Index feature class, and if there is initialize the Map
'               Index form to that feature's data.
'               If there is no selection, or more than one feature selected
'               than the Map Index form is unloaded, and the tool disabled
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_bEnabled - Local flag that indicates the enabled state
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

Private Sub m_pEditorEvents_OnSelectionChanged()
  On Error GoTo ErrorHandler
    'Determine if MapIndex command should be enabled
    Dim pMIFlayer As esriCarto.IFeatureLayer2
    Dim pMIFclass As esriGeoDatabase.IFeatureClass
    
    ' Determines if there is a selection of a single feature
    Set pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then
        Reset
        Exit Sub
    End If
    If HasSelectedFeatures(pMIFlayer) Then
        m_bEnabled = True
    Else
        If g_pForms.IsFormOpen("frmMapIndex") Then Unload m_pFrmMapIndex
        Reset
    End If
    
    ' Determines if this command/tool is currently operating and acts accordingly
    If g_pForms.IsFormOpen("frmMapIndex") Then
        m_pFrmMapIndex.InitForm
    End If

    Exit Sub
ErrorHandler:
    HandleError False, _
                "m_pEditorEvents_OnSelectionChanged " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
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
'Description:   Resets the current tool
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
'Subsequent Author:     <<Type your name here>>
'Created:               2/12/2007
'Purpose:       Insure proper disablement of the control when a map document
'               is changed
'Called From:   m_pDocEvents
'Description:   Resets the current tool
'Methods:       None
'Inputs:        Item - A valid map layer object
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

Private Sub m_pViewEvents_ItemAdded(ByVal Item As Variant)
    ' Variable declarations
    Dim sName As String
    
    ' Determines if this class is either the Map Index or Taxlot feature class
    If TypeOf Item Is esriCarto.IFeatureLayer Then
        sName = Item.Name
        If sName = g_pFldnames.FCMapIndex Or sName = g_pFldnames.FCTaxlot Then
            Reset
        End If
    End If
End Sub

'***************************************************************************
'Name:                  m_pViewEvents_ItemAdded
'Initial Author:        John Walton
'Subsequent Author:     <<Type your name here>>
'Created:               2/12/2007
'Purpose:       Insure proper disablement of the control when a map document
'               is changed
'Called From:   m_pDocEvents
'Description:   Resets the current tool
'Methods:       None
'Inputs:        Item - A valid map layer object
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

Private Sub m_pViewEvents_ItemDeleted(ByVal Item As Variant)
    ' Variable declarations
    Dim pFeatClass As esriCarto.IFeatureLayer
    Dim sName As String
    
    ' Determines if this class is either the Map Index or Taxlot feature class
    If TypeOf Item Is esriCarto.IFeatureLayer Then
        sName = Item.Name
        If sName = g_pFldnames.FCMapIndex Or sName = g_pFldnames.FCTaxlot Then
            ' Resets the current tool
            Reset
        End If
    End If
End Sub

Private Sub Reset()
On Error Resume Next
    ' Hide the taxlot assignment form
    If g_pForms.IsFormOpen("frmMapIndex") Then m_pFrmMapIndex.Frame.Visible = False
    If Not m_pFrmMapIndex Is Nothing Then
        Unload m_pFrmMapIndex
    End If
    
    ' Reset all class level variables
    Set m_pViewEvents = Nothing
    Set m_pDocEvents = Nothing
    Set m_pMap = Nothing
    Set m_pDoc = Nothing
    
    ' Reflects the enabled state of the tool
    m_bEnabled = False
End Sub


