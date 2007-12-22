VERSION 5.00
Begin VB.Form frmLocate 
   Caption         =   "Locate"
   ClientHeight    =   1455
   ClientLeft      =   3270
   ClientTop       =   4605
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   1830
      TabIndex        =   4
      Top             =   1020
      Width           =   800
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   990
      TabIndex        =   3
      Top             =   1020
      Width           =   800
   End
   Begin VB.ComboBox cmbMapNumber 
      Height          =   315
      Left            =   1230
      TabIndex        =   0
      Top             =   210
      Width           =   1335
   End
   Begin VB.TextBox txtTaxlot 
      Height          =   315
      Left            =   1230
      TabIndex        =   1
      Top             =   570
      Width           =   885
   End
   Begin VB.Label Label2 
      Caption         =   "Taxlot:"
      Height          =   225
      Left            =   600
      TabIndex        =   5
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Map Number:"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmLocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' File name:            frmLocate
'
' Initial Author:       <<Unknown>>
'
' Date Created:         10/11/2006
'
' Description:
'       Form used to locate taxlots or map index extents based on a map number and/or
'       taxlot specified by the user
'
' Entry points:
'       Form Object
'       Methods
'           Frame
'               The Frame object that ArcGIS uses to display the form
'
' Dependencies:
'       File References:
'           esriArcMapUI
'           esriCarto
'           esriFramework
'           esriGeoDatabase
'           esriGeometry
'           esriSystem
'       File Dependencies
'           basGlobals
'           basWin32API
'
' Issues:
'       None known at this time (2/6/2007 JWalton
'
' Method:
'       This form is implemented as a modeless form that can sit on top of ArcMap
'       while allowing the user continued access to ArcMap.
'       This implementation is made possible through the Frame property of the
'       form that implements ESRI's Modeless Window interface.
'
' Updates:
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)

Option Explicit
'******************************
' Private Definitions
'------------------------------
'------------------------------
' Private Variables
'------------------------------
Private m_pMxDoc As esriArcMapUI.IMxDocument
Private m_pMIFlayer As esriCarto.IFeatureLayer
Private m_pTaxlotFLayer As esriCarto.IFeatureLayer
Private m_pApp As esriFramework.IApplication
'++ START JWalton 1/29/2007 Variable declarations
Private m_pFrame As esriFramework.IModelessFrame
'++ END JWalton 1/29/2007
Private m_pMIFclass As esriGeoDatabase.IFeatureClass
Private m_pTaxlotFClass As esriGeoDatabase.IFeatureClass
Private m_pMIFields As esriGeoDatabase.IFields
'++ START JWalton 1/29/2007 Variable declarations
Private m_pStatBar As esriSystem.IStatusBar
'++ END JWalton 1/29/2007

'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "frmLocate.frm"

'***************************************************************************
'Name:  cmdApply_Click
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Process the Locate query, zoom to MapIndex or Taxlot, and
'               select the located polygon
'Called From:   cmdApply
'Description:   Event handler for command button cmdApply control
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
'James Moore    10/11/2006      Initial creation
'***************************************************************************

Private Sub cmdApply_Click()
On Error GoTo ErrorHandler
    '++ START JWalton 2/6/2007 Centralized Variable Declarations
    Dim pDoc As esriArcMapUI.IMxDocument
    Dim pMap As esriCarto.IMap
    Dim pFeature As esriGeoDatabase.IFeature
    Dim pFeatureCursor As esriGeoDatabase.IFeatureCursor
    Dim pQueryFilter As esriGeoDatabase.IQueryFilter
    Dim pEnvelope As esriGeometry.IEnvelope
    Dim bMapIndexOnly As Boolean
    Dim sMNum As String
    Dim sTLNum As String
    '++ END JWalton 2/6/2007
  
  
    '++ START JWalton 2/6/2007
    ' Removed dead variables
      
    ' Removed variable initialization for sMsg - No longer needed
    
    ' Busy signal for the user
    Screen.MousePointer = MousePointerConstants.vbHourglass
    '++ END JWalton 2/6/2007
  
  
    '++ START JWM 10/11/2006 trim and then test for length
    sTLNum = Trim$(Me.txtTaxlot.Text)
    sMNum = Trim$(Me.cmbMapNumber.Text)
  
  
    Set pQueryFilter = New esriGeoDatabase.QueryFilter
    If Len(sTLNum) = 0 Then 'Just Query MapIndex
        '++ START JWalton 1/29/2007
        ' Removed first block of if block - No longer needed
        ' Removed dataset validation for the map index feature class
        
        ' Flag for later selection
        bMapIndexOnly = True
        '++ END JWalton 1/29/2007
    
        pQueryFilter.whereClause = "[" & g_pFldnames.MIMapNumberFN & "] = '" & Me.cmbMapNumber.Text & "'"
        '++ START JWalton 1/29/2007
            ' Removed object initialization for pMIFClass
        '++ STOP JWalton 1/26/2007
        Set pFeatureCursor = m_pMIFclass.Search(pQueryFilter, False)
        '++ START JWalton 1/29/2007 ElseIf Operator with Criteria to Else Operator
      Else 'Both values entered
        '++ STOP JWalton 1/29/2007
          pQueryFilter.whereClause = "[" & g_pFldnames.TLMapNumberFN & "] = '" & Me.cmbMapNumber.Text & "' and [" & g_pFldnames.TLTaxlotFN & "]= '" & Me.txtTaxlot & "'"
          Set pFeatureCursor = m_pTaxlotFClass.Search(pQueryFilter, False)
        '++ START JWalton 1/29/2007
          ' Removed else clause - No longer needed
        '++ END JWalton 1/29/2007
    End If
    If pFeatureCursor Is Nothing Then GoTo Process_Exit
    Set pFeature = pFeatureCursor.NextFeature
    
    If pFeature Is Nothing Then
        If Len(sTLNum) = 0 Then
            '++ START 1/29/2007 JWalton Transformed MsgBox to application messages
            m_pStatBar.Message(0) = "Map Index could not be found"
            'With lblMessages
            '    .Caption = "Map Index could not be found"
            '    .Refresh
            'End With
          Else
            m_pStatBar.Message(0) = "Taxlot could not be found"
            'With lblMessages
            '    .Caption = "Taxlot could not be found"
            '    .Refresh
            'End With
            '++ END 1/29/2007 JWalton
        End If
        '++ END JWM 10/11/2006
        Me.txtTaxlot.SetFocus
        GoTo Process_Exit
      Else
        'Zoom to selected feature
        Set pDoc = g_pApp.Document
        Set pMap = pDoc.FocusMap
        Set pEnvelope = pFeature.Shape.Envelope
        ZoomToExtent pEnvelope, m_pMxDoc
        
        ' Select the feature on the map
        pMap.ClearSelection
        If bMapIndexOnly Then
            pMap.SelectFeature m_pMIFlayer, pFeature
          Else
            pMap.SelectFeature m_pTaxlotFLayer, pFeature
        End If
        
        ' Updates the selection inside the envelope
        pDoc.ActiveView.PartialRefresh esriViewGeoSelection, Nothing, pEnvelope
    End If
  
    '++ START JWalton 1/29/2007
      ' Removed the unload command
    '++ END JWalton 1/29/2007
    
Process_Exit:
    '++ START JWalton 1/29/2007
    Screen.MousePointer = MousePointerConstants.vbDefault
    Exit Sub
    '++ END JWalton 1/29/2007
    
ErrorHandler:
  HandleError True, _
              "cmdApply_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
              Err.Number, _
              Err.Source, _
              Err.Description, _
              4
  
  '++ START JWalton 1/29/2007
  ' Necessary because of the mouse cursorassignment
  GoTo Process_Exit
  '++ END JWalton 1/29/2007
End Sub

'++ START JWalton 1/29/2007
    ' Removed cmdCancel_Click() Routine - No longer necessary
'++ END JWalton 1/29/2007

'***************************************************************************
'Name:                  cmdHelp
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:           Help for Locate Tool
'Called From:       cmdHelp
'Description:       Event handler for cmdHelp command button control
'Methods:           None
'Inputs:            None
'Parameters:        None
'Outputs:           None
'Returns:           Nothing
'Errors:            This routine raises no known errors.
'Assumptions:       None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007        Initial creation
'***************************************************************************

Private Sub cmdHelp_Click()
    '++ START JWalton 2/6/2007 Centralized Variable Declarations
    ' Variable declartions
    Dim sFilePath As String
    '++ END JWalton 2/6/2007    Dim sFilePath As String
    
    sFilePath = app.Path & "\Locate_help.rtf"
    If FileExists(sFilePath) Then
        '++ START JWM 10/16/2006 using new method to open help file
        gsb_StartDoc Me.hwnd, sFilePath
        '++ START/END JWM 10/16/2006
    Else
        MsgBox "No help file available in current directory.", vbOKOnly + vbInformation
    End If
End Sub

'***************************************************************************
'Name:                  Form_Initialize
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Initialize the form position
'Called From:   Form Object
'Description:   Event Handler for the form initialize event
'Methods:       The window position is not, and cannot be, set through the
'               form left and top properties.
'               The window position is set using the IWindowPosition
'               interface of the Frame object exposed by ESRI.
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:       None
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub Form_Initialize()
    ' Variable declarations
    Dim pWindowPos As esriFramework.IWindowPosition
    
    ' Initialize objects
    Set pWindowPos = Me.Frame
    
    ' Get the initial position
    pWindowPos.Top = GetSetting("ArcGIS.ArcMap.ORMAP.Tools", "FrmLocate", "Top", 0)
    pWindowPos.Left = GetSetting("ArcGIS.ArcMap.ORMAP.Tools", "FrmLocate", "Left", 0)
    
    ' Makes the form the topmost form
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    
End Sub

'***************************************************************************
'Name:                  Form_Load
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Methods:       Populate the MapIndex dropdown list so user can choose from
'               all available MapIndex values.
'Inputs:        None
'Parameters:    None
'Outputs:       m_pTaxlotFLayer, m_pTaxlotFClass, m_pMIFLayer, m_pMIFClass
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    01-09-07    moved code to load cmbMapNumber combobox to a subroutine.commented dead variables
'***************************************************************************

Private Sub Form_Load()
On Error GoTo Err_Handler
    '++ START JWalton 1/29/2007
    
    ' Initialize objects
    Set m_pStatBar = g_pApp.StatusBar
    Set m_pMxDoc = g_pApp.Document
    
    'Get the MapIndex feature layer and fclass
    Set m_pTaxlotFLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    Set m_pTaxlotFClass = m_pTaxlotFLayer.FeatureClass
    
    ' Get the Taxlot feature layer and class
    Set m_pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    Set m_pMIFclass = m_pMIFlayer.FeatureClass
    '++ END JWalton 1/29/2007

    ' Initialize the combobox control on the form
    fsb_LoadMapNumCombo
    
    '++ START JWalton 1/29/2007 Register the form
    ' Sets the form status to open
    g_pForms.SetFormStatus Me.Name, True
    '++ END JWalton 1/29/2007

Proc_Exit:
    Exit Sub

Err_Handler:
    HandleError True, _
                "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  Form_Unload
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Save the current window position
'Called From:   Form Object
'Description:   Saves the current position of the frame that ESRI ArcMap
'               uses to draw the form
'Methods:       The window position does not, and cannot, come from the
'               form, since the form is modeless.
'               The window position must be obtained through the IWindowPos
'               interface of the Frame object.
'Inputs:        None
'Parameters:    None
'Outputs:       Registry Settings
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

'++ START JWalton 1/29/2007
Private Sub Form_Unload(Cancel As Integer)
    ' Variable declarations
    Dim pWindowPos As esriFramework.IWindowPosition
    
    ' Initialize objects
    Set pWindowPos = Me.Frame
    
    ' Saves the position of the form
    SaveSetting "ArcGIS.ArcMap.ORMAP.Tools", "FrmLocate", "Top", pWindowPos.Top
    SaveSetting "ArcGIS.ArcMap.ORMAP.Tools", "FrmLocate", "Left", pWindowPos.Left
End Sub

'***************************************************************************
'Name:                  Form_QueryUnload
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Unregister the form
'Called From:   Form Object
'Description:   Event handler for Form QueryUnload event
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Sets the form status to not open
    g_pForms.SetFormStatus Me.Name, False
End Sub

'***************************************************************************
'Name:                  Frame
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Create, if necessary, and return an ESRI ArcGIS Modeless
'               Frame
'Called From:   Multiple Location
'Description:   Creates and returns the Modeless Frame that is necessary to
'               display this form as modeless.
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pFrame
'Returns:       A frame through its IModelessFrame interface
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Function Frame() As IModelessFrame
    If m_pFrame Is Nothing Then
        ' Creates the frame if necessary
        Set m_pFrame = New esriFramework.ModelessFrame
        m_pFrame.Create Me
        m_pFrame.Caption = Me.Caption
    End If
    
    ' Returns the frame to the function
    Set Frame = m_pFrame
End Function
'++ END JWalton 1/29/2007

'***************************************************************************
'Name:                  fsb_LoadMapNumCombo
'Initial Author:        James Moore
'Subsequent Author:     Type your name here.
'Created:       01/09/2007
'Called From:   Form_Load
'Description:   Loads cmbMapNumber combo box with values from Mapindex data
'Methods:       Describe any complex details.
'Inputs:        None
'Parameters:    none
'Outputs:       What variables are changed in this routine? None
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    01/09/2007  Initial creation.
'jwm            01-09-07    It appears the Name method of a IDataset object
'                           will only return a Fully Qualified table Name if
'                           you have not changed the display value in the
'                           TOC.  So I have modified the assignment to the
'                           Tables property of the QueryDef object.
'***************************************************************************

Private Sub fsb_LoadMapNumCombo()
On Error GoTo fsb_LoadMapNumCombo_Error
    '++ START JWalton 2/6/2007 Centralized Variable Declarations
    ' Variable declarations
    Dim pCursor As esriGeoDatabase.ICursor
    Dim pDataset As esriGeoDatabase.IDataset
    Dim pFeatureWorkspace As esriGeoDatabase.IFeatureWorkspace
    Dim pQueryDef As esriGeoDatabase.IQueryDef
    Dim pRow As esriGeoDatabase.IRow
    '++ START JWalton 1/29/2007 Variable declarations
    Dim bFeaturesAdded As Boolean
    '++ END JWalton 1/29/2007
    '++ END JWalton 2/6/2007
    
    
    Set pDataset = m_pMIFlayer
    Set pFeatureWorkspace = pDataset.Workspace
    Set pQueryDef = pFeatureWorkspace.CreateQueryDef
    With pQueryDef
        '++ START The pDataset.Name method may or may not return a fully qualified name JWM 01/09/2007
        .Tables = g_pFldnames.FCMapIndex 'This gives us the Fully qualified table name JWM
        '++ END JWM 01/09/2007
        ' Problems with some values -- prevents the form from loading
        .SubFields = "DISTINCT(" & g_pFldnames.TLMapNumberFN & ")"
        Set pCursor = .Evaluate
    End With
    
    Set pRow = pCursor.NextRow
    Do Until pRow Is Nothing
        If Not IsNull(pRow.Value(0)) Then
            '++ START 1/29/2007 JWalton
            ' Flag indicating that an item has been added
            bFeaturesAdded = True
            '++ END 1/29/2007 JWalton
            Me.cmbMapNumber.AddItem pRow.Value(0) ' Note only one field in the cursor
        End If
        Set pRow = pCursor.NextRow
    Loop
    
    '++ START 1/29/2007 JWalton Selects the first item in the dropdown if there is one
    If bFeaturesAdded Then
        Me.cmbMapNumber.ListIndex = 0
        Me.cmbMapNumber.Text = Me.cmbMapNumber.List(0)
    End If
    '++ END 1/29/2007 JWalton
    
fsb_LoadMapNumCombo_Exit:
    Exit Sub
    
fsb_LoadMapNumCombo_Error:
    HandleError True, _
                "fsb_LoadMapNumCombo " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  cmdNumber_Change
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Validation of user choices
'Called From:   cmdMapNumber Control
'Description:   Validates the choices of the user and either enables or
'               enables or disables the Locate command button accordingly
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

'++ START JWalton 1/29/2007 Event handlers for control of Find Button
Private Sub cmbMapNumber_Change()
    If CBool(Len(cmbMapNumber.Text)) Then
        cmdApply.Enabled = True
      Else
        cmdApply.Enabled = False
    End If
    m_pStatBar.Message(0) = ""
End Sub

'***************************************************************************
'Name:                  txtTaxlot_Change
'Initial Author:        <Unknown>
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Clear the status bar message
'Called From:   txtTaxlot Control
'Description:   Clears the status bar message when the entry in txtTalot
'               changes
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub txtTaxlot_Change()
    m_pStatBar.Message(0) = ""
End Sub
'++ END JWalton 1/29/2007
