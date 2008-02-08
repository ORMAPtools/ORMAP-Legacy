VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdArrows"
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
' File name:            cmdArrows
'
' Initial Author:       Type your name here
'
' Date Created:
'
'       Programmatic description for Arrows tool
'
' Entry points:
'       The tool that this class describes is exposed on the toolbar defined by
'       clsTaxlotEditingToolbar.
'
' Dependencies:
'       File References
'           esriArcMapUI
'           esriCarto
'           esriDisplay
'           esriFramework
'           esriGeoDatabase
'           esriGeomtry
'           esriSystemUI
'           esriSystemUI
'       File Dependencies
'           frmArrows
'           Globals
'
' Issues:
'
'
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

'------------------------------
' Interface Implementations
'------------------------------
Implements esriSystemUI.ICommand
Implements esriSystemUI.ITool

'******************************
' Private Definitions
'------------------------------
' Private Variables
'------------------------------
Private m_pDoc As esriArcMapUI.IMxDocument
Private m_pMap As esriCarto.IMap
Private m_pLineSymbol As esriDisplay.ILineSymbol
Private m_pApp As esriFramework.IApplication
Private m_pArrowPt1 As esriGeometry.IPoint
Private m_pArrowPt2 As esriGeometry.IPoint
Private m_pArrowPt3 As esriGeometry.IPoint
Private m_pArrowPt4 As esriGeometry.IPoint
Private m_pFromBreakPoint As esriGeometry.IPoint        ' hooks
Private m_pStartPoint As esriGeometry.IPoint            ' hooks
Private m_pTextPoint As esriGeometry.IPoint
Private m_pToBreakPoint As esriGeometry.IPoint          ' hooks
Private m_pLinePolyline As esriGeometry.IPolyline
Private m_pTextSymbol As esriDisplay.ITextSymbol
Private m_dHookAngle As Double
Private m_bDoOnce As Boolean                            ' testing
Private m_bInUse As Boolean

'++ START Added by Laura Gordon, 02/20/2007
Private m_pArrowPtTemp As IPoint
Private m_pArrowPtTemp2 As IPoint
Private m_pPt As IPoint
Private m_pSnapAgent As IFeatureSnapAgent
Private m_bMouseHasMoved As Boolean
Private m_bInTol As Boolean
Private m_pMxApp As IMxApplication
Private m_pMxDoc As IMxDocument
Private m_pActive As IActiveView
Private WithEvents m_pEditorEvents As Editor
Attribute m_pEditorEvents.VB_VarHelpID = -1
Private m_pEditor As IEditor2
Private m_bToolJustCompletedTask As Boolean
'++ END Added by Laura Gordon, 02/20/2007

'++  START Laura Gordon, 03/08/2007
Private m_bolDimensionChanged As Boolean
'++  END Laura Gordon

'++ START Laura Gordon 05/23/07, changing arrow draw to allow for more arrow types
    Dim colArrowPoints As Collection
'++ END Laura Gordon

'++ START Nick Seigal (LCOG) 12/20/2007 - Support for bitmap
Private m_pBitmap As IPictureDisp
'++ END Nick Seigal (LCOG) 12/20/2007
'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "clsArrows.cls"

'++ START Nick Seigal (LCOG) 12/20/2007 - Added event handler, bitmap
'***************************************************************************
'Name:                  ClassInitialize
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Initialize class level variables
'Called From:   Class Object
'Description:   Loads the bitmap that describes the tool on the toolbar
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
    
    Set m_pBitmap = LoadResPicture("ARROWS_BITMAP", vbResBitmap)
    
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

    '++ START JWalton 2/6/2007 Enabled status dependent on DLL enabled state
    ICommand_Enabled = True And g_bDLLEnabled
    '++ END JWalton 2/6/2007

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
    
    ICommand_Name = "AddArrowsAndHooks"

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

    '++ START Laura Gordon 05/23/07, Adding additional arrow types, name change no longer hooks all combined in one button
    ICommand_Caption = "Add Arrows"
    'ICommand_Caption = "Add Arrows/Hooks"
    '++ END Laura Gordon
    
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

    ICommand_Tooltip = "Open a dialog to add arrows and hooks"

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
    
    ICommand_Message = "Initialize the Add Arrows and Hooks form"

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
On Error GoTo ErrorHandler

    ICommand_HelpContextID = ""

    Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_HelpContextID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
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
'Subsequent Author:     Type your name here.
'Created:               <<Unknown>>
'Purpose:       Event Handler for ICommand interface OnCreate Event
'Called From:   Class Object
'Description:   Saves a reference to the calling application if it is
'               ArcMap, and enables or disables the DLL accordingly.
'Methods:       None
'Inputs:        hook - A ArcGIS Application
'Parameters:    None
'Outputs:       g_pApp - Global ArcMap reference
'               g_bDLLEnabled - Global flag that indicates enabled state of
'                               the DLL
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
    
    '++ START JWalton 2/7/2007
    ' Saves a reference to the application and checks to see that it is ArcMap
    If TypeOf hook Is esriArcMapUI.IMxApplication Then
        Set g_pApp = hook
        g_bDLLEnabled = True
      Else
        Set g_pApp = Nothing
        g_bDLLEnabled = False
        Exit Sub
    End If
    
        ' Removed initialization of m_pDoc and m_pMap to ICommand_OnClick
    '++ END JWalton 2/7/2007

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
'Subsequent Author:     Type your name here.
'Created:               <<Unknown>>
'Purpose:       Activate the arrows command
'Called From:   Class Object
'Description:   Verifies that the current document is a valid document to
'               use this tool on, and enable or disable this command
'               accordingly.
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
'***************************************************************************

Private Sub ICommand_OnClick()
On Error GoTo ErrorHandler
    'Get the MapIndex feature layer and fclass
    Dim pCLFLayer As esriCarto.IFeatureLayer
    Dim pTaxlotFLayer As esriCarto.IFeatureLayer
    Dim pMIFlayer As esriCarto.IFeatureLayer
    
    ' Initialize document and map references
    Set m_pDoc = g_pApp.Document
    Set m_pMap = m_pDoc.FocusMap

    ' Validate the existence of the Taxlot feature class
    Set pTaxlotFLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If pTaxlotFLayer Is Nothing Then
        MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        GoTo Process_Exit
    End If
    
    ' Validate the existence of the Map Index feature class
    Set pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then
        MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
        GoTo Process_Exit
    End If
    
    ' Validate the existence of the Cartographic Lines feature class
    Set pCLFLayer = FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
    If pCLFLayer Is Nothing Then
        MsgBox "Unable to locate Cartographic Lines layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCCartoLines
        GoTo Process_Exit
    End If
    
    '++ START Laura Gordon 05/23/07, changing arrow draw to allow for more arrow types
    Set colArrowPoints = New Collection
    '++ END Laura Gordon
    
    ' Shows the options form
    frmArrows.Show vbModal

Process_Exit:
    Exit Sub
    
ErrorHandler:
    HandleError True, _
                "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub
 
Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    ' Not Implemented
End Property

Private Sub ITool_OnDblClick()
    ' Not Implemented
End Sub
 
Private Sub ITool_OnKeyDown(ByVal keyCode As Long, ByVal Shift As Long)
  On Error GoTo ErrorHandler
    
    '++ START Added by Laura Gordon 02/20/2007
    'End the dimension arrow tool if the "q" key is pressed
    If keyCode = vbKeyQ Then
        If frmArrows.ArrowType = "Dimension" Then
            'Deactivate the tool
            Set m_pArrowPt1 = Nothing
            Set m_pArrowPt2 = Nothing
            Set m_pArrowPt3 = Nothing
            Set m_pArrowPt4 = Nothing
            Set m_pPt = Nothing
            Set m_pArrowPtTemp = Nothing
            Set m_pArrowPtTemp2 = Nothing
            Set g_pApp.CurrentTool = Nothing
            g_pApp.RefreshWindow
            m_bToolJustCompletedTask = True
            
        '++ START Laura Gordon 05/23/07, adding additional arrow options types
        ElseIf frmArrows.ArrowType = "Arrow" Then
            'Deactivate the tool
            Set m_pArrowPt1 = Nothing
            Set m_pArrowPt2 = Nothing
            Set m_pArrowPt3 = Nothing
            Set m_pArrowPt4 = Nothing
            Set colArrowPoints = Nothing
            Set g_pApp.CurrentTool = Nothing
            g_pApp.RefreshWindow
            m_bToolJustCompletedTask = True
        '++ END Laura Gordon
        Else 'Hooks
            'Deactivate the tool
            Set g_pApp.CurrentTool = Nothing
            g_pApp.RefreshWindow
            m_bToolJustCompletedTask = True
        End If
    End If
    '++  START Laura Gordon 03/08/07, call form
    If keyCode = vbKeyD Then
        If frmArrows.ArrowType = "Dimension" Then
            m_bolDimensionChanged = True
            frmDimensionArrowSizes.Show vbModal
        End If
    End If
    '++  END Laura Gordon

  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnKeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
'++ END Added by Laura Gordon 02/20/2007
End Sub
 
Private Sub ITool_OnKeyUp(ByVal keyCode As Long, ByVal Shift As Long)
    ' Not Implemented
End Sub
 
Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean
    ' Not Implemented
End Function

'***************************************************************************
'Name:                  ITool_Deactivate
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Clean up after the tool
'Called From:   Class Object
'Description:   Hides the form and deactivates the tool
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
'***************************************************************************
  
Private Function ITool_Deactivate() As Boolean
On Error GoTo ErrorHandler
    
    ' Reset tool level variables
    Set m_pTextSymbol = Nothing
    Set m_pTextPoint = Nothing
    Set m_pLinePolyline = Nothing
    Set m_pLineSymbol = Nothing
    m_bInUse = False

    '++ START Laura Gordon 05/23/07, changing arrow draw to allow for more arrow types
    Set colArrowPoints = Nothing
    '++ END Laura Gordon

    ' Allow the tool to deactivate
    ITool_Deactivate = True


    Exit Function
ErrorHandler:
    HandleError True, _
                "ITool_Deactivate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  ITool_OnMouseDown
'Initial Author:        <<Unknown>>
'Subsequent Author:     Type your name here.
'Created:               <<Unknown>>
'Purpose:       Handle the MouseDown event for the defined tool
'Called From:   Class Object
'Description:   Create a record of the point the mouse was clicked for
'               annotation arrows, and draws them when enough points have
'               been entered.
'               Create a record of the first point the mouse was clicked for
'               creating a land hook, and let the OnMouseMove, and OnMouseUp
'               events finish the work.
'Methods:       None
'Inputs:        Button, Shift, X, Y
'Parameters:    None
'Outputs:       m_bDoOnce - Initialization flag
'               m_pArrowP1 - The first point for an annotation arrow
'               m_pArrowP2 - The second point for an annotation arrow
'               m_pArrowP3 - The third point for an annotation arrow
'               m_pArrowP4 - The fourth point for an annotation arrow
'               m_pLinePolyline - The temporary polyline
'               m_pLineSymbol - Line Symbol of the temporary line
'               m_pTextPoint - Point for temporary text
'               m_pTextSymbol -
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point
'JWM            10/16/2006  Using case instead of if
'Laura Gordon   02/20/2007  Adding dimension arrows
'Laura Gordon   3/8/2007    Added required field calc (automethod, autowho, autodate, mapnumber, mapscale) and dimension arrow size changes
'Laura Gordon   05/23/07    Adding additional arrow type options
'***************************************************************************

Private Sub ITool_OnMouseDown( _
  ByVal Button As Long, _
  ByVal Shift As Long, _
  ByVal X As Long, _
  ByVal Y As Long)
On Error GoTo ErrorHandler
    '++ START JWalton 2/8/2007 Centralized variable declarations
    ' Variable declarations
    Dim pMxDoc As esriArcMapUI.IMxDocument
    Dim pActiveView As esriCarto.IActiveView
    Dim pArrowLayer As esriCarto.IFeatureLayer
    Dim pMIFL As esriCarto.IFeatureLayer
    Dim pDSet As esriGeoDatabase.IDataset
    Dim pFeature As esriGeoDatabase.IFeature
    Dim pArrowFC As esriGeoDatabase.IFeatureClass
    Dim pMIFC As esriGeoDatabase.IFeatureClass
    Dim pArrowLine As esriGeometry.IPolyline4
    Dim pArrowPoints As esriGeometry.IPointCollection4
    Dim pWSEdit As esriGeoDatabase.IWorkspaceEdit
    Dim lCLMNfld As Long
    Dim lLineTypeFld As Long
    Dim sCurMapNum As String
    Dim sScale As String
    Dim sToolCaption As String
    '++ END JWalton 2/8/2007
    '++ START Laura Gordon 02/20/2007
    Dim sCurrentMapScale As String
    Dim sSide As String
    Dim iChange As Integer
    '++ END Laura Gordon 02/20/2007
    
    '++ START Laura Gordon 03/08/2007
    Dim dRatioLine As Double
    Dim dRatioCurve As Double
    '++ END Laura Gordon
    
    '++ START Laura Gordon 05/21/07, adding user input smooth option
    Dim dSmoothRatio As Double
    '++ END Laura Gordon
    
    '++ START Laura Gordon 05/23/07, changing arrow draw to allow for more arrow types
    'Dim colArrowPoints As Collection
    Dim bolFinish As Boolean
    Dim lngIndex As Long
    '++ END Laura Gordon
    
    ' Set the in use flag
    m_bInUse = True
    
    ' Initialize active view
    '++ START JWalton 2/8/2007 m_pApp to g_pApp
    Set pMxDoc = g_pApp.Document
    '++ END JWalton
    Set pActiveView = pMxDoc.FocusMap
  
    ' Get reference to MapIndex FC
    Set pMIFL = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    Set pMIFC = pMIFL.FeatureClass
    
    '++ START JWalton 2/9/2007 Form control to Form Property
    sToolCaption = UCase$(frmArrows.ArrowType)
    '++ END JWalton 2/9/2007
    Select Case sToolCaption
        '++ Start Laura Gordon 02/20/2007, Had to UCase HOOK because it wasn't matching
        Case "HOOK"  '"Hook" 'If drawing hooks
        '++ End Laura Gordon, 02/20/2007
            ' Get point to measure distance from
            Set m_pStartPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
            m_bDoOnce = False 'testing
            Set m_pFromBreakPoint = m_pStartPoint
            
            ' Get the scale of the current mapindex
            '++ START Nick Seigal (LCOG) 01/25/2008 - Added sort-by field
            sScale = GetValueViaOverlay(m_pStartPoint, pMIFC, g_pFldnames.MIMapScaleFN, g_pFldnames.MIMapNumberFN)
            '++ END Nick Seigal (LCOG) 01/25/2008
            sScale = ConvertToDescription(pMIFC.Fields, g_pFldnames.MIMapScaleFN, sScale)
        
            
        Case "ARROW" 'If drawing annotation arrows
            'If m_bToolJustCompletedTask = False Then
                If bolFinish = False Then
                    If Button = 1 And Shift = 1 Then
                        '[Left mouse click plus Shift key down...]
                        If colArrowPoints.Count = 0 Then
                        '[Save the first point...]
                            If m_pArrowPt1 Is Nothing Then
                                Set m_pArrowPt1 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                                colArrowPoints.Add m_pArrowPt1
                                
                                'Clear existing point
                                Set m_pArrowPt1 = Nothing
                            End If
                            
                        'Save the last point
                        Else
                        '[Save the last point...]
                            If m_pArrowPt1 Is Nothing Then
                                Set m_pArrowPt1 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                                colArrowPoints.Add m_pArrowPt1
                                DrawArrows
                                
                                'Clear existing point
                                Set m_pArrowPt1 = Nothing
                                bolFinish = True
                            End If
                        End If
                        
                    ElseIf Button = 1 And Shift = 0 Then
                    '[Left mouse click...]
                        'Add vertex
                        If m_pArrowPt1 Is Nothing Then
                            Set m_pArrowPt1 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                            colArrowPoints.Add m_pArrowPt1
                            
                            'Clear existing point
                            Set m_pArrowPt1 = Nothing
                        End If
                    End If
                End If
    '            If m_pArrowPt1 Is Nothing Then
    '                ' Save the first point
    '                Set m_pArrowPt1 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
    '            ElseIf m_pArrowPt2 Is Nothing Then
    '                ' Saves the second point and draws the arrow
    '                Set m_pArrowPt2 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
    '                DrawArrows
    '            ElseIf m_pArrowPt3 Is Nothing Then
    '                ' Saves the third point and draws the arrow
    '                Set m_pArrowPt3 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
    '                DrawArrows
    '            ElseIf m_pArrowPt4 Is Nothing Then
    '                ' Saves the fourth point and draws the arrow
    '                Set m_pArrowPt4 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
    '                DrawArrows
                
                If colArrowPoints.Count > 1 And bolFinish = True Then
                    
                    ' Creates a new polyline from the points and smoothes it
                    Set pArrowPoints = New esriGeometry.Polyline
                    
                    For lngIndex = 1 To colArrowPoints.Count
                        pArrowPoints.AddPoint colArrowPoints.Item(lngIndex)
                    Next
                    
                    Set pArrowLine = pArrowPoints
                    pArrowLine.Smooth pArrowLine.Length / 10
                    
    '                ' Creates a new polyline from the points and smoothes it
    '                Set pArrowPoints = New esriGeometry.Polyline
    '                pArrowPoints.AddPoint m_pArrowPt1
    '                pArrowPoints.AddPoint m_pArrowPt2
    '                pArrowPoints.AddPoint m_pArrowPt3
    '                pArrowPoints.AddPoint m_pArrowPt4
    '                Set pArrowLine = pArrowPoints
    '                pArrowLine.Smooth pArrowLine.Length / 10
            
                    ' Get a reference to the Cartographic Lines feature class
                    Set pArrowLayer = FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
                    If pArrowLayer Is Nothing Then
                        MsgBox "The layer, " & g_pFldnames.FCCartoLines & ", is not in the map.", vbExclamation, "Layer not found"
                        GoTo Process_Exit
                    End If
                    Set pArrowFC = pArrowLayer.FeatureClass
                    Set pDSet = pArrowFC
                    Set pWSEdit = pDSet.Workspace
                    
                    ' Start an edit operation to encompass the creation of the feature
                    pWSEdit.StartEditOperation
                    
                    ' Create the arrow feature
                    Set pFeature = pArrowFC.CreateFeature
                    Set pFeature.Shape = pArrowLine
                    
                    ' Locates fields in the feature's dataset
                    lCLMNfld = LocateFields(pArrowFC, g_pFldnames.MIMapNumberFN)
                    lLineTypeFld = LocateFields(pArrowFC, g_pFldnames.CLinesLineTypeFN)
                    
                    ' Insure that the feature's fields are found
                    If lLineTypeFld = -1 Then GoTo Process_Exit
                    
                    ' Populate field values in the feature
                    '++ START Nick Seigal (LCOG) 01/25/2008 - Added sort-by field
                    sCurMapNum = GetValueViaOverlay(pFeature.Shape, pMIFC, g_pFldnames.MIMapNumberFN, g_pFldnames.MIMapNumberFN)
                    '++ END Nick Seigal (LCOG) 01/25/2008
                    pFeature.Value(lCLMNfld) = sCurMapNum
                    '++ START Laura Gordon 05/23/07, adding additional arrow types
                    'pFeature.Value(lLineTypeFld) = 100
                    pFeature.Value(lLineTypeFld) = Left(frmArrows.cmbArrow.Text, 3) '100
                    '++  END Laura Gordon
                    
                    '++ START Laura 03/08/07, calc additional fields
                    'Set the AutoMethod Field
                    lLineTypeFld = LocateFields(pArrowFC, g_pFldnames.AutoMethodFN)
                    If lLineTypeFld = -1 Then Exit Sub
                    pFeature.Value(lLineTypeFld) = "UNK"
                    
                    'Set the AutoWho Field
                    lLineTypeFld = LocateFields(pArrowFC, g_pFldnames.AutoWhoFN)
                    If lLineTypeFld = -1 Then Exit Sub
                    pFeature.Value(lLineTypeFld) = UserName
                    
                    'Set the AutoDate Field
                    lLineTypeFld = LocateFields(pArrowFC, g_pFldnames.AutoDateFN)
                    If lLineTypeFld = -1 Then Exit Sub
                    pFeature.Value(lLineTypeFld) = Format(Now, "MM/DD/YYYY")
                    
                    'Set the MapScale Field
                    '++ START Nick Seigal (LCOG) 01/25/2008 - Added sort-by field
                    sCurrentMapScale = GetValueViaOverlay(pFeature.Shape, pMIFC, g_pFldnames.MIMapScaleFN, g_pFldnames.MIMapNumberFN)
                    '++ END Nick Seigal (LCOG) 01/25/2008
                    lLineTypeFld = LocateFields(pArrowFC, g_pFldnames.MIMapScaleFN)
                    If lLineTypeFld = -1 Then Exit Sub
                    '++ START Nick Seigal (LCOG) 12/20/2007 - Fixed data type
                    pFeature.Value(lLineTypeFld) = Val(sCurrentMapScale)
                    '++ END Nick Seigal (LCOG) 12/20/2007
                    '++  End Laura Gordon
                            
                     '++ START Nick Seigal (LCOG) 12/20/2007 - Moved store to after all fields are edited
                    pFeature.Store
                    '++ END Nick Seigal (LCOG) 12/20/2007
                    ' Finalize the edit operation
                    pWSEdit.StopEditOperation
                    
                    '++ START Nick Seigal (LCOG) 12/20/2007 - Better refresh
                    'Partially refresh the display
                    pActiveView.PartialRefresh _
                            esriViewBackground Or _
                            esriViewGeography Or _
                            esriViewGraphics Or _
                            esriViewGraphicSelection, Nothing, pFeature.Extent.Envelope
                    '++ END Nick Seigal (LCOG) 12/20/2007
                    
                    ' Clean up
                    Set pArrowLine = Nothing
                    Set m_pArrowPt1 = Nothing
                    Set m_pArrowPt2 = Nothing
                    Set m_pArrowPt3 = Nothing
                    Set m_pArrowPt4 = Nothing
                    Set m_pTextSymbol = Nothing
                    Set m_pTextPoint = Nothing
                    Set m_pLinePolyline = Nothing
                    Set m_pLineSymbol = Nothing
                    
                    'START Laura Gordon 05/23/07, adding arrow options
                    Set colArrowPoints = Nothing
                    Set colArrowPoints = New Collection
                    'END Laura Gordon
                
                    ' Deactivate the tool
                    '++ START JWalton 2/8/2007 m_pApp to g_pApp
                    'Set g_pApp.CurrentTool = Nothing
                    '++ END JWalton 2/8/2007
                End If
'            Else ' Reset everything
'
'
'                Set m_pArrowPt1 = Nothing
'                Set m_pArrowPt2 = Nothing
'                Set m_pArrowPt3 = Nothing
'                Set m_pArrowPt4 = Nothing
'                Set colArrowPoints = Nothing
'                m_bToolJustCompletedTask = True
'
'                ' Deactivate the tool
'                '++ START JWalton 2/8/2007 m_pApp to g_pApp
'                Set g_pApp.CurrentTool = Nothing
'                '++ END JWalton 2/8/2007
'
'            End If
        Case "DIMENSION"
            '++ START Added by Laura Gordon, 02/20/2007
                'Get the 3 input points
                If m_pArrowPt1 Is Nothing Then
                    Set m_pArrowPt1 = m_pPt
                ElseIf m_pArrowPt2 Is Nothing Then
                    Set m_pArrowPt2 = m_pPt
                    DrawArrows
                ElseIf m_pArrowPt3 Is Nothing Then
                    Set m_pArrowPt3 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                    DrawArrows
                    
                        'Check to be sure the 2 points are not equal
                        If (m_pArrowPt1.X = m_pArrowPt2.X) And (m_pArrowPt1.Y = m_pArrowPt2.Y) Then
                            Call MsgBox("Two input points can't be equal, dimension arrows terminated.", vbInformation, "Invalid Input")
                            '++ START Laura Gordon 03/05/07, Prevent error.
                            Call ITool_OnKeyDown(vbKeyQ, 1)
                            Exit Sub
                            'm_pEditor.AbortOperation
                            '++ END Laura Gordon 03/05/07
                        End If
                
                        'Get the mapscale
                        sCurrentMapScale = GetCurrentMapScale(pMIFC)
                        
                        'Get the side the dimension arrows should be on based upon the 3rd input point
                        sSide = GetDimensionArrowSide
                        
                        'Create two dimension arrows, one on the left and one on the right
                        'Dim lngIndex As Long
                        For lngIndex = 1 To 2
                        
                            Dim pDimensionLine As IPolyline4
                            Dim pDimensionPoints As IPointCollection4
                            Set pDimensionPoints = New Polyline
                            
                            'START Laura Gordon 05/21/07, manually add dimension arrow
                            If frmDimensionArrowSizes.chkAddManually = False Then
                            'END Laura Gordon 05/21/07
                            
                                'Add starting point
                                pDimensionPoints.AddPoint m_pArrowPt1
                                
                                'Instanciate the temp points
                                Set m_pArrowPtTemp = New Point
                                Set m_pArrowPtTemp2 = New Point
                                
                                'Set the distance of change for the dimension arrows based upon the mapscale
                                iChange = GetChange(sCurrentMapScale, Shift)
                                
                                'Get 3 calculated points for a hook
                                    'Create the line from input
                                    pDimensionPoints.AddPoint m_pArrowPt1
                                    pDimensionPoints.AddPoint m_pArrowPt2
                                    Set pDimensionLine = pDimensionPoints
                                
                                    '++ START Laura Gordon 03/08/07, get user sizes
                                    'Check the ratio changes
                                    If m_bolDimensionChanged = True Then
                                        If frmDimensionArrowSizes.RatioLine > 0 Then
                                            dRatioLine = frmDimensionArrowSizes.RatioLine
                                        Else
                                            dRatioLine = 1.75
                                        End If
                                        
                                        If frmDimensionArrowSizes.RatioCurve > 0 Then
                                            dRatioCurve = frmDimensionArrowSizes.RatioCurve
                                        Else
                                            dRatioCurve = 1.35
                                        End If
                                        
                                        '++ START Laura Gordon 05/21/07, adding user input smooth option
                                        If frmDimensionArrowSizes.SmoothRatio >= 0 Then
                                            dSmoothRatio = frmDimensionArrowSizes.SmoothRatio
                                        Else
                                            dSmoothRatio = 10
                                        End If
                                    Else
                                        dRatioLine = 1.75
                                        dRatioCurve = 1.35
                                        dSmoothRatio = 10
                                    End If
                                    '++ END Laura Gordon
                                    
                                    'Create a line iChange from the beginning
                                    Dim pLine As ILine
                                    Set pLine = New esriGeometry.Line
                                    If lngIndex = 1 Then
                                        If Shift = 0 Then 'shift not pressed
                                            If Trim$(UCase$(sSide)) = "RIGHT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, iChange, False, (iChange / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, iChange, False, ((iChange / dRatioLine) / 2), pLine
                                                '++ END Laura Gordon
                                            ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, iChange, False, -(iChange / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, iChange, False, -((iChange / dRatioLine) / 2), pLine
                                                '++ END Laura Gordon
                                            End If
                                        ElseIf Shift = 1 Then 'shift pressed
                                            If Trim$(UCase$(sSide)) = "RIGHT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, iChange, False, ((iChange * 2) / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, iChange, False, (((iChange * 2) / dRatioLine) / 2), pLine
                                                '++ END Laura Gordon
                                            ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, iChange, False, -((iChange * 2) / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, iChange, False, -(((iChange * 2) / dRatioLine) / 2), pLine
                                                '++ END Laura Gordon
                                            End If
                                        End If
                                    ElseIf lngIndex = 2 Then
                                        If Shift = 0 Then 'shift not pressed
                                            If Trim$(UCase$(sSide)) = "RIGHT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - iChange), False, (iChange / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - iChange), False, ((iChange / dRatioLine) / 2), pLine
                                                '++ END Laura Gordon
                                            ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - iChange), False, -(iChange / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - iChange), False, -((iChange / dRatioLine) / 2), pLine
                                                '++ END Laura Gordon
                                            End If
                                        ElseIf Shift = 1 Then 'shift pressed
                                            If Trim$(UCase$(sSide)) = "RIGHT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - iChange), False, ((iChange * 2) / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - iChange), False, (((iChange * 2) / dRatioLine) / 2), pLine
                                                '++ END Laura Gordon
                                            ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - iChange), False, -((iChange * 2) / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - iChange), False, -(((iChange * 2) / dRatioLine) / 2), pLine
                                                '++ END Laura Gordon
                                            End If
                                        End If
                                    End If
                                    
                                    'Save the to and from points of the line
                                    m_pArrowPtTemp.X = pLine.ToPoint.X
                                    m_pArrowPtTemp.Y = pLine.ToPoint.Y
                                    Set pLine = Nothing
                                    
                                    
                                    'Create a line (iChange/.25) from the beginning
                                    Set pLine = New esriGeometry.Line
                                    If lngIndex = 1 Then
                                        If Shift = 0 Then 'shift not pressed
                                            If Trim$(UCase$(sSide)) = "RIGHT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (iChange / dRatioCurve), False, ((iChange / dRatioLine) / 1.75), pLine
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (iChange / dRatioCurve), False, (iChange / dRatioLine), pLine
                                                
                                                '++ END Laura Gordon
                                            ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (iChange / dRatioCurve), False, -((iChange / dRatioLine) / 1.75), pLine
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (iChange / dRatioCurve), False, -(iChange / dRatioLine), pLine
                                                '++ END Laura Gordon
                                            End If
                                        ElseIf Shift = 1 Then 'shift pressed
                                            If Trim$(UCase$(sSide)) = "RIGHT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (iChange / dRatioCurve), False, ((iChange * 2) / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (iChange / dRatioCurve), False, (((iChange * 2) / dRatioLine) / 1.75), pLine
                                                '++ END Laura Gordon
                                            ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (iChange / dRatioCurve), False, -((iChange * 2) / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (iChange / dRatioCurve), False, -(((iChange * 2) / dRatioLine) / 1.75), pLine
                                                '++ END Laura Gordon
                                            End If
                                        End If
                                    ElseIf lngIndex = 2 Then
                                        If Shift = 0 Then 'shift not pressed
                                            If Trim$(UCase$(sSide)) = "RIGHT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, (iChange / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, ((iChange / dRatioLine) / 1.75), pLine
                                                '++ END Laura Gordon
                                            ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, -(iChange / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, -((iChange / dRatioLine) / 1.75), pLine
                                                '++ END Laura Gordon
                                            End If
                                        ElseIf Shift = 1 Then 'shift pressed
                                            If Trim$(UCase$(sSide)) = "RIGHT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, ((iChange * 2) / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, (((iChange * 2) / dRatioLine) / 1.75), pLine
                                                '++ END Laura Gordon
                                            ElseIf Trim$(UCase$(sSide)) = "LEFT" Then
                                                '++ START Laura Gordon 05/21/07, shape change per groups suggestions
                                                'pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, -((iChange * 2) / dRatioLine), pLine
                                                pDimensionLine.QueryNormal esriExtendAtFrom, (pDimensionLine.Length - (iChange / dRatioCurve)), False, -(((iChange * 2) / dRatioLine) / 1.75), pLine
                                                '++ END Laura Gordon
                                            End If
                                        End If
                                    End If
                                    
                                    'Save the to and from points of the line
                                    m_pArrowPtTemp2.X = pLine.ToPoint.X
                                    m_pArrowPtTemp2.Y = pLine.ToPoint.Y
                                    Set pLine = Nothing
                                    
                                    Set pDimensionPoints = Nothing
                                    Set pDimensionPoints = New Polyline
                                    
                                    'Add the points to the line/hook to be created
                                    If lngIndex = 1 Then
                                        pDimensionPoints.AddPoint m_pArrowPt1
                                    ElseIf lngIndex = 2 Then
                                        pDimensionPoints.AddPoint m_pArrowPt2
                                    End If
                                    pDimensionPoints.AddPoint m_pArrowPtTemp2
                                    pDimensionPoints.AddPoint m_pArrowPtTemp
                               
                               'START Laura Gordon 05/21/07, manually add dimension arrow
                               Else 'Adding dimension arrow manually
                                    lngIndex = 3
                                    pDimensionPoints.AddPoint m_pArrowPt1
                                    pDimensionPoints.AddPoint m_pArrowPt2
                                    pDimensionPoints.AddPoint m_pArrowPt3
                                    
                                    If frmDimensionArrowSizes.SmoothRatio >= 0 Then
                                        dSmoothRatio = frmDimensionArrowSizes.SmoothRatio
                                    Else
                                        dSmoothRatio = 10
                                    End If
                               End If
                               'END Laura Gordon 05/21/07
                               
                            'Create the dimension arrow from the collection of 3 points (1 input, 2 calculated)
                            Set pDimensionLine = pDimensionPoints
                
                
                            'START Laura Gordon 04/23/07, changed the direction fo the line, per Lane Co. request.
                            'Reverse the line to = from and from = to so arrow points the correct way
                            'pDimensionLine.ReverseOrientation
                            'END Laura Gordon 04/23/07
                        
                            
                            'START Laura Gordon 05/21/07, add user input smooth option
                            'pDimensionLine.Smooth pDimensionLine.Length / 10
                            
                            If dSmoothRatio > 0 Then
                                pDimensionLine.Smooth pDimensionLine.Length / dSmoothRatio
                            End If
                            'END Laura Gordon
                            
                            Dim pDimensionFeature As IFeature
                            Dim pDimensionWSEdit As IWorkspaceEdit
                            Dim pDimensionDSet As IDataset
                            Dim pDimensionArrowLayer As IFeatureLayer
                            Dim pDimensionArrowFC As IFeatureClass
                            Set pDimensionArrowLayer = FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
                            If pDimensionArrowLayer Is Nothing Then
                              MsgBox "The layer, " & g_pFldnames.FCCartoLines & ", is not in the map.", vbExclamation, "Layer not found"
                              m_pEditor.AbortOperation
                              Exit Sub
                            End If
                            Set pDimensionArrowFC = pDimensionArrowLayer.FeatureClass
                            Set pDimensionDSet = pDimensionArrowFC
                            Set pDimensionWSEdit = pDimensionDSet.Workspace
                            
                            'create the arrow feature
                            pDimensionWSEdit.StartEditOperation
                            Set pDimensionFeature = pDimensionArrowFC.CreateFeature
                            Set pDimensionFeature.Shape = pDimensionLine
                            lLineTypeFld = LocateFields(pDimensionArrowFC, g_pFldnames.CLinesLineTypeFN)
                            If lLineTypeFld = -1 Then Exit Sub
                            pDimensionFeature.Value(lLineTypeFld) = 134 'Bearing Distance Arrow
                            
                            'Get the current MapNumber
                            Dim sCurrentMapNums As String
                            '++ START Nick Seigal (LCOG) 01/25/2008 - Added sort-by field
                            sCurrentMapNums = GetValueViaOverlay(pDimensionFeature.Shape, pMIFC, g_pFldnames.MIMapNumberFN, g_pFldnames.MIMapNumberFN)
                            '++ END Nick Seigal (LCOG) 01/25/2008
                            Dim lCLMNSfld As Long
                            lCLMNSfld = LocateFields(pDimensionArrowFC, g_pFldnames.MIMapNumberFN)
                            pDimensionFeature.Value(lCLMNSfld) = sCurrentMapNums
                
                            '++ START Laura 03/08/07, calc additional fields
                            'Set the AutoMethod Field
                            lLineTypeFld = LocateFields(pDimensionArrowFC, g_pFldnames.AutoMethodFN)
                            If lLineTypeFld = -1 Then Exit Sub
                            pDimensionFeature.Value(lLineTypeFld) = "UNK"
                            
                            'Set the AutoWho Field
                            lLineTypeFld = LocateFields(pDimensionArrowFC, g_pFldnames.AutoWhoFN)
                            If lLineTypeFld = -1 Then Exit Sub
                            pDimensionFeature.Value(lLineTypeFld) = UserName
                            
                            'Set the AutoDate Field
                            lLineTypeFld = LocateFields(pDimensionArrowFC, g_pFldnames.AutoDateFN)
                            If lLineTypeFld = -1 Then Exit Sub
                            pDimensionFeature.Value(lLineTypeFld) = Format(Now, "MM/DD/YYYY")
                            
                            'Set the MapScale Field
                            lLineTypeFld = LocateFields(pDimensionArrowFC, g_pFldnames.MIMapScaleFN)
                            If lLineTypeFld = -1 Then Exit Sub
                            '++ START Nick Seigal (LCOG) 12/20/2007 - Fixed invalid value for field type error
                            pDimensionFeature.Value(lLineTypeFld) = Val(sCurrentMapScale) * 12
                            '++ END Nick Seigal (LCOG) 12/20/2007
                            '++  End Laura Gordon
                
                            pDimensionFeature.Store
                            pDimensionWSEdit.StopEditOperation
                            
                            '++ START Nick Seigal (LCOG) 12/20/2007 - Better refresh
                            'Partially refresh the display
                            pActiveView.PartialRefresh _
                                    esriViewBackground Or _
                                    esriViewGeography Or _
                                    esriViewGraphics Or _
                                    esriViewGraphicSelection, Nothing, pDimensionFeature.Extent.Envelope
                            '++ END Nick Seigal (LCOG) 12/20/2007
                            Set pDimensionLine = Nothing
                            Set m_pArrowPtTemp = Nothing
                            Set m_pArrowPtTemp2 = Nothing
                            
                            Set m_pTextSymbol = Nothing
                            Set m_pTextPoint = Nothing
                            Set m_pLinePolyline = Nothing
                            Set m_pLineSymbol = Nothing
                            
                        Next
                        
                        Set m_pArrowPt1 = Nothing
                        Set m_pArrowPt2 = Nothing
                        Set m_pArrowPt3 = Nothing
                        Set m_pArrowPt4 = Nothing
                    
        Else 'Reset everything
            Set m_pArrowPt1 = Nothing
            Set m_pArrowPt2 = Nothing
            Set m_pArrowPt3 = Nothing
            Set m_pArrowPt4 = Nothing
            Set m_pArrowPtTemp = Nothing
            Set m_pArrowPtTemp2 = Nothing
            Set m_pSnapAgent = Nothing
            m_bToolJustCompletedTask = True
            
            'Deactivate the tool
            Set g_pApp.CurrentTool = Nothing
            
        End If
        '++ END Added by Laura Gordon, 02/20/2007
    End Select
  
Process_Exit:
    Exit Sub
    
ErrorHandler:
    HandleError True, _
                "ITool_OnMouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  ITool_OnMouseMove
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Handle the OnMouseMove Event for the tool
'Called From:   Class Object
'Description:   Draws a temporary line if the arrow type is a hook using
'               the original mouse click coordinates and the current
'               mouse coordinates.
'Methods:       None
'Inputs:        Button, Shift, X, Y (System)
'Parameters:    None
'Outputs:       m_bDoOnce - Initialization flag
'               m_pLinePolyline - The temporary polyline
'               m_pLineSymbol - Line Symbol of the temporary line
'               m_pTextSymbol - Text symbol of the temporary text
'               m_pTextPoint - Point for temporary text
'               m_pToBreakPoint - The current mouse position
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point created
'James Moore    11-1-06     commented out dead variables and assignments
'Laura Gordon   02/20/2007  added point for dimension arrows
'***************************************************************************

Private Sub ITool_OnMouseMove( _
  ByVal Button As Long, _
  ByVal Shift As Long, _
  ByVal X As Long, _
  ByVal Y As Long)
On Error GoTo ErrorHandler
    '++ START JWalton 2/9/2007 Centralized Variable Declarations
    ' Variable Declarations
    Dim pMxDoc As esriArcMapUI.IMxDocument
    Dim pActiveView As esriCarto.IActiveView
    Dim pRGBColor As esriDisplay.IRgbColor
    Dim pSymbol As esriDisplay.ISymbol
    Dim pCPoint As esriGeometry.IConstructPoint
    Dim pFCPoint As esriGeometry.IConstructPoint
    Dim pCv As esriGeometry.ICurve
    Dim pLine As esriGeometry.ILine
    Dim pPoint As esriGeometry.IPoint
    Dim pPLine As esriGeometry.IPolyline
    Dim pPolyline As esriGeometry.IPolyline
    Dim pSegColl As esriGeometry.ISegmentCollection
    Dim bfirstTime As Boolean
    Dim angle As Double
    Dim deltaX As Double
    Dim deltaY As Double
    Dim lLineLength As Long
    Dim pFont As stdOle.IFontDisp
    '++ END JWalton 2/9/2007
    
    '++ START JWalton 2/9/2007
        ' If Statement conditional change from Form control to Form property
    '++ END JWalton 2/9/2007
    
    ' Draw the temporary line the user sees while moving the mouse
    If frmArrows.ArrowType = "Hook" Then
        If (Not m_bInUse) Then GoTo Process_Exit
        
        '++ START JWalton 2/8/2007 m_pApp to g_pApp
        Set pMxDoc = g_pApp.Document
        '++ END JWalton 2/8/2007
        Set pActiveView = pMxDoc.FocusMap
        
        ' Checks to see if the line symbol is defined
        If (m_pLineSymbol Is Nothing) Then bfirstTime = True
        
        ' Get current point
        Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
        Set m_pToBreakPoint = pPoint 'the unextended to point used to break the hooks
        
        '++ START Added by Laura Gordon, 02/20/2007
        'Check to be sure there is a start point for the hook to prevent an error
        If m_pStartPoint Is Nothing Then Exit Sub
        '++ END Added by Laura Gordon, 02/20/2007
        
        ' Draw a virtual line that represents the extended line
        Set pPLine = New esriGeometry.Polyline
        pPLine.FromPoint = m_pStartPoint
        pPLine.ToPoint = pPoint
        Set pCv = pPLine
                      
        Set pCPoint = New esriGeometry.Point
        pCPoint.ConstructAlong pCv, esriExtendAtTo, pCv.Length + CDbl(lLineLength), False
        Set pPoint = pCPoint
         
         If Not m_bDoOnce Then
              Set pFCPoint = New esriGeometry.Point
              pFCPoint.ConstructAlong pCv, esriExtendAtFrom, -(CDbl(lLineLength)), False
              Set m_pStartPoint = pFCPoint
              m_bDoOnce = True
         End If
        
        'Draw the line
        pActiveView.ScreenDisplay.StartDrawing pActiveView.ScreenDisplay.hdc, -1
          
        ' Initialize or draw the temporary line
        If bfirstTime Then
            'Line Symbol
            Set m_pLineSymbol = New esriDisplay.SimpleLineSymbol
            m_pLineSymbol.Width = 2
            Set pRGBColor = New esriDisplay.RgbColor
            With pRGBColor
                .Red = 223
                .Green = 223
                .Blue = 223
            End With
            m_pLineSymbol.Color = pRGBColor
            Set pSymbol = m_pLineSymbol
            pSymbol.ROP2 = esriDisplay.esriRasterOpCode.esriROPXOrPen
              
            'Text Symbol
            Set m_pTextSymbol = New esriDisplay.TextSymbol
            m_pTextSymbol.HorizontalAlignment = esriDisplay.esriTextHorizontalAlignment.esriTHACenter
            m_pTextSymbol.VerticalAlignment = esriDisplay.esriTextVerticalAlignment.esriTVACenter
            m_pTextSymbol.Size = 16
            Set pSymbol = m_pTextSymbol
            Set pFont = m_pTextSymbol.Font
            pFont.Name = "Arial"
            pSymbol.ROP2 = esriDisplay.esriRasterOpCode.esriROPXOrPen
            
            'Create point to draw text in
            Set m_pTextPoint = New esriGeometry.Point
          Else
            'Use existing symbols and draw existing text and polyline
            pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
            pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
            pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
            If (m_pLinePolyline.Length > 0) Then _
              pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
        End If
        
        'Get line between from and to points, and angle for text
        Set pLine = New esriGeometry.Line
        pLine.PutCoords m_pStartPoint, pPoint
        angle = pLine.angle
        angle = angle * (180# / 3.14159)
        If ((angle > 90#) And (angle < 180#)) Then
            angle = angle + 180#
          ElseIf ((angle < 0#) And (angle < -90#)) Then
            angle = angle - 180#
          ElseIf ((angle < -90#) And (angle > -180)) Then
            angle = angle - 180#
          ElseIf (angle > 180) Then
            angle = angle - 180#
        End If
        
        
        'For drawing text, get text(distance), angle, and point
        'jwm dead variable Dim distance As Double
        deltaX = pPoint.X - m_pStartPoint.X
        deltaY = pPoint.Y - m_pStartPoint.Y
        m_pTextPoint.X = m_pStartPoint.X + deltaX / 2#
        m_pTextPoint.Y = m_pStartPoint.Y + deltaY / 2#
        m_pTextSymbol.angle = angle
        'jwm   distance = Round(Sqr((deltaX * deltaX) + (deltaY * deltaY)), 3)
        m_pTextSymbol.Text = "" '"[" & distance & "]" distance commented out by unknown
        
        'Draw text
        pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
        pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
        
        'Get polyline with blank space for text
        Set pPolyline = New esriGeometry.Polyline
        Set pSegColl = pPolyline
        pSegColl.AddSegment pLine
        Set m_pLinePolyline = GetSmashedLine(pActiveView.ScreenDisplay, m_pTextSymbol, m_pTextPoint, pPolyline)
        
        'Draw polyline
        pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
        If (m_pLinePolyline.Length > 0) Then _
          pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
        pActiveView.ScreenDisplay.FinishDrawing
    End If
    
    '++ START Added by Laura Gordon, 02/20/2007
    If frmArrows.ArrowType = "Dimension" Then
    
        'Set variables
        Dim pId As New esriSystem.UID
        pId = "esriEditor.Editor"
        Set m_pEditor = g_pApp.FindExtensionByCLSID(pId)
        Set m_pEditorEvents = m_pEditor
        
        Set m_pMxApp = g_pApp
        'Set m_pDisp = m_pMxApp.Display
        Set m_pDoc = g_pApp.Document
        Set m_pMxDoc = m_pDoc
        Set m_pActive = m_pMxDoc.ActiveView
        
        'Set m_pMapEvents = m_pMxDoc.FocusMap
    
    
      If m_bToolJustCompletedTask Then 'gets rid of a stray editor agent
        m_pActive.Refresh
        m_bToolJustCompletedTask = False
        Exit Sub
      End If
      
      If m_bMouseHasMoved Then
        'Check to be sure m_pPt has a value, prevents an error if called after other tools
        If m_pPt Is Nothing Then
          m_bMouseHasMoved = False
          Exit Sub
        End If
        'erase the old agent
        m_pEditor.InvertAgent m_pPt, 0
        'get the new point
        GetMousePoint X, Y
      Else
        Set m_pPt = m_pActive.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
      End If
      m_bMouseHasMoved = True
    
      If Not m_pSnapAgent Is Nothing Then
        Dim pSnapenv As ISnapEnvironment
        Set pSnapenv = m_pEditor
        Dim pTmpPt As IPoint
        Set pTmpPt = New Point
        pTmpPt.PutCoords m_pPt.X, m_pPt.Y
        Dim dTol As Double
        dTol = pSnapenv.SnapTolerance
        m_bInTol = pSnapenv.SnapPoint(m_pPt)
        If Not m_bInTol Then
          Set m_pPt = New Point
          m_pPt.PutCoords pTmpPt.X, pTmpPt.Y  '+++ set the point back b/c it was not in the snap tol
        End If
      Else
        m_bInTol = False
      End If
      
      'If m_lCtr = 0 Then
        m_pEditor.InvertAgent m_pPt, 0 'draw the agent
        Exit Sub
      'End If
      Exit Sub
    
    End If
    '++ END Added by Laura Gordon, 02/20/2007

Process_Exit:
    Exit Sub
    
ErrorHandler:
    HandleError True, _
                "ITool_OnMouseMove " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  ITool_OnMouseUp
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Handle the tool's OnMouseUp Event
'Called From:   Class object
'Description:   Draws and finalizes any hook features, but only draws
'               temporary features for anything else.
'Methods:       None
'Inputs:        Button, Shift, X, Y (System)
'Parameters:    None
'Outputs:       m_bInUse - Boolean flag that indicates whether or not the
'                          tool is currently in use
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'Laura Gordon   02/20/2007  Added if statement back in so only hooks do the following
'***************************************************************************

Private Sub ITool_OnMouseUp( _
  ByVal Button As Long, _
  ByVal Shift As Long, _
  ByVal X As Long, _
  ByVal Y As Long)
On Error GoTo ErrorHandler
    '++ START JWalton 2/9/2007 Centralized Variable Declarations
    ' Variable Declarations
    Dim pMxDoc As esriArcMapUI.IMxDocument
    Dim pActiveView As esriCarto.IActiveView
    '++ END JWalton 2/9/2007
    
    '++ START JWalton 2/9/2007
        ' Removed If...Then Block (If frmArrows.lblDisplay.Caption="hook")
        ' Inner block code is now obigatory
    '++ END JWalton 2/9/2007
    
    '++ START Laura Gordon 02/20/2007 'Adding if block back in
    If frmArrows.ArrowType = "Hook" Then
    '++ END Laura Gordon 02/20/2007
    
        If (Not m_bInUse) Then GoTo Process_Exit
            
        If (m_pLineSymbol Is Nothing) Then GoTo Process_Exit
        
        '++ START JWalton 2/9/2007 m_pApp to g_pApp
    '    Set pMXDoc = m_pApp.Document Commented out by jwm on behalf of JWalton 2-20-07
        Set pMxDoc = g_pApp.Document
        '++ END JWalton 2/9/2007
        Set pActiveView = pMxDoc.FocusMap
        
        ' Draws a temporary line on the screen
        pActiveView.ScreenDisplay.StartDrawing pActiveView.ScreenDisplay.hdc, -1
        pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
        pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
        pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
        If (m_pLinePolyline.Length > 0) Then pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
        pActiveView.ScreenDisplay.FinishDrawing
        
        '++ START JWalton 2/9/2007 Form control to Form property
        If frmArrows.ArrowType = "Hook" Then
            ' Generate hooks based on the graphic polyline
            GenerateHooks m_pLinePolyline
          Else
            ' Do Nothing / Store nothing
        End If
        '++ END JWalton 2/9/2007
        
        ' Records that the tool is no longer in use
        m_bInUse = False
        
        ' Clean up
        Set m_pTextSymbol = Nothing
        Set m_pTextPoint = Nothing
        Set m_pLinePolyline = Nothing
        Set m_pLineSymbol = Nothing
    '++ START Laura Gordon 02/20/2007 ending if statement
    End If
    '++ END Laura Gordon 02/20/2007

Process_Exit:
    Exit Sub
    
ErrorHandler:
    HandleError True, _
                "ITool_OnMouseUp " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

Private Sub ITool_Refresh( _
  ByVal hdc As esriSystem.OLE_HANDLE)
    ' Not Implemented
End Sub

'***************************************************************************
'Name:                  GetSmashedLine
'Initial Author:        <<Unknown>>
'Subsequent Author:     Type your name here.
'Created:               <<Unknown>>
'Purpose:       Find a polyline that intersects a text element but omits
'               the section that intersects the text element
'Called From:   cmdArrows.ITool_OnMouseMove
'Description:   Give a screen display, pDisplay, a text symbol, pTextSymbol,
'               a point to fill the a polygon with, pPoint, and the polyline
'               to intersect, pPolyline.
'               Find the bounding polygon of pTextSymbol and fill it with
'               pPoint.  Then determine the intersection between pPolyline
'               and pTextSymbol.  Finally, return the difference of the
'               intersection and pPolyline.
'Methods:       None
'Inputs:        pDisplay - The current display
'               pTextSymbol - The text to overlay the polyline
'               pPoint - A filler symbol for the returned polyline
'               pPolyline - The polyline to intersect
'Parameters:    None
'Outputs:       None
'Returns:       A polyline that represents the difference between the
'               input polyline and the intersection of the input polyline
'               and the boundary of the input text symbol
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    11-1-06     Removed dead variable
'***************************************************************************

Private Function GetSmashedLine( _
  pDisplay As esriDisplay.IScreenDisplay, _
  pTextSymbol As esriDisplay.ISymbol, _
  pPoint As esriGeometry.IPoint, _
  pPolyline As esriGeometry.IPolyline) As esriGeometry.IPolyline
On Error GoTo ErrorHandler
    '++ START JWalton 2/9/2007 Centralized Variable Declarations
    ' Variable Declarations
    Dim pBoundary As esriGeometry.IPolygon
    Dim pIntersect As esriGeometry.IPolyline
    Dim pTopo As esriGeometry.ITopologicalOperator
    '++ END JWalton 2/9/2007
    
    ' Query the text symbol for its bounding polygon
    Set pBoundary = New esriGeometry.Polygon
    pTextSymbol.QueryBoundary pDisplay.hdc, pDisplay.DisplayTransformation, pPoint, pBoundary
    
    ' Returns the intersection of the polyline and the text symbol
    Set pTopo = pBoundary
    Set pIntersect = pTopo.Intersect(pPolyline, esriGeometry1Dimension)
    
    ' Returns the difference between the polyline and the intersection
    Set pTopo = pPolyline
    Set GetSmashedLine = pTopo.Difference(pIntersect)

    Exit Function
ErrorHandler:
    HandleError False, _
                "GetSmashedLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  GenerateHooks
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Generate landhooks
'Called From:   cmdArrows.ITool_OnMouseUp
'Description:   Given a base line, pSketch.
'               Create a land hook based on the endpoints of pSketch.
'Methods:       None
'Inputs:        pSketch - An object that supports the IGeometry interface
'Parameters:    None
'Outputs:       m_dHookAngle - The angle of the landhooks in relation to the
'                              base line.
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Created single exit point
'James Moore    11-1-06     Removed dead variables
'John Walton    2/9/2007    Removed variables bOk
'Laura Gordon   3/8/2007    Added required field calc (automethod, autowho, autodate, mapnumber, mapscale)
'***************************************************************************

Private Sub GenerateHooks( _
  ByRef pSketch As esriGeometry.IGeometry)
On Error GoTo ErrorHandler
    '++ START JWalton 2/9/2007 Centralized Variable Declarations
    ' Variable Declarations
    Dim pDoc As esriArcMapUI.IMxDocument
    Dim pMxDoc As esriArcMapUI.IMxDocument
    Dim pActiveView As esriCarto.IActiveView
    Dim pHookLayer As esriCarto.IFeatureLayer
    Dim pMIFlayer As esriCarto.IFeatureLayer
    Dim pApp As esriFramework.IApplication
    Dim pDSet As esriGeoDatabase.IDataset
    Dim pFeature As esriGeoDatabase.IFeature
    Dim pHookFC As esriGeoDatabase.IFeatureClass
    Dim pMIFclass As esriGeoDatabase.IFeatureClass
    Dim pWSEdit As esriGeoDatabase.IWorkspaceEdit
    Dim pCurve As esriGeometry.ICurve
    Dim pEnumSeg As esriGeometry.IEnumSegment
    Dim pGeomColl As esriGeometry.IGeometryCollection
    Dim pNormal As esriGeometry.ILine
    Dim pNewPointColl As esriGeometry.IPointCollection
    Dim pPointColl As esriGeometry.IPointCollection
    Dim pWholeLine As esriGeometry.IPolyline4
    Dim pSeg As esriGeometry.ISegment
    Dim pPath1 As esriGeometry.ISegmentCollection
    Dim pPath2 As esriGeometry.ISegmentCollection
    Dim pPath3 As esriGeometry.ISegmentCollection
    Dim pSegCollection As esriGeometry.ISegmentCollection
    Dim bBool As Boolean
    Dim bCreateDefaultHook As Boolean
    Dim bSplitHappened As Boolean
    Dim dHookLength As Double
    Dim dSideA As Double
    Dim dSideB As Double
    Dim dSideC As Double
    Dim lLineLength As Long
    Dim lLineTypeFld As Long
    Dim lNewPartIndex As Long
    Dim lNewSegIndex As Long
    Dim lPartIndex As Long
    Dim lSegIndex As Long
    Dim vMapScale1 As Variant
    Dim vMapScale2 As Variant
   '++ END JWalton 2/9/2007
   '++  START Laura Gordon 3/8/07, New variable for map number
   Dim sCurMapNum As String
   '++  END Laura Gordon
    
    ' Initialize the hook angle
    m_dHookAngle = 20
    
    'Make sure the edit sketch is a polyline
    If Not TypeOf pSketch Is IPolyline Then GoTo Process_Exit
    
    Set pCurve = pSketch
    
    ' Retrieve a reference to the Map Index layer
    Set pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then
        MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
        GoTo Process_Exit
    End If
    Set pMIFclass = pMIFlayer.FeatureClass
    
    ' Retrieve the map scale from the overlaying Map Index layer
    '++ START Nick Seigal (LCOG) 01/25/2008 - Added sort-by field
    vMapScale1 = GetValueViaOverlay(pCurve.FromPoint, pMIFclass, g_pFldnames.MIMapScaleFN, g_pFldnames.MIMapNumberFN)
    vMapScale2 = GetValueViaOverlay(pCurve.ToPoint, pMIFclass, g_pFldnames.MIMapScaleFN, g_pFldnames.MIMapNumberFN)
    '++ END Nick Seigal (LCOG) 01/25/2008
    
    ' Insure that the map scales exist and that they are equal
    If IsNull(vMapScale1) Or IsNull(vMapScale2) Then
        MsgBox "No mapscale for current MapIndex.  Unable to create hooks", vbOKOnly
        GoTo Process_Exit
    End If
    If vMapScale1 <> vMapScale2 Then
        MsgBox "Hook can not span Mapindex polygons with different scale", vbCritical
        GoTo Process_Exit
    End If
    
    ' Insures that the map scale is supported -- Not all scales are defined (Issue)
    If vMapScale1 = 600 Then
        lLineLength = 20
      ElseIf vMapScale1 = 1200 Then
        lLineLength = 20
      ElseIf vMapScale1 = 2400 Then
        lLineLength = 40
      ElseIf vMapScale1 = 4800 Then
        lLineLength = 80
      ElseIf vMapScale1 = 24000 Then
        lLineLength = 400
      Else
        MsgBox "Not a valid mapscale.  Unable to create hooks", vbOKOnly
        GoTo Process_Exit
    End If
    dHookLength = lLineLength * 0.1
    
    '++ START JWalton 2/9/2007 Removed m_pEditor.StopOperation (No operation was started)
    ' Insures that the polyline only has two vertices (Starting & Ending points only)
    If Not IsSketcha2PointLine(pSketch) Then
      GoTo Process_Exit
    End If
    '++ END JWalton 2/9/2007
    
    '++ START JWalton 2/9/2007 m_pApp to g_pApp
    Set pApp = g_pApp 'm_pEditor.Parent
    '++ END JWalton 2/9/2007
    Set pDoc = pApp.Document
    
    'Get the hook layer
    Set pHookLayer = FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
    If pHookLayer Is Nothing Then
        '++ START JWalton 2/9/2007
        ' m_pDSName to g_pFldNames.FCCartoLines
        MsgBox "The layer, " & g_pFldnames.FCCartoLines & ", is not in the map.", _
               vbExclamation, _
               "Layer not found"
               
        ' Removed m_pEditor.StopOperation (No operation was started)
        '++ END JWalton 2/9/2007
        GoTo Process_Exit
    End If
    Set pHookFC = pHookLayer.FeatureClass
    Set pDSet = pHookFC
    Set pWSEdit = pDSet.Workspace
    
    ' Locate the line type field
    lLineTypeFld = LocateFields(pHookFC, g_pFldnames.CLinesLineTypeFN)
    If lLineTypeFld = -1 Then GoTo Process_Exit
    
    ' Initialize line objects and collections to create a new line
    Set pNewPointColl = New esriGeometry.Polyline
    Set pNormal = New esriGeometry.Line
    Set pPointColl = pSketch
      
    ' Adds the head of the hook based on the specified angle and hook length
    dSideA = (dHookLength * Sin((360 - m_dHookAngle) * (3.14 / 180)))
    dSideC = dHookLength
    dSideB = Sqr((dSideC * dSideC) - (dSideA * dSideA))
    pCurve.QueryNormal esriNoExtension, _
                       dSideB, _
                       False, _
                       dSideA, _
                       pNormal
    pNewPointColl.AddPoint pNormal.ToPoint
    
    ' Adds the line points
    pNewPointColl.AddPoint pPointColl.Point(0)
    pNewPointColl.AddPoint pPointColl.Point(1)
    
    ' Adds the tail of the hook based on the specified angle and hook length
    dSideA = (dHookLength * Sin(m_dHookAngle * (3.14 / 180)))
    dSideC = dHookLength
    dSideB = Sqr((dSideC * dSideC) - (dSideA * dSideA))
    pCurve.QueryNormal esriGeometry.esriSegmentExtension.esriNoExtension, _
                      (pCurve.Length - dSideB), _
                      False, _
                      dSideA, _
                      pNormal
    pNewPointColl.AddPoint pNormal.ToPoint
    
    'Now get rid of the line between the start and end points (where user clicked)
    Set pWholeLine = pNewPointColl
    pWholeLine.SplitAtPoint m_pFromBreakPoint, True, bBool, bSplitHappened, lNewPartIndex, lNewSegIndex
    pWholeLine.SplitAtPoint m_pToBreakPoint, True, bBool, bSplitHappened, lNewPartIndex, lNewSegIndex
    
    ' Initialize new path objects and collections to create a new polyline
    Set pPath1 = New esriGeometry.Path
    Set pPath2 = New esriGeometry.Path
    Set pPath3 = New esriGeometry.Path
    
    ' QI to get the segment collection of the landhook
    Set pSegCollection = pWholeLine
    
    ' Retreive an enumeration of the segments
    Set pEnumSeg = pSegCollection.EnumSegments
    
    ' Add segments to the paths that will make the final land hook
    pEnumSeg.Next pSeg, lPartIndex, lSegIndex
    '++ START JWalton 2/9/2007 While...Wend to Do While...Loop
    Do While Not pSeg Is Nothing
      If lSegIndex < 1 Then
          pPath1.AddSegment pSeg
      ElseIf lSegIndex = 1 Then
          pPath2.AddSegment pSeg
      ElseIf lSegIndex = 2 Then
          pPath3.AddSegment pSeg
      End If
      pEnumSeg.Next pSeg, lPartIndex, lSegIndex
    Loop
    '++ END JWalton 2/9/2007
    
    ' Add the component paths to the final land hook
    Set pGeomColl = New esriGeometry.Polyline
    pGeomColl.AddGeometry pPath1
    pGeomColl.AddGeometry pPath2
    pGeomColl.AddGeometry pPath3
    pGeomColl.GeometriesChanged
    
    'Store the new land hook feature
    pWSEdit.StartEditOperation
    Set pFeature = pHookFC.CreateFeature
    Set pFeature.Shape = pGeomColl
    pFeature.Value(lLineTypeFld) = 101
    
    '++  START Laura Gordon 03/08/2007, Calc additional fields
    'Set the AutoMethod Field
    lLineTypeFld = LocateFields(pHookFC, g_pFldnames.AutoMethodFN)
    If lLineTypeFld = -1 Then Exit Sub
    pFeature.Value(lLineTypeFld) = "UNK"
    
    'Set the AutoWho Field
    lLineTypeFld = LocateFields(pHookFC, g_pFldnames.AutoWhoFN)
    If lLineTypeFld = -1 Then Exit Sub
    pFeature.Value(lLineTypeFld) = UserName
    
    'Set the AutoDate Field
    lLineTypeFld = LocateFields(pHookFC, g_pFldnames.AutoDateFN)
    If lLineTypeFld = -1 Then Exit Sub
    pFeature.Value(lLineTypeFld) = Format(Now, "MM/DD/YYYY")
    
    'Set the MapScale Field
    lLineTypeFld = LocateFields(pHookFC, g_pFldnames.MIMapScaleFN)
    If lLineTypeFld = -1 Then Exit Sub
    '++ START Nick Seigal (LCOG) 12/20/2007 - Fixed data type
    pFeature.Value(lLineTypeFld) = Val(vMapScale1)
    '++ END Nick Seigal (LCOG) 12/20/2007
    
    'Set the MapNumber Field
    '++ START Nick Seigal (LCOG) 01/25/2008 - Added sort-by field
    sCurMapNum = GetValueViaOverlay(pFeature.Shape, pMIFclass, g_pFldnames.MIMapNumberFN, g_pFldnames.MIMapNumberFN)
    '++ END Nick Seigal (LCOG) 01/25/2008)
    lLineTypeFld = LocateFields(pHookFC, g_pFldnames.MIMapNumberFN)
    If lLineTypeFld = -1 Then Exit Sub
    pFeature.Value(lLineTypeFld) = sCurMapNum
    '++  END Laura Gordon
    
    
    pFeature.Store
    pWSEdit.StopEditOperation
    
    'Refresh the display
    '++ START JWalton 2/9/2007 m_pApp to g_pApp
    Set pMxDoc = g_pApp.Document
    '++ END JWalton 2/9/2007
    Set pActiveView = pMxDoc.FocusMap
    '++ START Nick Seigal (LCOG) 12/20/2007 - Better refresh
    'Partially refresh the display
    pActiveView.PartialRefresh esriViewBackground Or _
            esriViewGeography Or _
            esriViewGraphics Or _
            esriViewGraphicSelection, Nothing, pFeature.Extent.Envelope
    '++ END Nick Seigal (LCOG) 12/20/2007

Process_Exit:
    Exit Sub
    
ErrorHandler:
    HandleError False, _
                "GenerateHooks " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  IsSketcha2PointLine
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Determine if a geometry is a two-point line
'Called From:   cmdArrows.GenerateHooks
'Description:   Given a geometry object, pGeom.
'               Evaluate pGeom the determine if it is a point collection that
'               has exactly two points
'Methods:       None
'Inputs:        pGeom - The geometry to evaluate
'Parameters:    None
'Outputs:       None
'Returns:       A boolean value that indicates if the geometry is a either a
'               two-point line or not
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/9/2007    Replaced ESRI Error Handler in favor of
'                           returning False in the event of an error
'***************************************************************************

Public Function IsSketcha2PointLine( _
  pGeom As esriGeometry.IGeometry) As Boolean
On Error GoTo ErrorHandler
    ' Variable declarations
    Dim pPointColl As esriGeometry.IPointCollection
    
    ' Validate the passed geometry
    If Not TypeOf pGeom Is esriGeometry.IPointCollection Then GoTo Process_Exit
    
    ' Validate the number of points in the collection
    Set pPointColl = pGeom
    If pPointColl.PointCount <> 2 Then
        MsgBox "When creating the parcel hook only digitize two points", vbOKOnly Or vbExclamation, "Parcel Hook Error"
        GoTo Process_Exit
    End If
    
    ' Returns the value of the function
    IsSketcha2PointLine = True

Process_Exit:
    Exit Function
    
ErrorHandler:
    '++ START JWalton 2/9/2007
    ' Return False in the event of an error
    IsSketcha2PointLine = False
    '++ END JWalton 2/9/2007
End Function

'***************************************************************************
'Name:                  ReturnExtended
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Find the tangent of a curve at a point
'Called From:   This function is not called from anywhere
'Description:   Given a polyline, pPolyline, a length to extend the
'               polyline, lLength, and an extension method, pExt.
'               Extend pPolyline according to the the pExt method to a
'               length of lLength.
'Methods:       None
'Inputs:        pExt - esriSegmentExtension constant
'               pPolyline - Polyline to extend
'               lLength - Length to extend the polyline by
'Parameters:    None
'Outputs:       None
'Returns:       An polyline
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    11-1-06     removed dead variables
'***************************************************************************

Private Function ReturnExtended( _
  pExt As esriGeometry.esriSegmentExtension, _
  pPolyline As esriGeometry.IPolyline, _
  lLength As Long) As esriGeometry.IPolyline
On Error GoTo ErrorHandler
    '++ START JWalton 2/9/2007 Centralized Variable Declarations
    ' Variable Declarations
    Dim pCurve As esriGeometry.ICurve
    Dim pLine As esriGeometry.ILine
    Dim pPLine As esriGeometry.IPolyline
    '++ END JWalton 2/9/2007
    
    ' Need to extend the end (creates an ILine object)
    Set pCurve = pPolyline
    Set pLine = New esriGeometry.Line
    pCurve.QueryTangent pExt, _
                        1, _
                        False, _
                        lLength, _
                        pLine
    
    ' Convert ILine to an IPolyline
    Set pPLine = New esriGeometry.Polyline
    pPLine.FromPoint = pLine.FromPoint
    pPLine.ToPoint = pLine.ToPoint
    Set ReturnExtended = pPLine


    Exit Function
ErrorHandler:
    HandleError False, _
                "ReturnExtended " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  DrawArrows
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Create a temporary polyline
'Called From:   cmdArrows.ITool_OnMouseDown
'Description:   Given four points, m_pArrowPt1 thru m_pArrowPt4.
'               Create a temporary polyline from the given points and
'               display it as a temporary line on the display
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
'James Moore    11-1-06     removed dead variable
'Laura Gordon   02/20/2007  change m_pApp to g_pApp as JWalton hadn't changed it.
'Laura Gordon   05/23/07    adding additional arrow types
'***************************************************************************

Private Sub DrawArrows()
On Error GoTo ErrorHandler
    '++ START JWalton 2/9/2007 Centralized Variable Declarations
    ' Variable Declarations
    Dim pMxDoc As esriArcMapUI.IMxDocument
    Dim pActiveView As esriCarto.IActiveView
    Dim pRGBColor As esriDisplay.IRgbColor
    Dim pSymbol As esriDisplay.ISymbol
    Dim pArrowLine As esriGeometry.IPointCollection4
    Dim pArrowLine2 As esriGeometry.IPolyline4
    '++ END JWalton 2/9/2007

    ' Initialize references to the current document and map
    '++ START Laura Gordon 02/20/2007
    Set pMxDoc = g_pApp.Document 'm_pApp.Document
    '++ STOP Laura Gordon 02/20/2007
    Set pActiveView = pMxDoc.FocusMap
    
    '++ START Laura Gordon 05/23/07, adding additional arrow types
    Dim lngIndex As Long
    '++ END Laura Gordon
    
    'Set up line symbol to display temporary line
    Set m_pLineSymbol = New esriDisplay.SimpleLineSymbol
    m_pLineSymbol.Width = 2
    Set pRGBColor = New esriDisplay.RgbColor
    With pRGBColor
      .Red = 223
      .Green = 223
      .Blue = 223
    End With
    m_pLineSymbol.Color = pRGBColor
    Set pSymbol = m_pLineSymbol
    pSymbol.ROP2 = esriDisplay.esriRasterOpCode.esriROPXOrPen
    
    ' Create the polyline from a point collection
    Set pArrowLine = New esriGeometry.Polyline
    
    '++ START Laura Gordon 05/23/07, adding additional arrow types
    If UCase$(frmArrows.ArrowType) = "ARROW" Then
        If colArrowPoints.Count > 1 Then
            For lngIndex = 1 To colArrowPoints.Count
                pArrowLine.AddPoint colArrowPoints.Item(lngIndex)
            Next
        End If
    Else
        If Not m_pArrowPt1 Is Nothing Then pArrowLine.AddPoint m_pArrowPt1
        If Not m_pArrowPt2 Is Nothing Then pArrowLine.AddPoint m_pArrowPt2
        If Not m_pArrowPt3 Is Nothing Then pArrowLine.AddPoint m_pArrowPt3
        If Not m_pArrowPt4 Is Nothing Then pArrowLine.AddPoint m_pArrowPt4
    End If
    
'    If Not m_pArrowPt1 Is Nothing Then pArrowLine.AddPoint m_pArrowPt1
'    If Not m_pArrowPt2 Is Nothing Then pArrowLine.AddPoint m_pArrowPt2
'    If Not m_pArrowPt3 Is Nothing Then pArrowLine.AddPoint m_pArrowPt3
'    If Not m_pArrowPt4 Is Nothing Then pArrowLine.AddPoint m_pArrowPt4
    '++ END Laura Gordon 05/23/07
    
    Set pArrowLine2 = pArrowLine
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''Make lines, then add to segment collection
    '''''''''see mouse move event
    ''''''''''''''''''''''''''''''''''''
        
    ' Draw the temporary line
    pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
    If (pArrowLine2.Length > 0) Then pActiveView.ScreenDisplay.DrawPolyline pArrowLine2
    pActiveView.ScreenDisplay.FinishDrawing

    Exit Sub
ErrorHandler:
    HandleError False, _
                "DrawArrows " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'++ START Added by Laura Gordon, 02/20/2007
'***************************************************************************
'Name:                  GetMousePoint
'Initial Author:        Laura Gordon
'Subsequent Author:     <<Type your name here>>
'Created:               02/20/2007
'Purpose:       Get the mousepoint using ISnapAgent
'Called From:   ITool_OnMouseDown
'Description:   Get the mousepoint using ISnapAgent
'Methods:       None
'Inputs:        None
'Parameters:    X As Long, Y As Long
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************
Private Sub GetMousePoint(X As Long, Y As Long)

On Error GoTo ErrorHandler

  '+++ Get the current map point (and invert the agent at that location)
  Set m_pPt = m_pActive.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  
  '+++ get the snap agent, if it is being used
  Set m_pSnapAgent = Nothing
  Dim pSnapAgent As ISnapAgent
  Dim pSnapenv As ISnapEnvironment
  Set pSnapenv = m_pEditor
  Dim i As Long
  For i = 0 To pSnapenv.SnapAgentCount - 1
    Set pSnapAgent = pSnapenv.SnapAgent(i)
    If TypeOf pSnapAgent Is IFeatureSnapAgent Then
        Dim pFSnapAgent As IFeatureSnapAgent
        Dim pEdLyrs As IEditLayers
        Dim pLayer As ILayer
        Set pFSnapAgent = pSnapAgent
        Set pEdLyrs = m_pEditor
        Set pLayer = pEdLyrs.CurrentLayer
        Dim ht As esriGeometryHitPartType
        ht = pFSnapAgent.HitType
        If ht <> 0 Then
            Set m_pSnapAgent = pFSnapAgent
            Exit For
        End If
        Set pLayer = Nothing
        Set pEdLyrs = Nothing
        Set pFSnapAgent = Nothing
    End If
  Next i
  
  Set pSnapAgent = Nothing
  Set pSnapenv = Nothing
Exit Sub

ErrorHandler:
    HandleError False, "ReturnExtended " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
    Set pLayer = Nothing
    Set pEdLyrs = Nothing
    Set pFSnapAgent = Nothing
    Set pSnapAgent = Nothing
    Set pSnapenv = Nothing
End Sub
'***************************************************************************
'Name:                  GetCurrentMapScale
'Initial Author:        Laura Gordon
'Subsequent Author:     <<Type your name here>>
'Created:               02/20/2007
'Purpose:       Get the mapscale of a temp feature
'Called From:   ITool_OnMouseDown
'Description:   Get the mapscale of a temp feature
'Methods:       None
'Inputs:        None
'Parameters:    pMIFC As IFeatureClass
'Outputs:       None
'Returns:       String
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************
Private Function GetCurrentMapScale(pMIFC As IFeatureClass) As String

On Error GoTo ErrorHandler

      Dim pDimensionFeatureTemp As IFeature
      Dim pDimensionWSEditTemp As IWorkspaceEdit
      Dim pDimensionDSetTemp As IDataset
      Dim pDimensionArrowLayerTemp As IFeatureLayer
      Dim pDimensionArrowFCTemp As IFeatureClass

      Set pDimensionArrowLayerTemp = FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
      If pDimensionArrowLayerTemp Is Nothing Then
        MsgBox "The layer, " & g_pFldnames.FCCartoLines & ", is not in the map.", vbExclamation, "Layer not found"
        m_pEditor.AbortOperation
        Exit Function
      End If
      Set pDimensionArrowFCTemp = pDimensionArrowLayerTemp.FeatureClass
      Set pDimensionDSetTemp = pDimensionArrowFCTemp
      Set pDimensionWSEditTemp = pDimensionDSetTemp.Workspace

      'create the arrow feature
      pDimensionWSEditTemp.StartEditOperation
      Set pDimensionFeatureTemp = pDimensionArrowFCTemp.CreateFeature
      Dim pDimensionpointsTemp As IPointCollection4
      Set pDimensionpointsTemp = New Polyline
      pDimensionpointsTemp.AddPoint m_pArrowPt1
      pDimensionpointsTemp.AddPoint m_pArrowPt2
    
      Dim pDimensionLineTemp As IPolyline
      Set pDimensionLineTemp = pDimensionpointsTemp

      Set pDimensionFeatureTemp.Shape = pDimensionLineTemp
    
      
      'Get the current MapScale
      Dim sCurrentMapScale As String
      '++ START Nick Seigal (LCOG) 01/25/2008 - Added sort-by field
      GetCurrentMapScale = GetValueViaOverlay(pDimensionFeatureTemp.Shape, pMIFC, g_pFldnames.MIMapScaleFN, g_pFldnames.MIMapNumberFN)
      '++ END Nick Seigal (LCOG) 01/25/2008
      GetCurrentMapScale = GetCurrentMapScale / 12
      pDimensionWSEditTemp.AbortEditOperation
      pDimensionWSEditTemp.StopEditOperation
              
      Set pDimensionpointsTemp = Nothing
      Set pDimensionLineTemp = Nothing
      Set pDimensionArrowFCTemp = Nothing
      Set pDimensionArrowLayerTemp = Nothing
      Set pDimensionDSetTemp = Nothing
      Set pDimensionWSEditTemp = Nothing
      Set pDimensionFeatureTemp = Nothing
Exit Function

ErrorHandler:
  HandleError False, "ReturnExtended " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
  Set pDimensionpointsTemp = Nothing
  Set pDimensionLineTemp = Nothing
  Set pDimensionArrowFCTemp = Nothing
  Set pDimensionArrowLayerTemp = Nothing
  Set pDimensionDSetTemp = Nothing
  Set pDimensionWSEditTemp = Nothing
  Set pDimensionFeatureTemp = Nothing
End Function

'***************************************************************************
'Name:                  GetDimensionArrowSide
'Initial Author:        Laura Gordon
'Subsequent Author:     <<Type your name here>>
'Created:               02/20/2007
'Purpose:       Get the side of the line the dimension arrows will be placed on; right or left
'Called From:   ITool_OnMouseDown
'Description:   Get the side of the line the dimension arrows will be placed on; right or left
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       String
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************
Private Function GetDimensionArrowSide() As String

On Error GoTo ErrorHandler

    'Determine point location is on the left or right by Dean Anderson, help of Nate Anderson
    Dim r1              As String
    Dim slope           As Double
    Dim yint            As Double
    Dim mess            As String
    Dim z               As Double

    slope = 0
    If m_pArrowPt2.Y <> m_pArrowPt1.Y Then
        slope = (m_pArrowPt1.Y - m_pArrowPt2.Y) / (m_pArrowPt1.X - m_pArrowPt2.X)
    End If
    
    yint = m_pArrowPt1.Y - (slope * m_pArrowPt1.X)
    
    z = (slope * m_pArrowPt3.X) + yint - m_pArrowPt3.Y
    
    If m_pArrowPt1.X = m_pArrowPt2.X Then           'vertical
        If m_pArrowPt3.X = m_pArrowPt1.X Then z = 0
        If m_pArrowPt1.Y < m_pArrowPt2.Y Then       'going up
            If m_pArrowPt3.X > m_pArrowPt1.X Then z = 1
            If m_pArrowPt3.X < m_pArrowPt1.X Then z = -1
        End If
        If m_pArrowPt1.Y > m_pArrowPt2.Y Then       'going down
            If m_pArrowPt3.X < m_pArrowPt1.X Then z = 1
            If m_pArrowPt3.X > m_pArrowPt1.X Then z = -1
        End If
    End If
    If m_pArrowPt1.Y = m_pArrowPt2.Y Then           'horizontal
        If m_pArrowPt3.Y = m_pArrowPt1.Y Then z = 0
        If m_pArrowPt1.X < m_pArrowPt2.X Then       'going right
            If m_pArrowPt3.Y > m_pArrowPt1.Y Then z = -1
            If m_pArrowPt3.Y < m_pArrowPt1.Y Then z = 1
        End If
        If m_pArrowPt1.X > m_pArrowPt2.X Then       'going left
            If m_pArrowPt3.Y < m_pArrowPt1.Y Then z = -1
            If m_pArrowPt3.Y > m_pArrowPt1.Y Then z = 1
        End If
    End If
    
    If z < 0 Then GetDimensionArrowSide = "left"
    If z > 0 Then GetDimensionArrowSide = "right"
    If z = 0 Then GetDimensionArrowSide = "left" '"online"
    
    If (m_pArrowPt1.X > m_pArrowPt2.X And m_pArrowPt1.Y > m_pArrowPt2.Y) Or (m_pArrowPt1.X > m_pArrowPt2.X And m_pArrowPt2.Y > m_pArrowPt1.Y) Then
     If z > 0 Then GetDimensionArrowSide = "left"
     If z < 0 Then GetDimensionArrowSide = "right"
     If z = 0 Then GetDimensionArrowSide = "left" '"online"
    End If

Exit Function

ErrorHandler:
  HandleError False, "ReturnExtended " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:                  GetChange
'Initial Author:        Laura Gordon
'Subsequent Author:     <<Type your name here>>
'Created:               02/20/2007
'Purpose:       Get the size of the dimension arrows so that all are standard
'Called From:   ITool_OnMouseDown
'Description:   Get the size of the dimension arrows so that all are standard right or left
'Methods:       None
'Inputs:        None
'Parameters:    sCurrentMapScale, ByVal Shift As Long
'Outputs:       None
'Returns:       Integer
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************
Private Function GetChange(ByVal sCurrentMapScale, ByVal Shift As Long) As Integer

On Error GoTo ErrorHandler

    If sCurrentMapScale = "100" Then
        GetChange = 15
    ElseIf sCurrentMapScale = "200" Then
        GetChange = 30
    ElseIf sCurrentMapScale = "400" Then
        GetChange = 60
    ElseIf sCurrentMapScale = "2000" Then
        GetChange = 300
    End If
    
Exit Function

ErrorHandler:
  HandleError False, "ReturnExtended " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
'++ END Added by Laura Gordon, 02/20/2007

