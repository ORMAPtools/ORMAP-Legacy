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
' SCC Revision number: $Revision: 35 $
' Date of last change: $Date: 2006-11-15 12:20:32 -0800 (Wed, 15 Nov 2006) $
'
'
' File name:            cmdArrows
'
' Initial Author:       Type your name here
'
' Date Created:
'
' Description: CLASS MODULE FOR THE HOOKS/ARROWS COMMAND
'
' Entry points:
'       List the public variables and their purposes.
'       List the properties and routines that the module exposes to the rest of the program.
'
' Dependencies: OPENS frmArrows
'       How does this file depend or relate to other files?
'
' Issues: Cannot use Option Explicit when compiling this class because of reference to m_pEditor object
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

'Option Explicit
'interfaces
Implements ICommand
Implements ITool
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

'------------------------------
' Private Variables
'------------------------------
Private m_pApp As esriFramework.IApplication
Private m_pDoc As esriArcMapUI.IMxDocument
Private m_pMap As IMap
Private m_dHookAngle As Double
Private m_bDoOnce As Boolean  'testing
Private m_bInUse As Boolean
Private m_pLineSymbol As ILineSymbol
Private m_pLinePolyline As IPolyline
Private m_pTextSymbol As ITextSymbol
Private m_pStartPoint As IPoint 'hooks
Private m_pTextPoint As IPoint
Private m_pFromBreakPoint As IPoint 'hooks
Private m_pToBreakPoint As IPoint 'hooks
Private m_pArrowPt1 As IPoint
Private m_pArrowPt2 As IPoint
Private m_pArrowPt3 As IPoint
Private m_pArrowPt4 As IPoint

'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "clsArrows.cls"

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------
 

Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler

    
    ICommand_Enabled = True


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Checked() As Boolean
  On Error GoTo ErrorHandler

    
    ICommand_Checked = False


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler

    
    ICommand_Name = "AddArrowsAndHooks"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler


    ICommand_Caption = "Add Arrows/Hooks"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler


    ICommand_Tooltip = "Open a dialog to add arrows and hooks"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Message() As String
  On Error GoTo ErrorHandler

    
    ICommand_Message = "Initialize the Add Arrows and Hooks form"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpFile() As String
  On Error GoTo ErrorHandler


    ICommand_HelpFile = ""


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_HelpContextID() As Long
  On Error GoTo ErrorHandler


    ICommand_HelpContextID = ""


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpContextID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here
    ' ICommand_Bitmap =


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Property Get ICommand_Category() As String
  On Error GoTo ErrorHandler


    ICommand_Category = "ORMAP"


  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 
Private Sub ICommand_OnCreate(ByVal hook As Object)
  On Error GoTo ErrorHandler

    
  Set m_pApp = hook
  If g_pApp Is Nothing Then
    Set g_pApp = hook
  End If
  Set m_pDoc = m_pApp.Document
  Set m_pMap = m_pDoc.FocusMap

  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
'***************************************************************************
'Name:  ICommand_OnClick
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:       10/11/2006
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

    'Get the MapIndex feature layer and fclass
    Dim pCLFLayer As IFeatureLayer
    Dim pTaxlotFlayer As IFeatureLayer
    Dim pMIFlayer As IFeatureLayer
    
    Set pTaxlotFlayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If pTaxlotFlayer Is Nothing Then
        MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        GoTo Process_Exit
    End If
    
    Set pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then
        MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
        GoTo Process_Exit
    End If
    
    Set pTaxlotFlayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If pTaxlotFlayer Is Nothing Then
        MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        GoTo Process_Exit
    End If
    
    Set pCLFLayer = FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
    If pCLFLayer Is Nothing Then
        MsgBox "Unable to locate Cartographic Lines layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCCartoLines
        GoTo Process_Exit
    End If
    'Open the form
    frmArrows.Show vbModal

Process_Exit:
    Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
  On Error GoTo ErrorHandler
    
    ' TODO: Add your implementation here

  Exit Property
ErrorHandler:
  HandleError True, "ITool_Cursor " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
 


Private Sub ITool_OnDblClick()
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here


  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnDblClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnKeyDown(ByVal keyCode As Long, ByVal Shift As Long)
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here


  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnKeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Sub ITool_OnKeyUp(ByVal keyCode As Long, ByVal Shift As Long)
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here


  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnKeyUp " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
 
Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean
  On Error GoTo ErrorHandler

    
    ' TODO: Add your implementation here


  Exit Function
ErrorHandler:
  HandleError True, "ITool_OnContextMenu " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function
 

 
Private Function ITool_Deactivate() As Boolean
  On Error GoTo ErrorHandler

  ' stop doing operation
  Set m_pTextSymbol = Nothing
  Set m_pTextPoint = Nothing
  Set m_pLinePolyline = Nothing
  Set m_pLineSymbol = Nothing
  m_bInUse = False

  ITool_Deactivate = True


  Exit Function
ErrorHandler:
  HandleError True, "ITool_Deactivate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  ITool_OnMouseDown
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point
'JWM            10/16/2006  Using case instead of if

'***************************************************************************
Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    On Error GoTo ErrorHandler
    m_bInUse = True
    Dim pMXDoc As IMxDocument
    Dim pActiveView As IActiveView
    Set pMXDoc = m_pApp.Document
    Set pActiveView = pMXDoc.FocusMap
  
    'Get reference to MapIndex FC
    Dim sToolCaption As String
    Dim sScale As String
    Dim pMIFC As IFeatureClass
    Dim pMIFL As IFeatureLayer
    Set pMIFL = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    Set pMIFC = pMIFL.FeatureClass
    
  sToolCaption = UCase$(frmArrows.lblCurrentTool.Caption)
  
    Select Case sToolCaption
        Case "HOOK"  'If drawing hooks
            'Get point to measure distance from
            Set m_pStartPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
            m_bDoOnce = False 'testing
            Set m_pFromBreakPoint = m_pStartPoint
            'Get the scale of the current mapindex
            sScale = GetValueViaOverlay(m_pStartPoint, pMIFC, g_pFldnames.MIMapScaleFN)
            sScale = ConvertToDescription(pMIFC.Fields, g_pFldnames.MIMapScaleFN, sScale)
        Case "ARROW" 'If drawing annotation arrows
            If m_pArrowPt1 Is Nothing Then
                Set m_pArrowPt1 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
            ElseIf m_pArrowPt2 Is Nothing Then
                Set m_pArrowPt2 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                DrawArrows
            ElseIf m_pArrowPt3 Is Nothing Then
                Set m_pArrowPt3 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                DrawArrows
            ElseIf m_pArrowPt4 Is Nothing Then
                Set m_pArrowPt4 = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                DrawArrows
                Dim pArrowLine As IPolyline4
                Dim pArrowPoints As IPointCollection4
                Set pArrowPoints = New Polyline
                pArrowPoints.AddPoint m_pArrowPt1
                pArrowPoints.AddPoint m_pArrowPt2
                pArrowPoints.AddPoint m_pArrowPt3
                pArrowPoints.AddPoint m_pArrowPt4
                Set pArrowLine = pArrowPoints
                pArrowLine.Smooth pArrowLine.Length / 10
        
                Dim pFeature As IFeature
                Dim pWSEdit As IWorkspaceEdit
                Dim pDSet As IDataset
                Dim pArrowLayer As IFeatureLayer
                Dim pArrowFC As IFeatureClass
                Set pArrowLayer = FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
                If pArrowLayer Is Nothing Then
                    MsgBox "The layer, " & g_pFldnames.FCCartoLines & ", is not in the map.", vbExclamation, "Layer not found"
                    m_pEditor.AbortOperation
                    GoTo Process_Exit
                End If
                Set pArrowFC = pArrowLayer.FeatureClass
                Set pDSet = pArrowFC
                Set pWSEdit = pDSet.Workspace
                
                'create the arrow feature
                pWSEdit.StartEditOperation
                Set pFeature = pArrowFC.CreateFeature
                Set pFeature.Shape = pArrowLine
                lLineTypeFld = LocateFields(pArrowFC, g_pFldnames.CLinesLineTypeFN)
                
                If lLineTypeFld = -1 Then GoTo Process_Exit
                
                pFeature.Value(lLineTypeFld) = 100
                
                'Get the current MapNumber
                Dim sCurMapNum As String
                sCurMapNum = GetValueViaOverlay(pFeature.Shape, pMIFC, g_pFldnames.MIMapNumberFN)
                Dim lCLMNfld As Long
                lCLMNfld = LocateFields(pArrowFC, g_pFldnames.MIMapNumberFN)
                pFeature.Value(lCLMNfld) = sCurMapNum
                
                pFeature.Store
                pWSEdit.StopEditOperation
                
                'Refresh the display
                pActiveView.PartialRefresh esriViewBackground, Nothing, pFeature.Extent.Envelope
                
                Set pArrowLine = Nothing
                Set m_pArrowPt1 = Nothing
                Set m_pArrowPt2 = Nothing
                Set m_pArrowPt3 = Nothing
                Set m_pArrowPt4 = Nothing
                
                Set m_pTextSymbol = Nothing
                Set m_pTextPoint = Nothing
                Set m_pLinePolyline = Nothing
                Set m_pLineSymbol = Nothing
                'Deactivate the tool
                Set m_pApp.CurrentTool = Nothing
            Else 'Reset everything
                Set m_pArrowPt1 = Nothing
                Set m_pArrowPt2 = Nothing
                Set m_pArrowPt3 = Nothing
                Set m_pArrowPt4 = Nothing
                'Deactivate the tool
                Set m_pApp.CurrentTool = Nothing
            End If
        Case "DIMENSION"
    End Select
  
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  ITool_OnMouseMove
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point created
'James Moore    11-1-06     commented out dead variables and assignments
'***************************************************************************
Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler

'Draw the temporary line the user sees while moving the mouse
If frmArrows.lblCurrentTool.Caption = "hook" Then
  If (Not m_bInUse) Then GoTo Process_Exit
  
  Dim pMXDoc As IMxDocument
  Dim pActiveView As IActiveView
  Set pMXDoc = m_pApp.Document
  Set pActiveView = pMXDoc.FocusMap
  
  Dim bfirstTime As Boolean
  If (m_pLineSymbol Is Nothing) Then bfirstTime = True

  'Get current point
  Dim pPoint As IPoint
  Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  Set m_pToBreakPoint = pPoint 'the unextended to point used to break the hooks
  
  
   'Draw a virtual line that represents the extended line
   Dim pPline As IPolyline
   Set pPline = New Polyline
   pPline.FromPoint = m_pStartPoint
   pPline.ToPoint = pPoint
   Dim pCv As ICurve
   Set pCv = pPline
   
   Dim pCPoint As IConstructPoint
   Set pCPoint = New Point
   pCPoint.ConstructAlong pCv, esriExtendAtTo, pCv.Length + CDbl(lLineLength), False
   Set pPoint = pCPoint
   
   If Not m_bDoOnce Then
        Dim pFCPoint As IConstructPoint
        Set pFCPoint = New Point
        pFCPoint.ConstructAlong pCv, esriExtendAtFrom, -(CDbl(lLineLength)), False
        Set m_pStartPoint = pFCPoint
        m_bDoOnce = True
   End If

  'Draw the line
  pActiveView.ScreenDisplay.StartDrawing pActiveView.ScreenDisplay.hdc, -1
    
  If bfirstTime Then
    Dim pRGBColor As IRgbColor
    Dim pSymbol As ISymbol
    Dim pFont As IFontDisp
    
    'Line Symbol
    Set m_pLineSymbol = New SimpleLineSymbol
    m_pLineSymbol.Width = 2
    Set pRGBColor = New RgbColor
    With pRGBColor
      .Red = 223
      .Green = 223
      .Blue = 223
    End With
    m_pLineSymbol.Color = pRGBColor
    Set pSymbol = m_pLineSymbol
    pSymbol.ROP2 = esriROPXOrPen
    
    'Text Symbol
    Set m_pTextSymbol = New TextSymbol
    m_pTextSymbol.HorizontalAlignment = esriTHACenter
    m_pTextSymbol.VerticalAlignment = esriTVACenter
    m_pTextSymbol.Size = 16
    Set pSymbol = m_pTextSymbol
    Set pFont = m_pTextSymbol.Font
    pFont.Name = "Arial"
    pSymbol.ROP2 = esriROPXOrPen
    
    'Create point to draw text in
    Set m_pTextPoint = New Point
    
  Else
    'Use existing symbols and draw existing text and polyline
    pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
    pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
    pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
    If (m_pLinePolyline.Length > 0) Then _
      pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
  End If

  'Get line between from and to points, and angle for text
  Dim pLine As ILine
  Set pLine = New esrigeometry.Line
  pLine.PutCoords m_pStartPoint, pPoint
  Dim angle As Double
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
  Dim deltaX As Double
  Dim deltaY As Double
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
  Dim pPolyline As IPolyline
  Set pPolyline = New Polyline
  Dim pSegColl As ISegmentCollection
  Set pSegColl = pPolyline
  pSegColl.AddSegment pLine
  Set m_pLinePolyline = GetSmashedLine(pActiveView.ScreenDisplay, m_pTextSymbol, m_pTextPoint, pPolyline)
  
  'Draw polyline
  pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
  If (m_pLinePolyline.Length > 0) Then _
    pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
  
  pActiveView.ScreenDisplay.FinishDrawing

End If
Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseMove " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  ITool_OnMouseUp
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler

  If frmArrows.lblCurrentTool.Caption = "hook" Then
    If (Not m_bInUse) Then GoTo Process_Exit
    m_bInUse = False
    
    If (m_pLineSymbol Is Nothing) Then GoTo Process_Exit
    
    Dim pMXDoc As IMxDocument
    Dim pActiveView As IActiveView
    Set pMXDoc = m_pApp.Document
    Set pActiveView = pMXDoc.FocusMap
    
    'Draw temp line
    pActiveView.ScreenDisplay.StartDrawing pActiveView.ScreenDisplay.hdc, -1
    pActiveView.ScreenDisplay.SetSymbol m_pTextSymbol
    pActiveView.ScreenDisplay.DrawText m_pTextPoint, m_pTextSymbol.Text
    pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
    If (m_pLinePolyline.Length > 0) Then pActiveView.ScreenDisplay.DrawPolyline m_pLinePolyline
    pActiveView.ScreenDisplay.FinishDrawing
    If frmArrows.lblCurrentTool.Caption = "hook" Then
    'Generate hooks based on the graphic polyline
      GenerateHooks m_pLinePolyline
    Else
      'Nothing yet
    End If
    
    Set m_pTextSymbol = Nothing
    Set m_pTextPoint = Nothing
    Set m_pLinePolyline = Nothing
    Set m_pLineSymbol = Nothing
  End If

Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseUp " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ITool_Refresh(ByVal hdc As esriSystem.OLE_HANDLE)
  On Error GoTo ErrorHandler



  Exit Sub
ErrorHandler:
  HandleError True, "ITool_Refresh " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  GetSmashedLine
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    11-1-06     removed dead variable
'***************************************************************************
Private Function GetSmashedLine(pDisplay As IScreenDisplay, pTextSymbol As ISymbol, pPoint As IPoint, pPolyline As IPolyline) As IPolyline
  On Error GoTo ErrorHandler

  'Returns a graphic Polyline
  Dim pBoundary As IPolygon
  Set pBoundary = New Polygon
  pTextSymbol.QueryBoundary pDisplay.hdc, pDisplay.DisplayTransformation, pPoint, pBoundary
  Dim pTopo As ITopologicalOperator
  Set pTopo = pBoundary
  
  Dim pIntersect As IPolyline
  Set pIntersect = pTopo.Intersect(pPolyline, esriGeometry1Dimension)
  Set pTopo = pPolyline
  Set GetSmashedLine = pTopo.Difference(pIntersect)


  Exit Function
ErrorHandler:
  HandleError False, "GetSmashedLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  GenerateHooks
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  created single exit point
'James Moore    11-1-06     removed dead variables
'***************************************************************************
Private Sub GenerateHooks(ByRef pSketch As esrigeometry.IGeometry)
  On Error GoTo ErrorHandler

m_dHookAngle = 20
'Make sure the edit sketch is a polyline
If Not TypeOf pSketch Is IPolyline Then GoTo Process_Exit

Dim pCurve As ICurve
Set pCurve = pSketch

Dim dHookLength As Double

'Get the map scale so hook length can be determined
Dim pMIFlayer As IFeatureLayer
Dim pMIFclass As IFeatureClass
Set pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
If pMIFlayer Is Nothing Then
    MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
    "This process requires a feature class called " & g_pFldnames.FCMapIndex
    GoTo Process_Exit
End If
Set pMIFclass = pMIFlayer.FeatureClass
Dim vMapScale1 As Variant
Dim vMapScale2 As Variant
vMapScale1 = GetValueViaOverlay(pCurve.FromPoint, pMIFclass, g_pFldnames.MIMapScaleFN)
vMapScale2 = GetValueViaOverlay(pCurve.ToPoint, pMIFclass, g_pFldnames.MIMapScaleFN)
If IsNull(vMapScale1) Or IsNull(vMapScale2) Then
    MsgBox "No mapscale for current MapIndex.  Unable to create hooks", vbOKOnly
    GoTo Process_Exit
End If
If vMapScale1 <> vMapScale2 Then
    MsgBox "Hook can not span Mapindex polygons with differenc scale", vbCritical
    GoTo Process_Exit
End If
'ISSUE:  NOT ALL SCALES DEFINED
Dim lLineLength As Long
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
    
'dHookLength = (m_dHookLength * pCurve.Length)

Dim bOK As Boolean
'Make sure the polyline only has two vertices (start and end point only)
bOK = IsSketcha2PointLine(pSketch)

If bOK = False Then
  m_pEditor.AbortOperation
  GoTo Process_Exit
End If

Dim pApp As esriFramework.IApplication
Set pApp = m_pApp 'm_pEditor.Parent
Dim pDoc As esriArcMapUI.IMxDocument
Set pDoc = pApp.Document


'Get the hook layer
Dim lLineTypeFld As Long
Dim pWSEdit As IWorkspaceEdit
Dim pDSet As IDataset
Dim pHookLayer As IFeatureLayer
Dim pHookFC As IFeatureClass
Set pHookLayer = FindFeatureLayerByDS(g_pFldnames.FCCartoLines)
If pHookLayer Is Nothing Then
    MsgBox "The layer, " & m_sHookDSName & ", is not in the map.", vbExclamation, "Layer not found"
    m_pEditor.AbortOperation
    GoTo Process_Exit
End If
Set pHookFC = pHookLayer.FeatureClass
Set pDSet = pHookFC
Set pWSEdit = pDSet.Workspace
lLineTypeFld = LocateFields(pHookFC, g_pFldnames.CLinesLineTypeFN)
If lLineTypeFld = -1 Then GoTo Process_Exit
'Create the new hook feature
Dim pFeature As IFeature
Dim pPointColl As esrigeometry.IPointCollection
Dim bCreateDefaultHook As Boolean

bCreateDefaultHook = True

Set pPointColl = pSketch

Dim dSideA As Double
Dim dSideB As Double
Dim dSideC As Double
Dim pNormal As ILine

'Create a long line that is essentially what the user drew
'Create a hook at each end of the line
'Create hooks based on specified angle and hook length (based on scale)
If bCreateDefaultHook = True Then 'Always will be true here
  Dim pNewPointColl As IPointCollection
  Set pNewPointColl = New Polyline
    
  dSideA = (dHookLength * Sin((360 - m_dHookAngle) * (3.14 / 180)))
  dSideC = dHookLength
  
  dSideB = (dSideC ^ 2) - (dSideA ^ 2)
  dSideB = Sqr(dSideB)
  
  Set pNormal = New esrigeometry.Line
  pCurve.QueryNormal esriNoExtension, dSideB, False, dSideA, pNormal
  
  pNewPointColl.AddPoint pNormal.ToPoint
  pNewPointColl.AddPoint pPointColl.Point(0)
  pNewPointColl.AddPoint pPointColl.Point(1)
      
  dSideA = (dHookLength * Sin(m_dHookAngle * (3.14 / 180)))
  dSideC = dHookLength
  
  dSideB = (dSideC ^ 2) - (dSideA ^ 2)
  dSideB = Sqr(dSideB)
  
  pCurve.QueryNormal esriNoExtension, (pCurve.Length - dSideB), False, dSideA, pNormal

  pNewPointColl.AddPoint pNormal.ToPoint
  
  'Now get rid of the line between the start and end points (where user clicked)
  Dim pGeomColl As IGeometryCollection
  Dim pPath1 As ISegmentCollection
  Dim pPath2 As ISegmentCollection
  Dim pPath3 As ISegmentCollection
  Dim pSegCollection As ISegmentCollection
  Dim pWholeLine As IPolyline4
  Dim lNewPartIndex As Long
  Dim lNewSegIndex As Long
  Dim bSplitHappened As Boolean
  Dim bBool As Boolean
  Dim pEnumSeg As IEnumSegment
  Dim pSeg As ISegment
'jwm dead variable  Dim pOutseg As ISegment
  Dim lPartIndex As Long
  Dim lSegIndex As Long
  Set pWholeLine = pNewPointColl
  pWholeLine.SplitAtPoint m_pFromBreakPoint, True, bBool, bSplitHappened, lNewPartIndex, lNewSegIndex
  pWholeLine.SplitAtPoint m_pToBreakPoint, True, bBool, bSplitHappened, lNewPartIndex, lNewSegIndex
  
  Set pPath1 = New Path
  Set pPath2 = New Path
  Set pPath3 = New Path
  Set pSegCollection = pWholeLine

  Set pEnumSeg = pSegCollection.EnumSegments
  pEnumSeg.Next pSeg, lPartIndex, lSegIndex
  While Not pSeg Is Nothing
    If lSegIndex < 1 Then
        pPath1.AddSegment pSeg
    ElseIf lSegIndex = 1 Then
        pPath2.AddSegment pSeg
    ElseIf lSegIndex = 2 Then
        pPath3.AddSegment pSeg
    End If
    pEnumSeg.Next pSeg, lPartIndex, lSegIndex
  Wend
  Set pGeomColl = New Polyline
  pGeomColl.AddGeometry pPath1
  pGeomColl.AddGeometry pPath2
  pGeomColl.AddGeometry pPath3
  pGeomColl.GeometriesChanged
  

  'Store the new hook feature
  pWSEdit.StartEditOperation
  Set pFeature = pHookFC.CreateFeature
  Set pFeature.Shape = pGeomColl 'pWholeLine 'pFinalPolyLine 'pNewPointColl
  pFeature.Value(lLineTypeFld) = 101
  pFeature.Store
  pWSEdit.StopEditOperation
  
  'Refresh the display
  Dim pMXDoc As IMxDocument
  Dim pActiveView As IActiveView
  Set pMXDoc = m_pApp.Document
  Set pActiveView = pMXDoc.FocusMap
  pActiveView.PartialRefresh esriViewBackground, Nothing, pFeature.Extent.Envelope
  
End If

Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError False, "GenerateHooks " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'***************************************************************************
'Name:  IsSketcha2PointLine
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************
Public Function IsSketcha2PointLine(pGeom As IGeometry) As Boolean
  On Error GoTo ErrorHandler


Dim pPointColl As IPointCollection
  
If Not TypeOf pGeom Is IPointCollection Then GoTo Process_Exit

Set pPointColl = pGeom
If pPointColl.PointCount <> 2 Then
  MsgBox "When creating the parcel hook only digitize two points", vbExclamation, "Parcel Hook Error"
  GoTo Process_Exit
End If

IsSketcha2PointLine = True

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, "IsSketcha2PointLine " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  ReturnExtended
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    11-1-06     removed dead variables
'***************************************************************************
Private Function ReturnExtended(pExt As esriSegmentExtension, pPolyline As IPolyline, lLength As Long) As IPolyline
  On Error GoTo ErrorHandler

    Dim pCurve As ICurve
    Dim pLine As ILine
    Dim pPline As IPolyline
    ' Need to extend the end (creates an ILine object)
    Set pCurve = pPolyline
    Set pLine = New esrigeometry.Line
    pCurve.QueryTangent pExt, 1, False, _
                        lLength, pLine
    ' Convert ILine to an IPolyline
    Set pPline = New Polyline
    pPline.FromPoint = pLine.FromPoint
    pPline.ToPoint = pLine.ToPoint
    Set ReturnExtended = pPline


  Exit Function
ErrorHandler:
  HandleError False, "ReturnExtended " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'***************************************************************************
'Name:  DrawArrows
'Initial Author:
'Subsequent Author:     Type your name here.
'Created:
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    11-1-06     removed dead variable
'***************************************************************************
Private Sub DrawArrows()
  On Error GoTo ErrorHandler

    Dim pMXDoc As IMxDocument
    Dim pActiveView As IActiveView
    Set pMXDoc = m_pApp.Document
    Set pActiveView = pMXDoc.FocusMap
    
    'Set up line symbol to display temporary line
    Dim pRGBColor As IRgbColor
    Dim pSymbol As ISymbol
    
    'Line Symbol
    Set m_pLineSymbol = New SimpleLineSymbol
    m_pLineSymbol.Width = 2
    Set pRGBColor = New RgbColor
    With pRGBColor
      .Red = 223
      .Green = 223
      .Blue = 223
    End With
    m_pLineSymbol.Color = pRGBColor
    Set pSymbol = m_pLineSymbol
    pSymbol.ROP2 = esriROPXOrPen
    
    Dim pArrowLine As IPointCollection4
    Dim pArrowLine2 As IPolyline4
    Set pArrowLine = New Polyline
    If Not m_pArrowPt1 Is Nothing Then pArrowLine.AddPoint m_pArrowPt1
    If Not m_pArrowPt2 Is Nothing Then pArrowLine.AddPoint m_pArrowPt2
    If Not m_pArrowPt3 Is Nothing Then pArrowLine.AddPoint m_pArrowPt3
    If Not m_pArrowPt4 Is Nothing Then pArrowLine.AddPoint m_pArrowPt4
    Set pArrowLine2 = pArrowLine
    'pArrowLine.Smooth pArrowLine.Length / 10
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''Make lines, then add to segment collection
    '''''''''see mouse move event
    ''''''''''''''''''''''''''''''''''''
        
    pActiveView.ScreenDisplay.SetSymbol m_pLineSymbol
    If (pArrowLine2.Length > 0) Then _
        pActiveView.ScreenDisplay.DrawPolyline pArrowLine2
  
    pActiveView.ScreenDisplay.FinishDrawing

    'Set m_pLinePolyline = GetSmashedLine(pActiveView.ScreenDisplay, m_pTextSymbol, m_pTextPoint, pPolyline)

  Exit Sub
ErrorHandler:
  HandleError False, "DrawArrows " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
