VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdTaxlotAssignment"
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
'
' File name:            cmdTaxlotAssignment
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
'           esriEditor
'           esriFramework
'           esriGeoDatabase
'           esriGeometry
'           esriSystemUI
'       File Dependencies
'           frmTaxlotAssignment
'           basGlobals
'           basUtilities
'
' Issues:
'       None are known at this time (JWalton 2/8/2007)
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
Private WithEvents m_pDocEvents As esriArcMapUI.DocumentEvents
Attribute m_pDocEvents.VB_VarHelpID = -1
Private m_pTaxlotFLayer As esriCarto.IFeatureLayer2
Private m_pMap As esriCarto.IMap
Private WithEvents m_pViewEvents As esriCarto.Map
Attribute m_pViewEvents.VB_VarHelpID = -1
Private m_pEditor As esriEditor.IEditor2
Private m_pMIFclass As esriGeoDatabase.IFeatureClass
Private m_pTaxlotFClass As esriGeoDatabase.IFeatureClass
Private m_pMIFlayer As esriCarto.IFeatureLayer2
Private m_pFrmTaxlots As frmTaxlotAssignment
Private m_pBitmap As IPictureDisp
Private m_pDisabledCursor As IPictureDisp
Private m_pEnabledCursor As IPictureDisp
Private m_bEditStatus As Boolean
Private m_bToolStatus As Boolean
Private m_lCNTaxlotFld As Long
Private m_lCNMNumFld As Long
Private m_lOMMapTaxlotFld As Long
Private m_lOMNumFld As Long
Private m_lOMTLNumFld As Long
Private m_lTLAnomalyFld As Long
Private m_lTLMNFld As Long
Private m_lTLTaxlotFld As Long
'++ START 2/12/2007
'   Added Variable Declarations
'       m_pDocEvents, m_pViewEvents, m_pFrame,
'       m_DisabledCursor, m_pEnabledCursor, m_bEditStatus
'++ END 2/12/2007
'++START 1/26/2007 JWalton
    ' Removed declaration of m_pEditorEvents as it was no longer needed
    ' Removed declaration of g_pApp as it was no longer needed
'++END 1/26/2007 JWalton

'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName As String = "cmdTaxlotAssignment.cls"

'------------------------------
' Interface Implementations
'------------------------------
Implements esriSystemUI.ICommand
Implements esriSystemUI.ITool


'***************************************************************************
'Name:                  ClassInitialize
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Initialize class level variables
'Called From:   Class Object
'Description:   Loads the bitmap that describes the tool on the toolbar, and
'               loads the enabled and disabled cursors for the tool from the
'               resource file
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pBitmap - The toolbar icon that describes the tool
'               m_pEnabledCursor - The enabled cursor icon
'               m_pDisabledCursor - The disabled cursor icon
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

    '++ START JWalton 1/31/2007
    Set m_pBitmap = LoadResPicture("Taxlot_Icon", vbResBitmap)
    Set m_pDisabledCursor = LoadResPicture(102, vbResCursor)
    Set m_pEnabledCursor = LoadResPicture(101, vbResCursor)
    '++ END JWalton 1/31/2007
    
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
'Purpose:       Clean up class level objects upon termination
'Called From:   Class Object
'Description:   Close and nullify the associated form variable, nullify any
'               references to graphics, feature classes, layers, and the
'               Editor.
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

Private Sub Class_Terminate()
On Error GoTo ErrorHandler
    
    '++START JWalton 1/31/2007 Removed unload of form in favor of setting form variable to Nothing
    'If g_pForms.IsFormOpen("frmTaxlotAssignment") Then Unload m_pFrmTaxlots
    'Set m_pFrmTaxlots = Nothing
    '++END JWalton 1/31/2007
    
    ' Clean up variables
    Set m_pDoc = Nothing
    Set m_pMap = Nothing
    Set m_pTaxlotFLayer = Nothing
    Set m_pTaxlotFClass = Nothing
    Set m_pMIFlayer = Nothing
    Set m_pMIFclass = Nothing
    Set m_pEditor = Nothing
    
    '++START 1/26/2007 JWalton
      ' Removed assignment of m_pEditorEvents to Nothing as it was no longer needed
      ' Removed assignment of g_pApp to Nothing
    '++END 1/26/2007 JWalton

    Exit Sub
ErrorHandler:
    HandleError True, _
                "Class_Terminate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

Private Property Get ICommand_Checked() As Boolean
'++START 1/26/2007 JWalton
    ' Not Implemented
'++END 1/26/2007 JWalton
End Property

Private Property Get ICommand_Enabled() As Boolean
    On Error GoTo ErrorHandler
    
'++START JWalton 1/25/2007 ICommand_Enabled state linked directly to the editor and the DLL
    If g_bDLLEnabled Then
        If m_pEditor Is Nothing Then
            ICommand_Enabled = False
          Else
            ICommand_Enabled = (m_pEditor.EditState = esriEditor.esriEditState.esriStateEditing)
        End If
      Else
        ICommand_Enabled = False
    End If
'++STOP JWalton 1/25/2007

    Exit Property
ErrorHandler:
    HandleError True, _
                "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Property

Private Property Get ICommand_Name() As String
    On Error GoTo ErrorHandler

     ICommand_Name = "TaxlotAssignment"

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

     ICommand_Caption = "Taxlot Assignment"

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

     ICommand_Tooltip = "Taxlot Assignment Tool"

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
    
  ICommand_Message = "Set starting value and increment value to populate values in Taxlots"

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
'++START 1/26/2007 JWalton
    ' Not Implemented
'++END 1/26/2007 JWalton
End Property
 
Private Property Get ICommand_HelpContextID() As Long
'++START 1/26/2007 JWalton
    ' Not Implemented
'++END 1/26/2007 JWalton
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
On Error GoTo ErrorHandler
    
   ICommand_Bitmap = m_pBitmap

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
On Error GoTo ErrorHandler
    ' Variable declarations
    Dim pId As New esriSystem.UID
  
'++ START JWalton 1/26/2007 Only continue for ArcMap
    If TypeOf hook Is esriArcMapUI.IMxApplication Then
        If g_pApp Is Nothing Then
          Set g_pApp = hook
        End If
        g_bDLLEnabled = True
      Else
        g_bDLLEnabled = False
        Exit Sub
    End If
    
    'Get a reference to the editor extension
    pId = "esriEditor.Editor"
    Set m_pEditor = g_pApp.FindExtensionByCLSID(pId)
'++End JWalton 1/26/2007

'++ START JWalton 1/25/2007
    ' Removed event editor, document , and map declarations to ICommand_OnClick
'++ END  JWalton 1/25/2007++++++++++

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
'Purpose:       Set references to feature classes that will be used
'               throughout the life of the class
'Called From:   Class Object
'Description:   Verify the existence of the Taxlot and Map Index feature
'               classes, and allow the user to load such feature classes if
'               they do not exist.
'               Catalog the taxlot fields from the Taxlot feature class.
'               Load and show the Taxlot Assignment form to the user.
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_lOMTLNumFld - The index of the Taxlot number field
'               m_lOMNumFld - The index of the ORMAP number field
'               m_lOMMapTaxlotFld - The index of the ORMAP Taxlot number
'                                   field
'               m_lTLAnomalyFld - The index of the anomaly field
'               m_lTLTaxlotFld - The index of the taxlot field
'               m_pDoc - Reference to the current document
'               m_pMap - Reference to the focus map in the current document
'               m_pMIFClass - The Map Index feature class
'               m_pMILayer - The Map Index feature layer
'               m_pTaxlotFCClass - The Taxlot feature class
'               m_pTaxlotFLayer - The Taxlot feature layer
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  defined single exit point
'***************************************************************************

Private Sub ICommand_OnClick()
On Error GoTo ErrorHandler
    '++ START JWalton 2/12/2007
    '   Removed variable declaration as no longer needed
    '++ END JWalton 2/12/2007
    
    ' Find Taxlot feature layer
    Set m_pTaxlotFLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If m_pTaxlotFLayer Is Nothing Then
    '++ START JWalton 1/31/2007 Added load option for this feature class
        ' Attempt to load and find the Map Index in the map document
        If LoadFCIntoMap(g_pFldnames.FCTaxlot) Then
            Set m_pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
          Else
            GoTo Process_Exit
        End If
        
        ' Exit if unable to load the Map Index layer
        If m_pTaxlotFLayer Is Nothing Then GoTo Process_Exit
    '++ END JWalton 1/31/2007
    End If
    
    ' Find MapIndex feature layer
    Set m_pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If m_pMIFlayer Is Nothing Then
        ' Attempt to load and find the Map Index in the map document
        If LoadFCIntoMap(g_pFldnames.FCMapIndex) Then
            Set m_pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
          Else
            GoTo Process_Exit
        End If
        
        ' Exit if unable to load the Map Index layer
        If m_pMIFlayer Is Nothing Then GoTo Process_Exit
    End If

'++ START JWalton 1/25/2007 Removed from ICommand_OnCreate to this location
    ' Initialize document and map objects, and their events for tool reference only
    Set m_pDoc = g_pApp.Document
    Set m_pMap = m_pDoc.FocusMap
    Set m_pDocEvents = m_pDoc
    Set m_pViewEvents = m_pMap
    
    '++START JWalton 1/26/2007
        ' Removed assignment of m_pEditorEvents to m_EditorEventsas it was no longer needed
    '++END JWalton 1/26/2007
'++End JWalton 1/25/2007
    
    '++ START JWalton 1/31/2007
    ' Obtain references to feature layer feature classes
    Set m_pTaxlotFClass = m_pTaxlotFLayer.FeatureClass
    Set m_pMIFclass = m_pMIFlayer.FeatureClass
    
    ' Find the ORMAP Taxlot field
    m_lOMTLNumFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLOrmapTaxlotFN)
    If m_lOMTLNumFld = -1 Then GoTo Process_Exit
    
    ' Find the ORMap Map Number field
    m_lOMNumFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLOrmapMapNumberFN)
    If m_lOMNumFld = -1 Then GoTo Process_Exit
    
    ' Find the Map Taxlot field
    m_lOMMapTaxlotFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLMapTaxlotFN)
    If m_lOMMapTaxlotFld = -1 Then GoTo Process_Exit

    ' Find the Taxlot Taxlot field
    m_lTLTaxlotFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLTaxlotFN)
    If m_lTLTaxlotFld = -1 Then GoTo Process_Exit
    
    ' Find the Taxlot Anomaly field
    m_lTLAnomalyFld = LocateFields(m_pTaxlotFClass, g_pFldnames.TLAnomalyFN)
    If m_lTLAnomalyFld = -1 Then GoTo Process_Exit
    
    ' Initialize the TaxlotAssignment form
    If m_pFrmTaxlots Is Nothing Then
        Set m_pFrmTaxlots = New frmTaxlotAssignment
    End If
    m_pFrmTaxlots.Frame.Visible = True
    '++ END JWalton 1/31/2007

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
On Error GoTo ErrorHandler
    '++ START JWalton 1/31/2007
    ' Sets the proper cursor given the data
    If g_pForms.IsFormOpen("frmTaxlotAssignment") Then
        If m_pFrmTaxlots.IsValidData() Then
            ITool_Cursor = m_pEnabledCursor
            m_bEditStatus = True
          Else
            ITool_Cursor = m_pDisabledCursor
            m_bEditStatus = False
        End If
      Else
        ITool_Cursor = m_pDisabledCursor
        m_bEditStatus = False
    End If
    '++ END JWalton 1/31/2007

    Exit Property
ErrorHandler:
    HandleError True, _
                "ITool_Cursor " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Property
 
'***************************************************************************
'Name:                  ITool_OnMouseDown
'Initial Author:        James Moore
'Subsequent Author:     <<Type your name here>>
'Created:               10/11/2006
'Purpose:       Populate attributes in the Taxlot feature class
'Called From:   Class Object
'Description:   This code is fired when the user points to a taxlot with
'               the TaxlotAssignment tool and clicks
'Methods:       Search for underlying MapIndex.
'               Populate attributes based on Map Index
'               Assign Taxlot based on input taxlot and optional incremented
'               Value settings
'Inputs:        Button, Shift, X, Y (System)
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Using strcomp function. Changed to intrinsic
'                           constant, established single exit point
'james moore    11-9-2006   Added code to update the MapTaxlot field
'John Walton    2/8/2007    Renamed strShapeFieldName to sShapeFieldName to
'                           conform to variable naming conventions
'***************************************************************************

Private Sub ITool_OnMouseDown( _
  ByVal Button As Long, _
  ByVal Shift As Long, _
  ByVal X As Long, _
  ByVal Y As Long)
On Error GoTo ErrorHandler
'++ START JWalton 1/31/2007 Centralized variable declarations
    Dim pActiveView As esriCarto.IActiveView
    Dim pMap As esriCarto.IMap
    '++ START JWalton 2/14/2007 Additional Variable Declarations
    Dim pCmdItem As esriFramework.ICommandItem
    Dim pDoc As esriFramework.IDocument
    '++ END JWalton 2/14/2007
    Dim pFeature As esriGeoDatabase.IFeature
    Dim pFCursor As esriGeoDatabase.IFeatureCursor
    Dim pSpatialFilter As esriGeoDatabase.ISpatialFilter
    Dim pEnv As esriGeometry.IEnvelope
    Dim pGeometry As esriGeometry.IGeometry
    Dim pPoint As esriGeometry.IPoint
    '++ START JWalton 2/14/2007 Additional Variable Declarations
    Dim pUID As New esriSystem.UID
    '++ END JWalton 2/14/2007
    Dim iResponse As Integer
    Dim lIncrement As Long
    Dim lValue As Long
    Dim sAnom As String
    Dim sExistOMNum As String
    Dim sExistTLNumVal As String
    Dim sNewTLNum As String
    Dim sNewTLNum_5digit As String
    Dim sOMMapTaxlot As String
    Dim sOMTLNval As String
    Dim sShapeFieldName As String
    Dim sShortOMNum As String
    Dim sTLMapSufNumVal As String
    Dim sTLMapSufTypeVal As String
'++ END JWalton 1/31/2007
    
    '++ START JWalton 2/8/2007
    If Button <> vbLeftButton Then Exit Sub
    '++ END JWalton 2/8/2007

    '++ START JWalton 1/31/2007
    ' Insure that it is the tool has the proper data to continue
    If Not m_bEditStatus Then Exit Sub
    '++ END JWalton 1/31/2007
    
    'If "NUMBER" selected, then make sure value is numeric
    If StrComp(m_pFrmTaxlots.PolygonType, "NUMBER", vbTextCompare) = 0 Then
        '++ START JWalton 2/8/2007
        If IsNumeric(m_pFrmTaxlots.CurrentValue) Then
            lValue = m_pFrmTaxlots.CurrentValue
            lIncrement = m_pFrmTaxlots.Increment
          Else
            Exit Sub
        End If
        '++ END JWalton 2/8/2007
    End If
    
    ' Create a search shape out of the point that the user clicked
    Set pMap = m_pEditor.Map
    Set pPoint = m_pEditor.Display.DisplayTransformation.ToMapPoint(X, Y)
    Set pGeometry = m_pEditor.CreateSearchShape(pPoint)
    
    ' Verify the validity of the specified taxlot number, and uniqueness of
    ' numeric taxlot numbers
    If StrComp(m_pFrmTaxlots.PolygonType, "NUMBER", vbTextCompare) = 0 Then
        If Not ValidateTaxlotNum(m_pFrmTaxlots.CurrentValue, pGeometry) Then
            '++ START JWalton 1/31/2007 Added choice to the duplicate value problem
            If MsgBox("The current Taxlot value (" & m_pFrmTaxlots.CurrentValue & _
               ") is not unique withing this MapIndex.  " & vbCrLf & _
               "Attribute feature with value anyways?", vbYesNo Or vbQuestion) = vbNo Then
                GoTo Process_Exit
            End If
            '++ END JWalton 1/31/2007
        End If
    End If
    
    ' Insures the validity of the underlying map index polygon
    Set pEnv = pGeometry 'QI
    Set pSpatialFilter = New esriGeoDatabase.SpatialFilter
    Set pSpatialFilter.Geometry = pGeometry
    sShapeFieldName = m_pMIFclass.ShapeFieldName
    Set pSpatialFilter.OutputSpatialReference(sShapeFieldName) = pMap.SpatialReference
    pSpatialFilter.GeometryField = m_pMIFclass.ShapeFieldName
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    Set pFCursor = m_pMIFclass.Search(pSpatialFilter, False)
    Set pFeature = pFCursor.NextFeature
    If pFeature Is Nothing Then
        MsgBox "Unable to assign taxlot values to polygons that are not within a Map Index polygon", vbOKOnly Or vbCritical
        GoTo Process_Exit
    End If
    
    ' Select any feature under the given point in the target layer
    Set pEnv = pGeometry 'QI
    Set pSpatialFilter = New esriGeoDatabase.SpatialFilter
    Set pSpatialFilter.Geometry = pGeometry
    sShapeFieldName = m_pTaxlotFClass.ShapeFieldName
    Set pSpatialFilter.OutputSpatialReference(sShapeFieldName) = pMap.SpatialReference
    pSpatialFilter.GeometryField = m_pTaxlotFClass.ShapeFieldName
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    Set pFCursor = m_pTaxlotFClass.Search(pSpatialFilter, False)
    
    ' Exit if no taxlots are selected
    If pFCursor Is Nothing Then GoTo Process_Exit
    
    Set pFeature = pFCursor.NextFeature
    If Not pFeature Is Nothing Then
        'Update the feature
        m_pEditor.StartOperation
        
        ' The current ORMAP Number
        sExistOMNum = IIf(IsNull(pFeature.Value(m_lOMNumFld)), "", pFeature.Value(m_lOMNumFld))
        
        ' Obtain it the ORMAP Number from a MapIndex polygon if it is not present
        If Len(sExistOMNum) = 0 Then
            ' This call will point m_pMIFlayer to the MapIndex feature class
            CalcTaxlotValues pFeature, m_pMIFlayer
            
            ' Refresh the ORMAP Number
            sExistOMNum = IIf(IsNull(pFeature.Value(m_lOMNumFld)), "", pFeature.Value(m_lOMNumFld))
            
            ' Stops if there is still no ORMAP number
            If Len(sExistOMNum) = 0 Then
                '++ START JWalton 2/1/2007 Changed user message
                MsgBox "ORMAPMapNumber not present in this taxlot or MapIndex" & vbCrLf & _
                    "Use the MapIndex tool to populate the ORMAPMapNumber field before using this tool", _
                    vbOKOnly, _
                    "Create Tax Number"
                GoTo Process_Exit
            End If
        End If
    End If
    
    'Assign Taxlot value
    sExistTLNumVal = IIf(IsNull(pFeature.Value(m_lTLTaxlotFld)), "", pFeature.Value(m_lTLTaxlotFld))
    
    '++ START JWalton 12/14/2007 Added test for not 0 in conditional
    ' Optionally updates the taxlot number field
    If Len(sExistTLNumVal) > 0 And sExistTLNumVal <> "0" Then
        iResponse = MsgBox("Taxlot currently has a Taxlot value (" & sExistTLNumVal & ").  Update it?", vbYesNo)
        If iResponse = vbNo Then GoTo Process_Exit
    End If
    '++ END JWalton 12/14/2007
    
    'Taxlot can be less than 5-digits
    'The Taxlot value in OrMapMapNum must be 5 digits.
    'Two versions of the taxlot number will be used for these purposes.
    '++ START JWalton 2/1/2007 Replaced form references to object variable/property references
    If StrComp(m_pFrmTaxlots.PolygonType, "NUMBER", vbTextCompare) = 0 Then
        sNewTLNum = m_pFrmTaxlots.CurrentValue  'User entered number
        '++ END JWalton 2/1/2007
        sNewTLNum_5digit = sNewTLNum
        ' Make sure number is 5 characters
        If Len(sNewTLNum_5digit) < ORMAP_TAXLOT_FIELD_LENGTH Then
            Do Until Len(sNewTLNum_5digit) = ORMAP_TAXLOT_FIELD_LENGTH
                sNewTLNum_5digit = "0" & sNewTLNum_5digit
            Loop
        End If
      Else
        'Remove leading Zeros for taxlot number if any exist
        '++ START JWalton 2/1/2007 Replaced form references to object variable/property references
        sNewTLNum_5digit = m_pFrmTaxlots.PolygonType         'Predefined selection
        '++ END JWalton 2/1/2007
        sNewTLNum = Replace(sNewTLNum_5digit, "0", "")
    End If
    
    '++ START Laura Gordon, November 29, 2005
    'Determine if Special Interests field is something other than default
    'If so, include it in ORMAPtaxlot
    sTLMapSufTypeVal = GetMapSufType(pFeature)
    sTLMapSufNumVal = GetMapSufNum(pFeature)
    If IsNull(sTLMapSufTypeVal) Then sTLMapSufTypeVal = "0"
    If IsNull(sTLMapSufNumVal) Then sTLMapSufNumVal = "000"
    '++ END Laura Gordon, November 29, 2005
    pFeature.Value(m_lTLTaxlotFld) = sNewTLNum
    
    sShortOMNum = ShortenOMMapNum(sExistOMNum)
    '++ START Laura Gordon, November 29, 2005
    sOMTLNval = sShortOMNum & sTLMapSufTypeVal & sTLMapSufNumVal & sNewTLNum_5digit
    '++ END Laura Gordon, November 29, 2005
    '++ START JWM 11/09/2006 create  masked value from a combination of ORMapNum and the new taxlot
    sOMMapTaxlot = gfn_s_CreateMapTaxlotValue(sExistOMNum & sNewTLNum_5digit, g_pFldnames.MapTaxlotFormatString)
    pFeature.Value(m_lOMMapTaxlotFld) = sOMMapTaxlot
    '++ END JWM 11/09/2006
    
    'Copy Anomaly from MapIndex
    sAnom = ParseOMMapNum(sExistOMNum, "anomaly")
    pFeature.Value(m_lTLAnomalyFld) = sAnom
    
    'Assign OrmapTaxlot value
    pFeature.Value(m_lOMTLNumFld) = sOMTLNval
    pFeature.Store
    
    'AutoIncrement if necessary
    '++ START JWalton 1/31/2007 Removed form control references in favor of Form Properties
    If m_pFrmTaxlots.Increment > 0 And StrComp(m_pFrmTaxlots.PolygonType, "NUMBER", vbTextCompare) = 0 Then
        m_pFrmTaxlots.CurrentValue = lValue + lIncrement
    End If
    '++ END JWalton 1/31/2007
    
    ' End the edit operation
    m_pEditor.StopOperation "AutoIncrement Attribute"
    
    '++ START JWalton 2/14/2007
    ' Select the feature
    pMap.ClearSelection
    pMap.SelectFeature m_pTaxlotFLayer, pFeature
    
    ' Insures that this tool keeps the focus
    m_bToolStatus = True
    pUID = "TaxlotEditing.cmdTaxlotAssignment"
    Set pDoc = m_pDoc
    Set pCmdItem = pDoc.CommandBars.Find(pUID, True, False)
    Set g_pApp.CurrentTool = pCmdItem
    m_bToolStatus = False
    '++ END JWalton 2/14/2007
    
    ' Update the view
    Set pActiveView = pMap
    pActiveView.PartialRefresh esriCarto.esriViewDrawPhase.esriViewGeoSelection Or _
                               esriCarto.esriViewDrawPhase.esriViewGraphics, _
                               Nothing, _
                               Nothing

Process_Exit:
    Exit Sub
    
ErrorHandler:
    '++ START JWalton 2/1/2007
    ' Abort any ongoing edit operations
    m_pEditor.AbortOperation
    '++ END JWalton 2/1/2007
    
    ' Handle the error
    HandleError True, _
                "ITool_OnMouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub
 
Private Sub ITool_OnMouseMove( _
  ByVal Button As Long, _
  ByVal Shift As Long, _
  ByVal X As Long, _
  ByVal Y As Long)
'++START 1/26/2007 JWalton
    ' Not Implemented
'++END 1/26/2007 JWalton
End Sub
 
Private Sub ITool_OnMouseUp( _
  ByVal Button As Long, _
  ByVal Shift As Long, _
  ByVal X As Long, _
  ByVal Y As Long)
'++START 1/26/2007 JWalton
    ' Not Implemented
'++END 1/26/2007 JWalton
End Sub

Private Sub ITool_OnDblClick()
'++START 1/26/2007 JWalton
    ' Not Implemented
'++END 1/26/2007 JWalton
End Sub
 
Private Sub ITool_OnKeyDown( _
  ByVal keyCode As Long, _
  ByVal Shift As Long)
'++START 1/31/2007 JWalton
    ' Deimplemented due to implementation of form as modeless
'++END 1/31/2007
End Sub
 
Private Sub ITool_OnKeyUp( _
  ByVal keyCode As Long, _
  ByVal Shift As Long)
'++START 1/26/2007 JWalton
    ' Not Implemented
'++END 1/26/2007 JWalton
End Sub
 
Private Function ITool_OnContextMenu( _
  ByVal X As Long, _
  ByVal Y As Long) As Boolean
'++START 1/26/2007 JWalton
    ' Not Implemented
'++END 1/26/2007 JWalton
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
'
'***************************************************************************

Private Function ITool_Deactivate() As Boolean
On Error GoTo ErrorHandler
    '++ START JWalton 1/31/2007
    ' Reset the tool
    Reset m_bToolStatus
    '++ END JWalton 1/31/2007
    
    ' Deactivate the tool
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

Private Sub ITool_Refresh( _
  ByVal hdc As esriSystem.OLE_HANDLE)
'++START 1/26/2007 JWalton
    ' Not Implemented
'++END 1/26/2007 JWalton
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
    ' Reset the current tool
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
            ' Reset the current tool
            Reset True
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
            ' Reset the current tool
            Reset True
        End If
    End If
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

Private Sub Reset(Optional ByVal bDoNotResetAll As Boolean = False)
On Error Resume Next
    ' Hide the taxlot assignment form
    If bDoNotResetAll = False Then
        If g_pForms.IsFormOpen("frmTaxlotAssignment") Then Unload m_pFrmTaxlots
        Set m_pFrmTaxlots = Nothing
    End If
    
    ' Reset all class level variables
    Set m_pMIFclass = Nothing
    Set m_pMIFlayer = Nothing
    Set m_pViewEvents = Nothing
    Set m_pDocEvents = Nothing
    Set m_pMap = Nothing
    Set m_pDoc = Nothing
End Sub

'++ START 1/26/2007 JWalton
    ' The following event handlers have been removed
    '   They were removed to cmdAutoUpdate and implemented there as a tool
    '   that could be both started and stopped as needed.
    '       m_pEditorEvents_OnChangeFeature
    '       m_pEditorEvents_OnCreateFeature
    '       m_pEditorEvents_OnDeleteFeature
'++ END 1/26/2007 JWalton

'++ START JWalton 1/26/2007
    ' Removed the following event handlers as they were no longer needed
    '   m_pEditorEvents_OnStartEditing()
    '   m_pEditorEvents_OnStopEditing()
'++ STOP JWalton 1/26/2007
