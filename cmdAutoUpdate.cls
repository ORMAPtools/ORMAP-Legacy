VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cmdAutoUpdate"
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
' SCC Revision number: $Revision: 18 $
' Date of last change: $Date: 2006-11-14 13:15:34 -0800 (Tue, 14 Nov 2006) $
'
'
' File name:            clsAutoUpdate
'
' Initial Author:       JWalton
'
' Date Created:         1/26/2007
'
'       Programmatic description for Auto Update command
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
'           esriGeoDatabase
'           esriSystem
'       File Dependencies
'           Globals
'
' Issues:
'       This class could be more properly implemented as a extension to ArcMap, but that implemen-
'       would involve registry issues that are not easily dispelled.  Thus, this is implemented as
'       an ICommand instead.
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
'++ START JWalton 2/12/2007
Private WithEvents m_pDocEvents As esriArcMapUI.DocumentEvents
Attribute m_pDocEvents.VB_VarHelpID = -1
'++ END JWalton 2/12/2007
Private WithEvents m_pEditorEvents As esriEditor.Editor
Attribute m_pEditorEvents.VB_VarHelpID = -1
'++ START JWalton 2/12/2007
Private WithEvents m_pViewEvents As esriCarto.Map
Attribute m_pViewEvents.VB_VarHelpID = -1
'++ END JWalton 2/12/2007
Private m_pEditor As esriEditor.IEditor
Private m_pMap As esriCarto.IMap
Private m_pDoc As esriArcMapUI.IMxDocument
Private m_pTaxlotFClass As esriGeoDatabase.IFeatureClass
Private m_pMIFclass As esriGeoDatabase.IFeatureClass
Private m_pMIFlayer As esriCarto.IFeatureLayer2
Private m_lTLTaxlotFld As Long
Private m_lTLMNFld As Long
Private m_lCNTaxlotFld As Long
Private m_lCNMNumFld As Long
Private m_bAutoUpdate As Boolean
Private m_pBmp As IPictureDisp

'------------------------------
'Private Constants and Enums
'------------------------------
Private Const c_sModuleFileName = "cmdAutoUpdate"

'------------------------------
' Interfaces Implemented
'------------------------------
Implements ICommand

'***************************************************************************
'Name:                  Class_Initialize
'Initial Author:        <<Unknown>>
'Subsequent Author:     Type your name here.
'Created:               <<Unknown>>
'Purpose:       Event Handler for Class Initialize event
'Called From:   Class Object
'Description:   Load the toolbar bitmap
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pBmp - The bitmap that display on the toolbar button
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
    ' Gets a reference to the command bitmap
    Set m_pBmp = LoadResPicture("AutoUpdate_Icon", vbResBitmap)
End Sub
 
Private Property Get ICommand_Enabled() As Boolean
On Error Resume Next
    ' Enable the control if the editor exists and is in direct edit mode
    If m_pEditor Is Nothing Then
        ICommand_Enabled = False
      Else
        ICommand_Enabled = (m_pEditor.EditState = esriEditor.esriEditState.esriStateEditing) And g_bDLLEnabled
    End If
End Property
 
Private Property Get ICommand_Checked() As Boolean
    ICommand_Checked = m_bAutoUpdate
End Property
 
Private Property Get ICommand_Name() As String
    ICommand_Name = "ORMAP_AutoUpdate"
End Property
 
Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Auto Update"
End Property
 
Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Auto Update Feature(s) Tool"
End Property
 
Private Property Get ICommand_Message() As String
    ICommand_Message = "Automatically update fields in tables that conform to the ORMAP standard"
End Property
 
Private Property Get ICommand_HelpFile() As String
    ' Not Implemented
End Property
 
Private Property Get ICommand_HelpContextID() As Long
    ' Not Implemented
End Property
 
Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_pBmp
End Property
 
Private Property Get ICommand_Category() As String
    ICommand_Category = "ORMAP"
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
On Error Resume Next
    ' Variable declarations
    Dim pUID As esriSystem.UID
    
    ' Initialize objects
    Set pUID = New esriSystem.UID
    
    ' Saves a reference to the application and checks to see that it is ArcMap
    If TypeOf hook Is esriArcMapUI.IMxApplication Then
        Set g_pApp = hook
        g_bDLLEnabled = True
      Else
        Set g_pApp = Nothing
        g_bDLLEnabled = False
    End If
    
    ' Gets a reference to the ESRI Editor
    pUID.Value = "esriEditor.Editor"
    Set m_pEditor = g_pApp.FindExtensionByCLSID(pUID)
End Sub

'***************************************************************************
'Name:                  ICommand_OnClick
'Initial Author:        <<Unknown>>
'Subsequent Author:     Type your name here.
'Created:               <<Unknown>>
'Purpose:       Toggle the auto update command
'Called From:   Class Object
'Description:   Verifies that the current document is a valid document to
'               use this tool on, and enable or disable this command
'               accordingly.
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_blnAutoUpdate - Current status of the command
'               m_pDoc - Local reference to the current document
'               m_pMap - Local reference to the focus map of the current
'                        document
'               m_pEditorEvents - Editor events interface variable
'               m_pMIFClass - Local reference to the Map Index feature
'                             class
'               m_pMIFLayer - Local reference to the Map Index feature
'                             layer
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

Private Sub ICommand_OnClick()
On Error Resume Next
    ' Variable declarations
    Dim iResponse As Integer
    
    ' Toggles the state of the control and exits if disabled
    m_bAutoUpdate = Not m_bAutoUpdate
    
    If m_bAutoUpdate Then
        ' Initialize document and map objects for tool reference only
        Set m_pDoc = g_pApp.Document
        Set m_pMap = m_pDoc.FocusMap
        Set m_pViewEvents = m_pDoc.FocusMap
        Set m_pDocEvents = m_pDoc
        
        ' Initialize the editor events interface
        Set m_pEditorEvents = m_pEditor
        
        'Find MapIndex
        Set m_pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        
        ' Takes action if the MapIndex is not found
        If m_pMIFlayer Is Nothing Then
            iResponse = MsgBox("Unable to locate Map Index layer in Table of Contents.  " & _
                               vbCrLf & _
                               "This process requires a feature class called " & g_pFldnames.FCMapIndex & ".  " & vbCrLf & _
                               "Load " & g_pFldnames.FCMapIndex & " automatically?", _
                               vbYesNo)
                               
            ' Opts out if the user does not want to continue
            If iResponse <> vbYes Then GoTo Process_Exit
            
            '++ START JWalton 1/31/2007
            ' Loads the map index feature layer to the map
            If LoadFCIntoMap(g_pFldnames.FCMapIndex) Then
                Set m_pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
            End If
            '++ END JWalton 1/31/2007
            If m_pMIFlayer Is Nothing Then GoTo Process_Exit
        End If
        
        ' Saves the map index layer feature class for later use
        Set m_pMIFclass = m_pMIFlayer.FeatureClass
      Else
        Reset
    End If

Process_Exit:
    Exit Sub
End Sub

'***************************************************************************
'Name:                  m_pEditorEvents_OnChangeFeature
'Initial Author:        <<Unknown>>
'Subsequent Author:     JWalton
'Created:               1/26/2007
'Purpose:       Update autofields and annotation
'Called From:   m_pEditorEvents
'Description:   Given a feature, obj.
'               Update the autofields in obj, and, if it is an annotation
'               feature, update the annotation size
'Methods:       <<None>>
'Inputs:        <<None>>
'Parameters:    obj - Passed by the Application
'Outputs:       <<None>>
'Returns:       <<None>>
'Errors:        This routine raises no known errors.
'Assumptions:   <<None>>
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'***************************************************************************

Private Sub m_pEditorEvents_OnChangeFeature( _
  ByVal obj As esriGeoDatabase.IObject)
On Error GoTo ErrorHandler
    ' Variable declarations
    Dim u As New esriSystem.UID
    Dim pFeat As esriGeoDatabase.IFeature
    Dim pAnnotationFeature As esriCarto.IAnnotationFeature

    'These are the autocalculated fields
    UpdateAutoFields obj
    
    If IsTaxlot(obj) Then
        'Obtain ORMAPMapnum via overlay and calculate other field values
        CalcTaxlotValues obj, m_pMIFlayer
      ElseIf IsAnno(obj) Then
        Set pAnnotationFeature = obj
        
        'Get the parent feature so mapnumber can be obtained
        Dim lParentID As Long
        lParentID = pAnnotationFeature.LinkedFeatureID
        If lParentID > -1 Then 'Feature linked
            Set pFeat = GetRelatedObjects(obj)
            If pFeat Is Nothing Then GoTo OnChangeFeature_Proc_Exit
        Else
            'Not feature linked anno, so we can use the feature as is
            Set pFeat = obj
        End If
        
        'Set anno size
        SetAnnoSize obj, pFeat
    End If
    
OnChangeFeature_Proc_Exit:
    Exit Sub
    
ErrorHandler:
    HandleError False, _
                "m_pEditorEvents_OnChangeFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  m_pEditorEvents_OnCreateFeature
'Initial Author:        <<Unknown>>
'Subsequent Author:     JWalton
'Created:               1/26/2007
'Purpose:       Automatically populate fields in a feature created within a
'               Map Index feature
'Called From:   m_pEditorEvents
'Description:   Given an object, obj.
'               Update obj with all the ORMAP Elements that can be derived
'               from an overlaying feature in Map Index.
'Methods:       None
'Inputs:        obj - A newly created feature
'Parameters:    None
'Outputs:       None
'Returns:       Nothing
'Errors:        This procedure produces no known errors
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************

Private Sub m_pEditorEvents_OnCreateFeature( _
  ByVal obj As esriGeoDatabase.IObject)
On Error GoTo ErrorHandler
    ' Variable declarations
    Dim sMapNum As String
    Dim pMIFlayer As esriCarto.IFeatureLayer
    Dim pMIFclass As esriGeoDatabase.IFeatureClass
    Dim pFeat As esriGeoDatabase.IFeature
    Dim lAnnoMapNumFld As Long
    Dim pAnnotationFeature As esriCarto.IAnnotationFeature
    Dim lParentID As Long
    Dim pGeometry As esriGeometry.IGeometry
    Dim pEnv As esriGeometry.IEnvelope
    Dim pCenter As esriGeometry.IPoint
    Dim lMapScaleFld As Long
    Dim lMapNumFld As Long
    Dim vMapScaleVal As Variant
    Dim vMapNumVal As Variant
    
    'These are the autocalculated fields
    UpdateAutoFields obj

    lMapNumFld = obj.Fields.FindField(g_pFldnames.MIMapNumberFN)
    lMapScaleFld = obj.Fields.FindField(g_pFldnames.MIMapScaleFN)
    
    If IsTaxlot(obj) Then
        'Obtain ORMAPMapnum via overlay and calculate other field values
        CalcTaxlotValues obj, m_pMIFlayer
        
    ElseIf IsAnno(obj) Then
        'Capture MapNumber for each anno feature created
        lAnnoMapNumFld = LocateFields(obj.Class, g_pFldnames.MIMapNumberFN)
        If lAnnoMapNumFld = -1 Then GoTo OnCreateFeature_Process_Exit
        Set pAnnotationFeature = obj
        
        'Get the parent feature so mapnumber can be obtained
        lParentID = pAnnotationFeature.LinkedFeatureID
        If lParentID > -1 Then 'Feature linked
            Set pFeat = GetRelatedObjects(obj)
            If pFeat Is Nothing Then GoTo OnCreateFeature_Process_Exit
        Else
            'Not feature linked anno, so we can use the feature as is
            Set pFeat = obj
        End If

        ' Retrieve the map number and scale from the overlaying map index polygon
        Set pCenter = New esriGeometry.Point
        Set pGeometry = pFeat.Shape
        If pGeometry.IsEmpty Then GoTo OnCreateFeature_Process_Exit
        Set pEnv = pGeometry.Envelope
        pCenter.PutCoords pEnv.XMin + (pEnv.XMax - pEnv.XMin) / 2, _
                          pEnv.YMin + (pEnv.YMax - pEnv.YMin) / 2
        Set pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        If pMIFlayer Is Nothing Then GoTo OnCreateFeature_Process_Exit
        Set pMIFclass = pMIFlayer.FeatureClass
        sMapNum = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
        obj.Value(lAnnoMapNumFld) = sMapNum
        If lMapScaleFld > -1 Then
            vMapScaleVal = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
            If Not IsNull(vMapScaleVal) Then
                obj.Value(lMapScaleFld) = vMapScaleVal
            End If
        End If
        
        ' Set size based on mapscale
        SetAnnoSize obj, pFeat
    Else
        ' Update MapScale for all features with a MapScale field (except MapIndex)
        If lMapScaleFld > -1 And Not IsMapIndex(obj) Then
            Set pCenter = New Point
            Set pFeat = obj
            Set pGeometry = pFeat.Shape
            If pGeometry.IsEmpty Then GoTo OnCreateFeature_Process_Exit
            Set pEnv = pGeometry.Envelope
            pCenter.PutCoords pEnv.XMin + (pEnv.XMax - pEnv.XMin) / 2, pEnv.YMin + (pEnv.YMax - pEnv.YMin) / 2
            Set pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
            If pMIFlayer Is Nothing Then GoTo OnCreateFeature_Process_Exit
            Set pMIFclass = pMIFlayer.FeatureClass
            vMapScaleVal = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
            If Not IsNull(vMapScaleVal) Then
            '++ START JWalton 1/29/2007 Corrects a data incompatibility error in value assignment to the field
                If vMapScaleVal = "" Then
                    obj.Value(lMapScaleFld) = Null
                  Else
                    obj.Value(lMapScaleFld) = vMapScaleVal
                End If
            '++ END JWalton 1/29/2007
            End If
            ' If a dataset with MapNumber, populate it
            If lMapNumFld > -1 Then
                vMapNumVal = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
                If Not IsNull(vMapNumVal) Then
                '++ START JWalton 1/29/2007 Corrects a data incompatibility error in value assignment to the field
                    If vMapNumVal = "" Then
                        obj.Value(lMapNumFld) = Null
                      Else
                        obj.Value(lMapNumFld) = vMapNumVal
                    End If
                '++ END JWalton 1/29/2007
                End If
            End If
        End If
    End If
    
OnCreateFeature_Process_Exit:
    Exit Sub
  
ErrorHandler:
    HandleError False, _
                "m_pEditorEvents_OnCreateFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
  Resume
End Sub

'***************************************************************************
'Name:                  m_pEditorEvents_OnDeleteFeature
'Initial Author:        <<Unknown>>
'Subsequent Author:     JWalton
'Created:               1/26/2007
'Purpose:       Record deleted features in an object class
'Called From:   m_pEditorEvents
'Description:   Records in the Cancelled Numbers object class the map number
'               and taxlot number from the feature that was just deleted.
'Methods:       None
'Inputs:        obj - The feature that was just deleted
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

Private Sub m_pEditorEvents_OnDeleteFeature( _
  ByVal obj As esriGeoDatabase.IObject)
On Error GoTo ErrorHandler
    ' Variable declarations
    Dim pDSet As esriGeoDatabase.IDataset
    Dim pFeat As esriGeoDatabase.IFeature
    Dim pFWSpace As esriGeoDatabase.IFeatureWorkspace
    Dim pRow As esriGeoDatabase.IRow
    Dim pCNTable As esriGeoDatabase.ITable
    Dim pWSpace As esriGeoDatabase.IWorkspace
    
    'This code listens for deletes
    If Not IsOrMapFeature(obj) Then GoTo Process_Exit
    
    'If deleting taxlots, capture the mapnumber and taxlot and record in CancelledNumbers
    If IsTaxlot(obj) Then
        ' Get reference to the Cancelled Numbers object table
        Set pFeat = obj
        Set m_pTaxlotFClass = pFeat.Class
        Set pDSet = m_pTaxlotFClass
        Set pWSpace = pDSet.Workspace
        Set pFWSpace = pWSpace
        Set pCNTable = pFWSpace.OpenTable(g_pFldnames.TabCancelled)
        
        ' Handles the event that Cancelled Numbers does not exist
        If pCNTable Is Nothing Then
            GoTo Process_Exit
        End If
        
        ' Retrieve field positions
        m_lTLTaxlotFld = m_pTaxlotFClass.FindField(g_pFldnames.TLTaxlotFN)
        m_lTLMNFld = m_pTaxlotFClass.FindField(g_pFldnames.TLMapNumberFN)
        m_lCNTaxlotFld = pCNTable.FindField(g_pFldnames.TLTaxlotFN)
        m_lCNMNumFld = pCNTable.FindField(g_pFldnames.TLMapNumberFN)
        
        ' Handles the event that a taxlot field or cancelled number field does not exist
        If m_lCNTaxlotFld = -1 Or m_lCNMNumFld = -1 Then GoTo Process_Exit
        
        'If no null values, copy them to Cancelled numbers
        If Not IsNull(obj.Value(m_lTLTaxlotFld)) And Not IsNull(obj.Value(m_lTLMNFld)) Then
            Set pRow = pCNTable.CreateRow
            If pRow Is Nothing Then GoTo Process_Exit
            pRow.Value(m_lCNTaxlotFld) = obj.Value(m_lTLTaxlotFld)
            pRow.Value(m_lCNMNumFld) = obj.Value(m_lTLMNFld)
            pRow.Store
            Set pRow = Nothing
        End If
    End If
        
Process_Exit:
    Exit Sub
  
ErrorHandler:
    HandleError False, _
                "m_pEditorEvents_OnDeleteFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
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
'Subsequent Author:     <<Type your name here>>
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
'Description:   Resets the current tool and all current class level
'               variables
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
            Reset
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

Private Sub Reset()
On Error Resume Next
    ' Set the current command to Nothing
    Set g_pApp.CurrentTool = Nothing
    
    ' Reset all class level variables
    Set m_pMIFclass = Nothing
    Set m_pMIFlayer = Nothing
    Set m_pEditorEvents = Nothing
    Set m_pViewEvents = Nothing
    Set m_pDocEvents = Nothing
    Set m_pMap = Nothing
    Set m_pDoc = Nothing
    
    ' Reflects the Auto Update OFF status
    m_bAutoUpdate = False
End Sub
