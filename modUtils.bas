Attribute VB_Name = "basUtilities"
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
' File name:            basUtilities
'
' Initial Author:       <<Unknown>>
'
' Date Created:         <<Unknown>>
'
' Description:
'       General Utility Module
'       Commonly used DLL-wide procedures and functions
'
' Entry points:
'       Methods
'           AddCodesToCmb
'           AddLeadingZeroes
'           AttributeQuery
'           CalcOMTLNum
'           CalcTaxlotValues
'           CompareAndSaveValue
'           ConvertCode
'           ConvertToDescription
'           CT_GetCenterOfEnvelope
'           FileExists
'           FindControlString
'           FindFeatureLayerByDS
'           FormatOMMapNum
'           GetAnnoSizeByScale
'           GetAppRef
'           GetCentroid
'           GetDomainDefaultValue
'           GetFWorkspace
'           GetMapSufNum
'           GetMapSufType
'           GetMXDocRef
'           GetRelatedObjects
'           GetSelectedFeature
'           GetSpecialInterests
'           GetValueViaOverlay
'           gfn_l_CountTokens
'           gfn_s_CreateMapTaxlotValue
'           gsb_StartDoc
'           HasSelectedFeatures
'           IsAnno
'           IsMapIndex
'           IsOrMapFeature
'           IsTaxlot
'           LoadFCInfoMap
'           LocateFields
'           ParseOMMapNum
'           ReadValue
'           SetAnnoSize
'           ShortenOMMapNum
'           SpatialQuery
'           SpatialQueryForEdit
'           UpdateAutoFields
'           UserName
'           Validate5Digits
'           ValidateTaxlotNum
'           ZoomToExtent

' Dependencies:
'       File References
'           esriArcMapUI
'           esriCarto
'           esriDataSourcesGDB
'           esriDisplay
'           esriGeoDatabase
'           esriGeometry
'           esriSystem
'       File Dependencies
'           basWin32API
'           clsCatalogFileDlg
'           clsRowChanged
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
'       10/11/2006 -- Added comment header to each function (JWM)
'       2/8/2007 -- All inline documentation reviewed/revised (JWalton)

Option Explicit
'******************************
' Private Definitions
'------------------------------
' Private API declarations
'------------------------------
'++ START JWalton 2/6/2007
    ' Removed the Win32API functions SendMessageString, ShellExecute&, and GetUserName to basWin32API
'++ END JWalton 2/6/2007

'------------------------------
' Private Variables
'------------------------------
Private m_bContinue As Boolean
'------------------------------
'Private Constants and Enums
'------------------------------
' Variables used by the Error handler function - DO NOT REMOVE
'++ JWM 10/11/2006 Reomved the path to this module as it will not always be in the same place
Private Const c_sModuleFileName As String = "basUtilities.bas"
'++ START JWM 10/16/2006 constants for gsb_StartDoc

'++ START JWalton 2/6/2007
    ' Removed Win32API Constant definitions to basWin32API
'++ END JWalton
'
'***************************************************************************
'Name:              FindFeatureLayerByDS
'Initial Author:    <<Unknown>>
'Subsequent Author: JWalton
'Created:           <<Unknown>>
'Called From:
'               basUtilities.CalcTaxlotValues
'               basUtilities.SetAnnoSize
'               basUtilities.ValidateTaxlotNum
'               cmdArrows.GenerateHooks
'               cmdArrows.ICommand_OnClick
'               cmdArrows.ITool_OnMouseDown
'               cmdAutuUpdate.ICommand_OnClick
'               cmdAutoUpdate.m_pEditorEvents_OnCreateFeature
'               cmdLocate.ICommand_Enabled
'               cmdLocate.ICommand_OnClick
'               cmdMapIndex.m_pEditorEvents_OnSelectionChanged
'               cmdTaxlotAssignment.ICommand_OnClick
'               cmdTaxlotCombine.ICommand_Enabled
'               frmMapIndex.Form_Load
'Description:   Return the Feature Layer based on its dataset name,
'               recursively.
'               This is an easy way to locate a feature layer in the TOC.
'Methods:
'Inputs:        asDatasetName -- The name of the dataset to find
'Parameters:    <<None>>
'Outputs:       <<None>>
'Returns:       A layer object of that supports the IFeatureLayer interface.
'Errors:        This routine raises no known errors.
'Assumptions:   <<None>>
'Updates:
'Developer:     Date:       Comments:
'----------     ------      ---------
'JWalton        1/26/2007   The function was rewritten almost in it
'                           entirety in order to recursively find feature
'                           datasets that are nested in group layers.
'***************************************************************************

Public Function FindFeatureLayerByDS( _
  ByRef asDatasetName As String) As IFeatureLayer
On Error GoTo Err_Handler
    ' Variable declarations
    Dim pMXDoc As esriArcMapUI.IMxDocument
    Dim pMap As esriCarto.IMap
    Dim pFeatureLayer As esriCarto.IFeatureLayer
    Dim pDataset As esriGeoDatabase.IDataset
    Dim pLayers As esriCarto.IEnumLayer
    Dim pLayer As esriCarto.ILayer
    Dim pId As esriSystem.UID
    Dim i As Integer
    
    ' Initialize objects
    Set pMXDoc = g_pApp.Document
    Set pMap = pMXDoc.FocusMap
    Set pId = New esriSystem.UID
    
    ' Gets a reference to the feature layers collection of the document
    pId.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
    Set pLayers = pMap.Layers(pId, True)
    
    ' Determines whether or not a layer has a dataset with the given name
    pLayers.Reset
    Set pLayer = pLayers.Next
    Do While Not pLayer Is Nothing
        Set pFeatureLayer = pLayer
        Set pDataset = pFeatureLayer.FeatureClass
        If Not pDataset Is Nothing Then
            If StrComp(pDataset.Name, asDatasetName, vbTextCompare) = 0 Then
                Set FindFeatureLayerByDS = pLayer
                Exit Function
            End If
        End If
        Set pLayer = pLayers.Next
    Loop
    
Err_Handler_Resume:
    ' Clean up
    Set pId = Nothing
    Set pMap = Nothing
    Set pMXDoc = Nothing
    Exit Function
    
Err_Handler:
    ' Return nothing in the event of an error
    Set FindFeatureLayerByDS = Nothing
    
    ' Clean up and exit
    Resume Err_Handler_Resume
End Function

'***************************************************************************
'Name:                  ReadValue
'Initial Author:        Chris Buhi
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Reads a value from a row, given a field name
'Called From:   basUtilities.CompareAndSaveValue
'               frmMapIndex.InitForm
'               frmMapIndex.InitWithFeature
'Description:   Given a row of data, pRow, a field name, sFldName, and a
'               type of data, pDataType.
'               Reads the value of a field with a domain and translates
'               the value from the coded value to the coded name.
'Methods:       None
'Parameters:    None
'Outputs:       None
'Returns:       If a domain field, the descriptive value is returned instead
'               of the stored code
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Replaced all Exit Function with Goto
'                           Process_Exit to have a single exit point
'James Moore    11-1-06     Commented out dead variable
'JWalton        2/1/2007    Changed type from Variant to String
'JWalton        2/7/2007    Removed ESRI Error Handler in favor or returning
'                           a zero-length string in the event of an error
'***************************************************************************

Public Function ReadValue( _
  pRow As esriGeoDatabase.IRow, _
  pFldName As String, _
  Optional pDataType As String) As String
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pCVDomain As esriGeoDatabase.ICodedValueDomain
    Dim pDomain As esriGeoDatabase.IDomain
    Dim pField As esriGeoDatabase.IField
    Dim dtDate As Date
    Dim i As Integer
    Dim lFld As Long
    Dim sVal As String
    Dim vDomainVal As Variant
    '++ END JWalton 2/7/2007

    lFld = pRow.Fields.FindField(pFldName)
    If lFld > -1 Then
        If pDataType = "date" Then
            'If a date and value is null, return a default date value
            '??? How should this be treated?
            sVal = IIf(IsNull(pRow.Value(lFld)), dtDate, pRow.Value(lFld))
          Else
            sVal = IIf(IsNull(pRow.Value(lFld)), "", pRow.Value(lFld))
        End If
        
        'Determine if domain field
        Set pField = pRow.Fields.Field(lFld)
        Set pDomain = pField.Domain
        If pDomain Is Nothing Then
            ReadValue = sVal
            GoTo Process_Exit
          Else
            'Determine type of domain  -If Coded Value, get the description
            If TypeOf pDomain Is ICodedValueDomain Then
                Set pCVDomain = pDomain
                vDomainVal = pRow.Value(lFld)
                'Search the domain for the code
                For i = 0 To pCVDomain.CodeCount - 1
                    If pCVDomain.Value(i) = vDomainVal Then
                        'return the description
                        ReadValue = pCVDomain.Name(i)
                        GoTo Process_Exit
                    End If
                Next i
              Else ' If range domain, return the numeric value
                ReadValue = sVal
                GoTo Process_Exit:
            End If
        End If  'If pDomain is nothing/Else
        ReadValue = sVal
      Else
        'Field not found
        ReadValue = ""
    End If

Process_Exit:
    Exit Function
    
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Return a zero-length string in the event of an error
    ReadValue = ""
    '++ END JWalton 2/7/2007
End Function

'***************************************************************************
'Name:                  AddCodesToCmb
'Initial Author:        Chris Buhi
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Add the descriptive values from each domain to the drop
'               down comboboxes
'Called From:   frmMapIndex.InitEmpty
'               frmMapIndex.InitWithFeature
'Description:   Given a collection of fields, pFields, a field, sFldName,
'               the current value of the field, vCurVal, and a combobox,
'               cboValues, the function will populate the list of the
'               cboValues with domain of the sFldName in pFields, and set
'               the current value to sCurVal.
'Methods:       None
'Inputs:        pFldName - Name of the field to draw the domain from
'               pFields - The fields collection that contains pFldName
'               cboValues - The combobox to populate
'               curVal - The current value of the field
'               blnAllowSpace - Allow a space/null entry in the list
'Parameters:    None
'Outputs:       None
'Returns:       A boolean value indicating the success of adding coded
'               values to the combobox
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point using Goto and made optional
'                           parameter true
'James Moore    11-13-06    Removed dead variables
'John Walton    2/7/2007    Renamed variables in accordance to variable
'                           naming conventions
'***************************************************************************

Public Function AddCodesToCmb( _
  ByVal sFldName As String, _
  ByVal pFields As esriGeoDatabase.IFields, _
  ByRef cboValues As ComboBox, _
  ByVal vCurVal As Variant, _
  Optional ByVal bAllowSpace As Boolean = True) As Boolean
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pCVDomain As esriGeoDatabase.ICodedValueDomain
    Dim pDomain As esriGeoDatabase.IDomain
    Dim pField As esriGeoDatabase.IField
    Dim i As Long
    Dim lCodes As Long
    Dim lFld As Long
    '++ END JWalton 2/7/2007
  
   'Get the Coded Value Domain from the field
    lFld = pFields.FindField(sFldName)
    If lFld > -1 Then
        Set pField = pFields.Field(lFld)
        Set pDomain = pField.Domain
        If pDomain Is Nothing Then
            AddCodesToCmb = False
            GoTo Process_Exit
          Else
            'Determine type of domain  -If Coded Value, get the description
            If TypeOf pDomain Is esriGeoDatabase.ICodedValueDomain Then
                Set pCVDomain = pDomain
                ' +++ Get a count of the coded values
                lCodes = pCVDomain.CodeCount
                ' +++ Loop through the list of values and add them
                ' +++ and their names to the combo box
                If Not bAllowSpace Then
                    With cboValues
                        If .ListCount > 0 Then
                            If (.List(0) = "") Then
                                .RemoveItem (0)
                            End If
                        End If
                        If .ListCount > 0 Then
                            '++ JWM 10/11/2006 Is this if statement comparing against the same thing ?
                            If (.List(.ListCount - 1) = "") Then
                                .RemoveItem (.ListCount - 1)
                            End If
                        End If
                    End With
                End If
                For i = 0 To lCodes - 1
                    'Commented line adds codes and description
                    cboValues.AddItem pCVDomain.Name(i)
                Next i
                'Successful completion of addition
                'If current value is null, add an empty string and make it active
                If vCurVal = "" Then
                    If bAllowSpace Then
                        cboValues.AddItem ""
                        cboValues.ListIndex = FindControlString(cboValues, "", 0, True)
                      Else
                        cboValues.ListIndex = 0
                    End If
                  Else 'Otherwise, select the existing value from the list
                    cboValues.ListIndex = FindControlString(cboValues, vCurVal, 0, True)
                End If
    
                AddCodesToCmb = True
              Else
                'if Range Domain, do not add values
                AddCodesToCmb = False
            End If
        End If 'if a valid domain
      Else 'Field not found
        AddCodesToCmb = False
    End If

Process_Exit:
    Exit Function
ErrorHandler:
    HandleError True, _
                "AddCodesToCmb " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  ConvertCode
'Initial Author:        Chris Buhi
'Subsequent Author:     <<Type your name here>>
'Created:               10/11/2006
'Purpose:       Converts a domain descriptive value to the stored code
'Called From:   basUtilities.CalcTaxlotValues
'               basUtilites.CompareAndSaveValue
'               frmMapIndex.cmbSufftype_Validate
'               frmMapIndex.cmdEditSave_Click
'Description:   Given a field, sFldName, in a collection of fields, pFields,
'               and a value, vVal.
'               Locates the sFldName in pFields and gets a reference to the
'               field's domain, and then finds the coded value in the domain
'               that corresponds to the coded name vVal
'Methods:       None
'Inputs:        pFields - An object that supports the IFields interface
'               sFldName - A field that exists in pFields
'               vVal - A coded name to covert to a coded value
'Parameters:    None
'Outputs:       None
'Returns:       A string that represents the domain coded value that
'               corresponds with the coded name, vVal, or a zero-length string.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point using goto
'James Moore    11-13-06    Removed dead variable
'JWalton        1/31/2007   Changed parameters -- pRow as IRow to pFields
'                           as IFields
'JWalton        2/1/2007    Changed return type from Variant to String
'JWalton        2/7/2007    Renamed variables in accordance with variable
'                           naming conventions
'                           Removed ESRI Error Handler in favor or returning
'                           a zero-length string
'***************************************************************************

Public Function ConvertCode( _
  ByVal pFields As esriGeoDatabase.IFields, _
  ByVal sFldName As String, _
  ByVal vVal As Variant) As String
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pDomain As esriGeoDatabase.IDomain
    Dim pField As esriGeoDatabase.IField
    Dim pCVDomain As esriGeoDatabase.ICodedValueDomain
    Dim i As Integer
    Dim lFld As Long
    '++ END JWalton 2/7/2007
    
    lFld = pFields.FindField(sFldName)
    If lFld > -1 Then
        'Determine if domain field
        Set pField = pFields.Field(lFld)
        Set pDomain = pField.Domain
        If pDomain Is Nothing Then
            ConvertCode = vVal
            GoTo Process_Exit
          Else
            'Determine type of domain  -If Coded Value, get the description
            If TypeOf pDomain Is esriGeoDatabase.ICodedValueDomain Then
                Set pCVDomain = pDomain
                'Given the description, search the domain for the code
                For i = 0 To pCVDomain.CodeCount - 1
                    If pCVDomain.Name(i) = vVal Then
                        ConvertCode = pCVDomain.Value(i) 'Return the code value
                        GoTo Process_Exit
                    End If
                Next i
              Else ' If range domain, return the numeric value
                ConvertCode = vVal
                GoTo Process_Exit
            End If
        End If  'If pDomain is nothing/Else
        ConvertCode = vVal
      Else
        'Field not found
        ConvertCode = ""
    End If
    
Process_Exit:
    Exit Function
  
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Returns a zero-length string in the case of an error
    ConvertCode = ""
    '++ END JWalton 2/7/2007
End Function
 
'***************************************************************************
'Name:                  ConvertToDescription
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Converts a domain descriptive value to the stored code
'Called From:   cmdArrow.ITool_OnMouseDown
'               frmMapIndex.cmbSufftype_Validate
'               frmMapIndex.InitEmpty
'               frmMapIndex.InitWithFeature
'Description:   Given a field, sFldName, in a collection of fields, pFields,
'               and a coded value name, vVal.
'               Locates the sFldName in pFields and gets a reference to the
'               field's domain, and then finds the coded name in the domain
'               that corresponds to the coded value vVal
'Methods:       None
'Inputs:        pFields - An object that supports the IFields interface
'               sFldName - A field that exists in pFields
'               vVal - A coded value to covert to a coded name
'Parameters:    None
'Outputs:       None
'Returns:       A string that represents the domain coded name that
'               corresponds with the coded value, vVal, or a zero-length
'               string.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point using goto
'James Moore    11-13-06    Removed dead variable
'JWalton        2/7/2007    Renamed variables in accordance with variable
'                           naming conventions
'                           Removed ESRI Error Handler in favor or returning
'                           a zero-length string
'***************************************************************************

Public Function ConvertToDescription( _
  pFlds As esriGeoDatabase.IFields, _
  sFldName As String, _
  vVal As Variant) As Variant
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pCVDomain As esriGeoDatabase.ICodedValueDomain
    Dim pField As esriGeoDatabase.IField
    Dim pDomain As esriGeoDatabase.IDomain
    Dim i As Integer
    Dim lFld As Long
    '++ END JWalton 2/7/2007
    
    lFld = pFlds.FindField(sFldName)
    If lFld > -1 Then
      'Determine if domain field
     Set pField = pFlds.Field(lFld)
      Set pDomain = pField.Domain
      If pDomain Is Nothing Then
        ConvertToDescription = vVal
        GoTo Process_Exit
      Else
        'Determine type of domain  -If Coded Value, get the description
        If TypeOf pDomain Is esriGeoDatabase.ICodedValueDomain Then
          Set pCVDomain = pDomain
          'Given the description, search the domain for the code
          For i = 0 To pCVDomain.CodeCount - 1
            If pCVDomain.Value(i) = vVal Then
              ConvertToDescription = pCVDomain.Name(i) 'Return the code value
              GoTo Process_Exit
            End If
          Next i
        Else ' If range domain, return the numeric value
          ConvertToDescription = vVal
          GoTo Process_Exit
        End If
      End If  'If pDomain is nothing/Else
      ConvertToDescription = vVal
    Else
      'Field not found
      ConvertToDescription = ""
    End If 'If lFld > -1/Else

Process_Exit:
    Exit Function
    
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Returns a zero-length string in the case of an error
    ConvertToDescription = ""
    '++ END JWalton 2/7/2007
End Function
 


'***************************************************************************
'Name:                  GetValueViaOverlay
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Overlay the passed in feature with a feature class
'Called From:   basUtilities.CalcTaxlotValues
'               basUtilities.SetAnnoSize
'               basUtilities.ValidateTaxlotNum
'               cmdArrows.GenerateHook
'               cmdArrows.ITool_OnMouseDown
'               cmdAutoUpdate.m_pEditorEvents_OnCreateFeature
'Description:   Given a search geometry, pGeom, and feature class that
'               overlays the search geometry, pOverlayFC, and a field
'               name, sFldName, to find the value of.
'               Find the first feature in pOverlayFC that intersects pGeom
'               and return the value of the field in that feature
'Methods:       None
'Inputs:        pGeom - Search geometry
'               pOverlayFC - Overlying feature class
'               sFldName - Name of field to return value for
'Parameters:    None
'Outputs:       None
'Returns:       Returns the value from the specified field as a variant
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'JWM            10-30-06    Checking the length of sFldName in 'if'
'                           statement instead of checking for a empty string
'***************************************************************************

Public Function GetValueViaOverlay( _
  ByRef pGeom As esriGeometry.IGeometry, _
  ByRef pOverlayFC As esriGeoDatabase.IFeatureClass, _
  ByVal sFldName As String) As Variant
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pFeat As esriGeoDatabase.IFeature
    Dim pFeatCur As esriGeoDatabase.IFeatureCursor
    Dim lFld As Long
    '++ END JWalton 2/7/2007
    
    If Not pGeom Is Nothing And Not pOverlayFC Is Nothing And Len(sFldName) > 0 Then
        Set pFeatCur = SpatialQuery(pOverlayFC, pGeom, esriSpatialRelIntersects)
        If Not pFeatCur Is Nothing Then
            'Get the first feature.  if more than one, let the user decide
            Set pFeat = pFeatCur.NextFeature
            If Not pFeat Is Nothing Then
                lFld = pFeat.Fields.FindField(sFldName)
                If lFld > -1 Then
                    'Get the  value
                    GetValueViaOverlay = IIf(IsNull(pFeat.Value(lFld)), Null, pFeat.Value(lFld))
                  Else
                    '++ START JWalton 2/1/2007 Added Else Clause
                    GetValueViaOverlay = Null
                    '++ END JWalton 2/1/2007
                End If
            End If
        End If
      Else
        '++ START JWalton 2/7/2007 Added Else Clause
        GetValueViaOverlay = Null
        '++ END JWalton 2/7/2007
    End If

Process_Exit:
  Exit Function

ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Return Null in the case of an error
    GetValueViaOverlay = Null
    '++END JWalton 2/7/2007
End Function

'***************************************************************************
'Name:                  FindControlString
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Find a string in a listbox or combobox control.
'Called From:   basUtilities.AddCodesToCmb
'Description:   Given a control, ctrl, a string to search for, sSearch,
'               an index to start at, lStartIdx, and whether not to search
'               for an exact match, bExactMatch
'               Send a message for the control instructing it to find the
'               first element that matches sSearch according the options
'               bExactMatch and lStartIdx, and return the position that the
'               match is found at.
'Methods:       Uses an API call to Windows in order to find the first
'               element that matchs sSearch according to the options
'               specified
'Inputs:        ctrl - A listbox or combobox
'               sSearch - The value to search for
'               lStartIdx - Index position to start at (Default at beginning)
'               bExactMatch - Find sSearch exactly (Default to False)
'Parameters:    None
'Outputs:       None
'Returns:       The index of the match, or -1 if not found.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point
'JWalton        2/7/2007    Removed ESRI Error Handler in favor or returning
'                           -1
'***************************************************************************

Public Function FindControlString( _
  ByRef ctrl As Control, _
  ByVal sSearch As String, _
  Optional ByRef lStartIdx As Long = -1, _
  Optional ByRef bExactMatch As Boolean = False) As Long
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim uMsg As Long
    '++ END JWalton 2/7/2007
    
    ' Determines the proper message to send
    If TypeOf ctrl Is ListBox Then
        uMsg = IIf(bExactMatch, LB_FINDSTRINGEXACT, LB_FINDSTRING)
    ElseIf TypeOf ctrl Is ComboBox Then
        uMsg = IIf(bExactMatch, CB_FINDSTRINGEXACT, CB_FINDSTRING)
    Else
        GoTo Process_Exit
    End If
    
    FindControlString = SendMessageString(ctrl.hwnd, uMsg, lStartIdx, sSearch)

Process_Exit:
    Exit Function
    
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Returns -1 in the case of an error
    FindControlString = -1
    '++ END JWalton 2/7/2007
End Function

'***************************************************************************
'Name:                  SpatialQuery
'Initial Author:        <<Unknown>>
'Subsequent Author:     Type your name here.
'Created:               <<Unknown>>
'Purpose:       Return a feature cursor based on the results of a spatial
'               query
'Called From:   basUtilites.GetValueViaOverlay
'Description:   Given a feature class, pFeatureClassIn, a search geometry,
'               pSearchGeometry, a spatial relationship, lSpatialRelation,
'               an Sql search statement, sWhereClause, and whether or not
'               the returned cursor should be updateable, bUpdateable.
'               Perform a spatial query on pFeatureClassIn where feature
'               which meet criteria sWhereClause have a relationship of
'               lSpatialRelation to pSearchGeometry.
'               The returned cursor is updatable if bUpdateable is True.
'Methods:       None
'Inputs:        pFeatureClassIn - Feature class to search
'               pSearchGeometry - Geometry to search in relation to
'               lSpatialRelation - Geometry relationship to
'                                  pSearchGeometry
'               sWhereClause - Sql Where clause
'               bUpdateable - Read/Write state of the return cursor
'Parameters:    None
'Outputs:       None
'Returns:       Returns a feature cursor that represents the results
'               of the spatial query
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Initial creation of this comment section
'John Walton    2/7/2007    Renamed strShpFld to sShpFld to conform to
'                           variable naming conventions
'John Walton    2/8/2007    Renamed arguments to conform to variable naming
'                           conventions
'                           Added argument bUpdateable that allows for the
'                           requested query to be returned read/write
'***************************************************************************

Public Function SpatialQuery( _
  ByRef pFeatureClassIn As esriGeoDatabase.IFeatureClass, _
  ByRef pSearchGeometry As esriGeometry.IGeometry, _
  ByRef lSpatialRelation As esriGeoDatabase.esriSpatialRelEnum, _
  Optional ByRef sWhereClause As String = "", _
  Optional ByVal bUpdateable As Boolean = False) As esriGeoDatabase.IFeatureCursor
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pFeatCursor As esriGeoDatabase.IFeatureCursor
    Dim pQueryFilter As esriGeoDatabase.IQueryFilter
    Dim pSpatialFilter As esriGeoDatabase.ISpatialFilter
    Dim sShpFld As String
    '++ END JWalton 2/7/2007
    
    ' create a spatial query filter
    Set pSpatialFilter = New esriGeoDatabase.SpatialFilter
    
    ' specify the geometry to query with
    Set pSpatialFilter.Geometry = pSearchGeometry
    
    ' specify what the geometry file is called on the Feature Class that we will be querying against
    sShpFld = pFeatureClassIn.ShapeFieldName
    pSpatialFilter.GeometryField = sShpFld
    
    'specify the type of spatial operation to use
    pSpatialFilter.SpatialRel = lSpatialRelation

    ' create the where statement
    pSpatialFilter.whereClause = sWhereClause
    
    ' perform the query
    Set pQueryFilter = pSpatialFilter
    '++ START JWalton 2/8/2007
    If bUpdateable Then
        Set pFeatCursor = pFeatureClassIn.Update(pQueryFilter, False)
      Else
        Set pFeatCursor = pFeatureClassIn.Search(pQueryFilter, False)
    End If
    '++ END JWalton 2/8/2007
    
    ' Returns the value of the function
    Set SpatialQuery = pFeatCursor

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, _
              "SpatialQuery " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
              Err.Number, _
              Err.Source, _
              Err.Description, _
              4
End Function

'***************************************************************************
'Name:                  SpatialQueryForEdit
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       See SpatialQuery
'Called From:   cmdCombin.cmdApply_Click
'Description:   See SpatialQuery
'               This function is no longer necessary, but is left intact.
'               This function is now no more than a shell that calls
'               SpatialQuery with the return updateable cursor set to True.
'Methods:       None
'Inputs:        pFeatureClassIn - Feature class to search
'               pSearchGeometry - Geometry to search in relation to
'               lSpatialRelation - Geometry relationship to
'                                  pSearchGeometry
'               sWhereClause - Sql Where clause
'Parameters:    None
'Outputs:       None
'Returns:       Returns a feature cursor that represents the results
'               of the spatial query
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/7/2007    Renamed strShpFld to sShpFld to conform to
'                           variable naming conventions
'***************************************************************************

Public Function SpatialQueryForEdit( _
  ByRef pFeatureClassIn As esriGeoDatabase.IFeatureClass, _
  ByRef pSearchGeometry As esriGeometry.IGeometry, _
  ByRef lSpatialRelation As esriGeoDatabase.esriSpatialRelEnum, _
  Optional ByRef sWhereClause As String = "") As esriGeoDatabase.IFeatureCursor
On Error GoTo ErrorHandler
    '++ START JWalton 2/8/2007
    Set SpatialQueryForEdit = SpatialQuery(pFeatureClassIn, _
                                           pSearchGeometry, _
                                           lSpatialRelation, _
                                           sWhereClause, _
                                           True)
    '++ END JWalton 2/8/2007

    Exit Function
ErrorHandler:
    HandleError True, _
                "SpatialQueryForEdit " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  AttributeQuery
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               10/11/2006
'Purpose:       Return a cursor that represents the results of an
'               attribute query
'Called From:   basUtilities.ValidateTaxlotNum
'Description:   Given a table, pTable, and a Where clause, sWhereClause.
'               Creates a cursor from pTable that contains all feature
'               records that meet the criteria in sWhereClause
'Methods:       None
'Inputs:        pTable - An object that supports the ITable interface
'               sWhereClause - An Sql Where clause
'Parameters:    None
'Outputs:       None
'Returns:       An object that supports the ICursor interface
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/7/2007    Renamed selCount to lSelCount to conform to
'                           variable naming conventions
'                           Removed variable lSelCount as it is no longer
'                           necesary
'***************************************************************************

Public Function AttributeQuery( _
  ByRef pTable As esriGeoDatabase.ITable, _
  Optional ByRef whereClause As String = "") As esriGeoDatabase.ICursor
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pCursor As esriGeoDatabase.ICursor
    Dim pQueryFilter As esriGeoDatabase.IQueryFilter
    '++ END JWalton 2/7/2007
    
    ' Return a cursor based on an attribute query
    Set pQueryFilter = New esriGeoDatabase.QueryFilter
    pQueryFilter.whereClause = whereClause
    Set pCursor = pTable.Search(pQueryFilter, False)
    
    ' Count the number of selected records
    '++ START JWalton 2/7/2007 Removed variable assignment of row count for testing
    If pTable.RowCount(pQueryFilter) = 0 Then
        Set AttributeQuery = Nothing
    Else
        Set AttributeQuery = pCursor
    End If
    '++ END JWalton 2/7/2007
    
    Exit Function
ErrorHandler:
    HandleError True, _
                "AttributeQuery " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  GetSelectedFeatures
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Return a cursor for the selected features
'Called From:   frmTaxlotCombine.ICommand_Enabled
'               frmCombine.cmdApply_Click
'               frmMapIndex.InitForm
'Description:   Given a feature layer, pFLayer.
'               References the currently selected features in pFLayer, and
'               return a cursor with the feature in it.
'Methods:       None
'Inputs:        pFLayer - The feature layer to return the selection from
'Parameters:    None
'Outputs:       None
'Returns:       An object that supports the IFeatureCursor interface
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point
'JWalton        2/7/2007    Removed ESRI Error Handler in favor or returning
'                           Nothing
'***************************************************************************

Public Function GetSelectedFeatures( _
  pFLayer As esriCarto.IFeatureLayer) As esriGeoDatabase.IFeatureCursor
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pFSelection As esriCarto.IFeatureSelection
    '++ END JWalton 2/7/2007

    '  exit if not applicable:
    If Not TypeOf pFLayer Is esriCarto.IFeatureLayer Then
        GoTo Process_Exit
    End If
    
    ' Returns a cursor with the selected items from the specified feature layer
    Set pFSelection = pFLayer
    pFSelection.SelectionSet.Search Nothing, False, GetSelectedFeatures
  
Process_Exit:
    Exit Function
    
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Return Nothing in the event of an error
    Set GetSelectedFeatures = Nothing
    '++ END JWalton 2/7/2007
End Function

'***************************************************************************
'Name:                  HasSelectedFeatures
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Determines if the feature layer has a selection
'Called From:   cmdMapIndex.m_pEditorEvents_OnSelectionChanged
'Description:   Given a feature layer, pFLayer.
'               Checking the selection set of pFlayer, determine if one,
'               many, or no features are selected.
'Methods:       None
'Inputs:        pFLayer - An object that supports the IFeatureLayer2
'                         interface
'Parameters:    None
'Outputs:       None
'Returns:       True or False
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point
'***************************************************************************

Public Function HasSelectedFeatures( _
  pFLayer As esriCarto.IFeatureLayer2) As Boolean
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pFeat As esriGeoDatabase.IFeature
    Dim pFeatCur As esriGeoDatabase.IFeatureCursor
    Dim pFSelection As esriCarto.IFeatureSelection
    '++ END JWalton 2/7/2007
    
    ' Do not continue if the passed feature layer is invalid
    If pFLayer Is Nothing Then GoTo Process_Exit
    
    '  exit if not applicable:
    If Not TypeOf pFLayer Is esriCarto.IFeatureLayer Then
        GoTo Process_Exit
    End If
  
    ' Determines if none, one, or many features are selected
    Set pFSelection = pFLayer
    pFSelection.SelectionSet.Search Nothing, False, pFeatCur
    If Not pFeatCur Is Nothing Then
        Set pFeat = pFeatCur.NextFeature
        If Not pFeat Is Nothing Then 'At least one feature selected
            Set pFeat = pFeatCur.NextFeature
            If Not pFeat Is Nothing Then 'More than one selected
                HasSelectedFeatures = False
                GoTo Process_Exit
              Else
                HasSelectedFeatures = True 'Just one selected
                GoTo Process_Exit
            End If
          Else 'nothing selected
            HasSelectedFeatures = False
            GoTo Process_Exit
        End If
    End If
  
Process_Exit:
    Exit Function
    
ErrorHandler:
    HandleError True, _
                "HasSelectedFeatures " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  ParseOMMapNum
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Return specific ORMAP values from this string as the whole
'               number represents multiple entities
'Called From:   cmdTaxlotAssignment.ITool_OnMouseDown
'Description:   Given an ORMAP Number, sVal, and a ORMAP Number part
'               specifier, sPartName.
'               Find sPartName in sVal and return it.
'               This function can be replaced by use of the ORMAPNumber
'               class.
'Methods:       None
'Inputs:        sVal - A valid ORMAP Number
'               sPartName - Part of the ORMAP Number to parse out
'Parameters:    None
'Outputs:       None
'Returns:       A string that represents the specified ORMAP Number part
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point
'John Walton    2/7/2007    Removed ESRI Error Handler in favor or returning
'                           a zero-length string in event of an error
'***************************************************************************

Public Function ParseOMMapNum( _
  ByRef sVal As String, _
  ByRef sPartName As String) As String
On Error GoTo ErrorHandler
    ' Validate the ORMAP Number by length
    If Not Len(sVal) = ORMAP_MAPNUM_FIELD_LENGTH Then
        ParseOMMapNum = ""
        GoTo Process_Exit
    End If
    
    ' Parse out the specified ORMAP part
    Select Case LCase$(sPartName)
        Case "county"
            ParseOMMapNum = ExtractString(sVal, 1, 2)
        Case "town"
            ParseOMMapNum = ExtractString(sVal, 3, 4)
        Case "townpart"
            ParseOMMapNum = ExtractString(sVal, 5, 7)
        Case "towndir"
            ParseOMMapNum = ExtractString(sVal, 8, 8)
        Case "range"
            ParseOMMapNum = ExtractString(sVal, 9, 10)
        Case "rangepart"
            ParseOMMapNum = ExtractString(sVal, 11, 13)
        Case "rangedir"
            ParseOMMapNum = ExtractString(sVal, 14, 14)
        Case "section"
            ParseOMMapNum = ExtractString(sVal, 15, 16)
        Case "qtr"
            ParseOMMapNum = ExtractString(sVal, 17, 17)
        Case "qtrqtr"
            ParseOMMapNum = ExtractString(sVal, 18, 18)
        Case "anomaly"
            ParseOMMapNum = ExtractString(sVal, 19, 20)
        Case "suffixtype"
             ParseOMMapNum = ExtractString(sVal, 21, 21)
        Case "suffixnum"
            ParseOMMapNum = ExtractString(sVal, 22, 24)
        Case Else
            '++ START JWalton 2/7/2007
            ParseOMMapNum = ""
            '++ END JWalton 2/7/2007
    End Select

Process_Exit:
  Exit Function
  
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Return a zero-length string in the event of an error
    ParseOMMapNum = ""
    '++ END JWalton 2/7/2007
End Function

'***************************************************************************
'Name:                  ExtractString
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Isolate elements of a string
'Called From:   basUtilites.ParseOMMapNum
'Description:   Given a string, sFullString, and lower bound, lLow, and an
'               upper bound, lHigh.
'               Extracts the substring from sFullString that at position
'               lLow and ends at position lHigh
'Methods:       None
'Inputs:        sFullString - The string to isolate the substring from
'               lLow - The starting character position
'               lHigh - The ending character position
'Parameters:    None
'Outputs:       None
'Returns:       A string that is a substring of sFullString.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'JWalton        2/7/2007    Changed the methodology used to isolate the
'                           substring from Left/Right to Mid
'                           Removed variables as they were no longer of any
'                           use
'                           Removed ESRI Error Handler in favor or returning
'                           a zero-length string
'***************************************************************************

Private Function ExtractString( _
  ByVal sFullString As String, _
  ByVal lLow As Long, _
  ByVal lHigh As Long) As String
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007
    If lLow <= lHigh Then
        ExtractString = Mid(sFullString, lLow, lHigh - lLow + 1)
      Else
        ExtractString = ""
    End If

  Exit Function
ErrorHandler:
    ExtractString = ""
End Function

'***************************************************************************
'Name:                  IsTaxlot
'Initial Author:        James Moore
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Determine if the feature belongs to the Taxlot feature class
'Called From:   cmdAutoUpdate.m_pEditorEvents_OnChangeFeature
'               cmdAutoUpdate.m_pEditorEvents_OnCreateFeature
'               cmdAutoUpdate.m_pEditorEvents_OnDeleteFeature
'Description:   Given an object, pObj.
'               Determine if pObj belongs to the Taxlot feature class by
'               checking the name of the dataset of pObj's feature class
'               againts the Taxlot Feature Class constant.
'Methods:       None
'Inputs:        pObj - A valid initialized geodatabase object
'Parameters:    None
'Outputs:       None
'Returns:       True or False
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'jwm            10-31-06    Using strcomp
'John Walton    2/7/2007    Removed ESRI Error Handler in favor of returning
'                           False in the case of an error
'***************************************************************************

Public Function IsTaxlot( _
  obj As esriGeoDatabase.IObject) As Boolean
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pOC As esriGeoDatabase.IObjectClass
    Dim pDS As esriGeoDatabase.IDataset
    '++ END JWalton 2/7/2007
    
    ' Extract the dataset from the object's class
    Set pOC = obj.Class
    Set pDS = pOC
    
    '++ START JWalton 2/7/2007
    ' Return the value of the function
    IsTaxlot = (StrComp(pDS.Name, g_pFldnames.FCTaxlot, vbTextCompare) = 0)
    '++ END JWalton 2/7/2007

    Exit Function
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Return False in the event of an error
    IsTaxlot = False
    '++ END JWalton 2/7/2007
End Function

'***************************************************************************
'Name:                  IsMapIndex
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Determine if the feature belongs to the Taxlot feature class
'Called From:   cmdAutoUpdate.m_pEditorEvents_OnCreateFeature
'Description:   Given an objec, pObj.
'               Compares the name of the dataset of pObj's feature class to
'               the Map Index layer name in order to determine if pObj
'               belongs to the Taxlot feature class.
'Methods:       None
'Inputs:        pObj - A valid initialized geodatabase object
'Parameters:    None
'Outputs:       None
'Returns:       True or False
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    12-18-2006  Using StrComp to compare strings
'John Walton    2/7/2007    Removed ESRI Error Handler in favor in returning
'                           False in case of an error
'***************************************************************************

Public Function IsMapIndex( _
  obj As esriGeoDatabase.IObject) As Boolean
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pOC As esriGeoDatabase.IObjectClass
    Dim pDS As esriGeoDatabase.IDataset
    '++ END JWalton 2/7/2007
    
    ' Extracts the dataset from the objects feature class
    Set pOC = obj.Class
    Set pDS = pOC

    '++ START JWalton 2/7/2007
    ' Returns the function's value
    IsMapIndex = (StrComp(pDS.Name, g_pFldnames.FCMapIndex, vbTextCompare) = 0)
    '++ END JWalton 2/7/2007
        
    Exit Function
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Return False in the event of an error
    IsMapIndex = False
    '++ END JWalton 2/7/2007
End Function

'***************************************************************************
'Name:                  IsAnno
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose        Determine if a feature is annotation
'Called From:   cmdAutoUpdate.m_pEditorEvents.OnChangeFeature
'               cmdAutoUpdate.m_pEditorEvents.OnCreateFeature
'Description:   Given an objec, pObj.
'               Compares the feature type of pObj with that of annotation
'               and return the truth value of the comparison.
'Methods:       None
'Inputs:        pObj - A valid initialized geodatabase object
'Parameters:    None
'Outputs:       None
'Returns:       True or False
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/7/2007    Removed ESRI Error Handler in favor or returning
'                           False in case of an error
'***************************************************************************

Public Function IsAnno( _
  pObj As esriGeoDatabase.IObject) As Boolean
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pDS As esriGeoDatabase.IDataset
    Dim pFC As esriGeoDatabase.IFeatureClass
    Dim pOC As esriGeoDatabase.IObjectClass
    '++ END JWalton 2/7/2007
    
    ' QI to get the dataset from the object's class
    Set pOC = pObj.Class
    Set pDS = pOC
    
    ' Determines if the object is annotation
    If TypeOf pObj Is esriGeoDatabase.IFeature Then
        Set pFC = pOC
        IsAnno = (pFC.FeatureType = esriGeoDatabase.esriFeatureType.esriFTAnnotation)
    End If

    Exit Function
ErrorHandler:
    HandleError True, _
                "IsAnno " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  ValidateTaxlotNum
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Determine the validity of a taxlot number
'Called From:   cmdTaxlotAssignment.ITool_OnMouseDown
'               frmCombine.cmdApply_Click
'Description:   Given a geometry from a feature, pGeometry, and a taxlot
'               number, sEnteredTLVal.
'               Determine if the feature represented by pGeometry with
'               taxlot sEnterTLVal is a unique and therefore valid.
'Methods:       None
'Inputs:        sEnteredTLVal - The new taxlot value to validate
'               pGeometry - The geometry of the feature to check
'Parameters:    None
'Outputs:       None
'Returns:       True or False
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point
'***************************************************************************

Public Function ValidateTaxlotNum( _
  sEnteredTLval As String, _
  pGeometry As esriGeometry.IGeometry) As Boolean
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pTaxlotFLayer As esriCarto.IFeatureLayer2
    Dim pCursor As esriGeoDatabase.ICursor
    Dim pTaxlotFClass As esriGeoDatabase.IFeatureClass
    Dim pMIFclass As esriGeoDatabase.IFeatureClass
    Dim pMIFlayer As esriCarto.IFeatureLayer2
    Dim pRow As esriGeoDatabase.IRow
    Dim sMIOMval As String
    Dim sWhere As String
    '++ END JWalton 2/7/2007
    
    ' Insure the existence of the Taxlot layer
    Set pTaxlotFLayer = FindFeatureLayerByDS(g_pFldnames.FCTaxlot)
    If pTaxlotFLayer Is Nothing Then
        MsgBox "Unable to locate Taxlot layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCTaxlot
        GoTo Process_Exit
    End If
    Set pTaxlotFClass = pTaxlotFLayer.FeatureClass
    
    ' Insures the existence of the Map Index layer
    Set pMIFlayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then
        MsgBox "Unable to locate Map Index layer in Table of Contents.  " & _
        "This process requires a feature class called " & g_pFldnames.FCMapIndex
        GoTo Process_Exit
    End If
    Set pMIFclass = pMIFlayer.FeatureClass
    

    ' Checks for the existence of a current ORMAP Number and Taxlot number
    sMIOMval = GetValueViaOverlay(pGeometry, pMIFclass, g_pFldnames.MIORMAPMapNumberFN)
    If sMIOMval = "" Then
        ValidateTaxlotNum = True
        GoTo Process_Exit
    End If
    
    'Make sure this number is unique within taxlots with this OM number
    sWhere = g_pFldnames.TLOrmapMapNumberFN & " = '" & sMIOMval & _
            "' and " & g_pFldnames.TLTaxlotFN & " = '" & sEnteredTLval & "'"
    Set pCursor = AttributeQuery(pTaxlotFClass, sWhere)
    If Not pCursor Is Nothing Then
        Set pRow = pCursor.NextRow
        If Not pRow Is Nothing Then
            ValidateTaxlotNum = False
        Else
            ValidateTaxlotNum = True
        End If
    Else
        ValidateTaxlotNum = True
    End If

Process_Exit:
    Exit Function
ErrorHandler:
    HandleError True, _
                "ValidateTaxlotNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  CalcTaxlotValues
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Calculates Taxlot values from ORMAPMapnum
'Called From:   cmdAutoUpdate.m_pEditorEvents_OnChangeFeature
'               cmdAutoUpdate.m_pEditorEvents_OnCreateFeature
'               cmdTaxlotAssignment.ITool_OnMouseDown,
'Description:   Given a Taxlot feature, a_pFeat, and the Map Index feature
'               layer, a_MIFLayer.
'               Update the ORMAP fields in a_pFeat to the reflect the
'               current ORMAP Number and Map Number elements in the
'               overlaying Map Index polygon.
'Methods:       None
'Inputs:        a_pFeat - A feature from the Taxlot feature class
'               a_MIFLayer - The Map Index feature layer
'Parameters:    None
'Outputs:       a_pFeat - By reference
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   a_pFeat is a feature in the Taxlot feature class
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point using goto
'James Moore    11-1-06     Removed dead variables
'James Moore    11/01/2006  Assigning the map taxlot field
'james Moore    12-19-06    Changed names of this subroutines arguments and
'                           implemented the logic implied for finding fields
'JWalton        2/7/2007    Replaced calls to ParseOMMapNumber with an
'                           ORMAPNumber object
'                           Removed variable sExistOMMAPNum as it is no
'                           longer necessary
'***************************************************************************

Public Sub CalcTaxlotValues( _
  ByRef a_pFeat As esriGeoDatabase.IFeature, _
  ByRef a_pMIFLayer As esriCarto.IFeatureLayer)
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pTaxlotFClass As esriGeoDatabase.IFeatureClass
    Dim pArea As esriGeometry.IArea
    Dim pCenter As esriGeometry.IPoint
    '++ START JWalton 2/7/2007 Additional variable declarations
    Dim pORMAPNumber As ORMAPNumber
    '++ END JWalton 2/7/2007
    Dim iResponse As Integer
    Dim lOMTLNumFld As Long
    Dim lOMNumFld As Long
    Dim lMNumFld As Long
    Dim lTaxlotMapAcres As Long
    Dim lTaxlotFld As Long
    Dim lTLAnomalyFld As Long
    Dim lTLCntyFld As Long
    Dim lTLMapSufTypeFld As Long
    Dim lTLMapSufNumFld As Long
    Dim lTLMapTaxlotFld As Long
    Dim lTLRangeFld As Long
    Dim lTLRangeDirFld As Long
    Dim lTLRangePartFld As Long
    Dim lTLQQFld As Long
    Dim lTLQtrFld As Long
    Dim lTLSectNumFld As Long
    Dim lTLSpecInterestFld As Long
    Dim lTLTownFld As Long
    Dim lTLTownDirFld As Long
    Dim lTLTownPartFld As Long
    Dim sExistMapNum As String
    Dim sExistOMTLNum As String
    Dim sExistVal As String
    Dim sTaxlotVal As String
    '++ START JWM 1/31/2006
    Dim sMapTaxlotID As String
    '++ END JWM 1/31/2006
    Dim sNewOMTLNum As String
    '++ END JWalton 2/7/2007
    
    ' Extract the Taxlot feature class from the passed feature
    Set pTaxlotFClass = a_pFeat.Class
    
    ' Locate the Map Index layer
    Set a_pMIFLayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If a_pMIFLayer Is Nothing Then
        '++ START JWalton 2/7/2007
            ' Removed user choice.  If the layer isn't here they cannot
            ' continue.  It makes sense that they should be given the
            ' chance to load the feature layer, and only quit if the
            ' load fails
        
        '++ START JWalton 1/31/2007
        ' Prompt the user for the location of the MapIndex feature class
        If LoadFCIntoMap(g_pFldnames.FCMapIndex, "Locate Database with Map Index") Then
            Set a_pMIFLayer = FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
        End If
        '++ END JWalton 1/31/2007
        
        ' Tests for a failure to load the MapIndex feature class
        If a_pMIFLayer Is Nothing Then GoTo Process_Exit
        '++ END JWalton 2/7/2007
    End If
    
'++ START JWM 12/19/2006 the module level boolean was not being modified anywhere
'   So to implement the logic implied here I am checking the variable's value from
'   each call.
'   The return value of -1 means the field was not found.
    
    ' Find all fields needed
    lOMTLNumFld = LocateFields(pTaxlotFClass, g_pFldnames.TLOrmapTaxlotFN)
    If lOMTLNumFld = -1 Then GoTo Process_Exit
    
    lOMNumFld = LocateFields(pTaxlotFClass, g_pFldnames.TLOrmapMapNumberFN)
    If lOMNumFld = -1 Then GoTo Process_Exit
    
    lMNumFld = LocateFields(pTaxlotFClass, g_pFldnames.TLMapNumberFN)
    If lMNumFld = -1 Then GoTo Process_Exit
    
    lTLCntyFld = LocateFields(pTaxlotFClass, g_pFldnames.TLCountyFN)
    If lTLCntyFld = -1 Then GoTo Process_Exit
    
    lTaxlotFld = LocateFields(pTaxlotFClass, g_pFldnames.TLTaxlotFN)
    If lTaxlotFld = -1 Then GoTo Process_Exit
    
    lTLTownFld = LocateFields(pTaxlotFClass, g_pFldnames.TLTownFN)
    If lTLTownFld = -1 Then GoTo Process_Exit
    
    lTLTownPartFld = LocateFields(pTaxlotFClass, g_pFldnames.TLTownPartFN)
    If lTLTownPartFld = -1 Then GoTo Process_Exit
    
    lTLTownDirFld = LocateFields(pTaxlotFClass, g_pFldnames.TLTownDirFN)
    If lTLTownDirFld = -1 Then GoTo Process_Exit
    
    lTLRangeFld = LocateFields(pTaxlotFClass, g_pFldnames.TLRangeFN)
    If lTLRangeFld = -1 Then GoTo Process_Exit
    
    lTLRangePartFld = LocateFields(pTaxlotFClass, g_pFldnames.TLRangePartFN)
    If lTLRangePartFld = -1 Then GoTo Process_Exit
    
    lTLRangeDirFld = LocateFields(pTaxlotFClass, g_pFldnames.TLRangeDirFN)
    If lTLRangeDirFld = -1 Then GoTo Process_Exit
    
    lTLSectNumFld = LocateFields(pTaxlotFClass, g_pFldnames.TLSectNumberFN)
    If lTLSectNumFld = -1 Then GoTo Process_Exit
    
    lTLQtrFld = LocateFields(pTaxlotFClass, g_pFldnames.TLQtrFN)
    If lTLQtrFld = -1 Then GoTo Process_Exit
    
    lTLQQFld = LocateFields(pTaxlotFClass, g_pFldnames.TLQtrQtrFN)
    If lTLQQFld = -1 Then GoTo Process_Exit
    
    lTLMapSufTypeFld = LocateFields(pTaxlotFClass, g_pFldnames.TLSufTypeFN)
    If lTLMapSufTypeFld = -1 Then GoTo Process_Exit
    
    lTLMapSufNumFld = LocateFields(pTaxlotFClass, g_pFldnames.TLSufNumFN)
    If lTLMapSufNumFld = -1 Then GoTo Process_Exit
    
    lTLSpecInterestFld = LocateFields(pTaxlotFClass, g_pFldnames.TLSpecInterestFN)
    If lTLSpecInterestFld = -1 Then GoTo Process_Exit
    
    lTLMapTaxlotFld = LocateFields(pTaxlotFClass, g_pFldnames.TLMapTaxlotFN)
    If lTLMapTaxlotFld = -1 Then GoTo Process_Exit
    
    lTaxlotMapAcres = LocateFields(pTaxlotFClass, g_pFldnames.TLMapAcresFN)
    If lTaxlotMapAcres = -1 Then GoTo Process_Exit
    
    lTLAnomalyFld = LocateFields(pTaxlotFClass, g_pFldnames.TLAnomalyFN)
    If lTLAnomalyFld = -1 Then GoTo Process_Exit
'++ END JWM 12/19/2006
    
    ' Obtain the map index poly via overlay
    Set pArea = a_pFeat.Shape
    Set pCenter = pArea.Centroid
    
    ' Update Acreage
    a_pFeat.Value(lTaxlotMapAcres) = pArea.Area / 43560  '(a_pFeat.Value(lTaxlotShapeArea) / 43560)
    
    ' Return and evaluate the ORMAP Number from the Map index
    Set pORMAPNumber = New ORMAPNumber
    If Not pORMAPNumber.ParseNumber(GetValueViaOverlay(pCenter, _
                                                       a_pMIFLayer.FeatureClass, _
                                                       g_pFldnames.MIORMAPMapNumberFN)) Then
        ' Exits if there is no value, or an invalid value
        GoTo Process_Exit
    End If
    
    ' Return and evaluate the Map Number from the Map Index
    sExistMapNum = GetValueViaOverlay(pCenter, a_pMIFLayer.FeatureClass, g_pFldnames.MIMapNumberFN)
    If Len(sExistMapNum) = 0 Then GoTo Process_Exit 'If no value for whatever reason, don't continue
    
    
    ' Store individual components of the map number in taxlot
    a_pFeat.Value(lOMNumFld) = pORMAPNumber.ORMAPNumber
    a_pFeat.Value(lMNumFld) = sExistMapNum
    
    ' County
    sExistVal = ConvertCode(a_pFeat.Fields, g_pFldnames.TLCountyFN, pORMAPNumber.County)
    If Len(sExistVal) And IsNumeric(sExistVal) Then
        a_pFeat.Value(lTLCntyFld) = CInt(sExistVal) 'Store county in county field
      Else
        a_pFeat.Value(lTLCntyFld) = Null
    End If
    
    ' Township
    a_pFeat.Value(lTLTownFld) = CInt(pORMAPNumber.Township)

    ' Partial Township Code
    a_pFeat.Value(lTLTownPartFld) = CDbl(pORMAPNumber.PartialRangeCode)

    ' Township Directional
    a_pFeat.Value(lTLTownDirFld) = pORMAPNumber.TownshipDirectional
    
    ' Range
    a_pFeat.Value(lTLRangeFld) = CInt(pORMAPNumber.Range)

    ' Partial Range Code
    a_pFeat.Value(lTLRangePartFld) = CDbl(pORMAPNumber.PartialRangeCode)
    
    ' Range Directional
    a_pFeat.Value(lTLRangeDirFld) = pORMAPNumber.RangeDirectional
    
    ' Section
    a_pFeat.Value(lTLSectNumFld) = CInt(pORMAPNumber.Section)
    
    ' Quarter
    a_pFeat.Value(lTLQtrFld) = pORMAPNumber.Quarter
    
    ' QuarterQuarter
    a_pFeat.Value(lTLQQFld) = pORMAPNumber.QuarterQuarter
    

    ' Map Suffix Type
    sExistVal = ConvertCode(a_pFeat.Fields, g_pFldnames.TLSufTypeFN, pORMAPNumber.SuffixType)
    a_pFeat.Value(lTLMapSufTypeFld) = sExistVal
    
    ' Map Suffix Number
    a_pFeat.Value(lTLMapSufNumFld) = pORMAPNumber.SuffixNumber
    
    ' Anomaly
    a_pFeat.Value(lTLAnomalyFld) = pORMAPNumber.Anomaly
    
    ' SpecialInterest
    sExistVal = IIf(IsNull(a_pFeat.Value(lTLSpecInterestFld)), "00000", a_pFeat.Value(lTLSpecInterestFld))
    If Len(sExistVal) <= 5 Then
        '++ START JWalton 2/7/2007
        sExistVal = String(5 - Len(sExistVal), "0") & sExistVal
        '++ END JWalton 2/7/2007
    End If
    a_pFeat.Value(lTLSpecInterestFld) = sExistVal
    
    ' Recalculate OMTaxlot
    If IsNull(a_pFeat.Value(lTaxlotFld)) Then GoTo Process_Exit
    
    ' Taxlot has actual taxlot number.  ORMAPTaxlot requires a 5-digit number, so leading zeros have to be added
    sTaxlotVal = a_pFeat.Value(lTaxlotFld)
    sTaxlotVal = AddLeadingZeros(sTaxlotVal, ORMAP_TAXLOT_FIELD_LENGTH)
    
    '++ START JWM 10/31/2006 assigning Maptaxlot field
    sMapTaxlotID = pORMAPNumber.ORMAPNumber & sTaxlotVal
    '@@ START NIS(LCOG) 02/5/2007
    '@@ DESCR: Add special code for Lane County (see comment below).
    Dim iCountyCode As Integer
    iCountyCode = CInt(Left$(sMapTaxlotID, 2))
    Select Case iCountyCode
    Case 1 To 19, 21 To 36
        a_pFeat.Value(lTLMapTaxlotFld) = gfn_s_CreateMapTaxlotValue(sMapTaxlotID, g_pFldnames.MapTaxlotFormatString)
    Case 20
        ' 1.  Lane County uses a 2-digit numeric identifier for ranges.
        '     Special handling is required for east ranges, where 02E is
        '     stored as 25, 03E as 35, etc.
        ' 2.  ORMAP standards (OCDES (pg 13); Taxmap Data Model (pg 11)) assert that
        '     this field should be equal to MAPNUMBER + TAXLOT. In this case, MAPNUMBER
        '     is already in the right format, thus removing the need for the
        '     gfn_s_CreateMapTaxlotValue function. Also, in this case, TAXLOT is padded
        '     on the left with zeros to make it always a 5-digit number (see comment
        '     above).
        a_pFeat.Value(lTLMapTaxlotFld) = sExistMapNum & sTaxlotVal
    End Select
    '@@ END NIS(LCOG) 02/5/2007
    '++ END JWM 10/31/2006
    
    ' Recalculate ORMAP Taxlot Number
    If IsNull(a_pFeat.Value(lOMTLNumFld)) Then GoTo Process_Exit
    
    ' Get the current and the new ORMAP Taxlot Numbers
    sExistOMTLNum = a_pFeat.Value(lOMTLNumFld)
    sNewOMTLNum = CalcOMTLNum(sExistOMTLNum, a_pFeat, sTaxlotVal)
    
    'If no changes, don't save value
    If StrComp(sExistOMTLNum, sNewOMTLNum, vbTextCompare) <> 0 Then
        a_pFeat.Value(lOMTLNumFld) = sNewOMTLNum
    End If

Process_Exit:
    ' Clean up
    Set pORMAPNumber = Nothing
    Exit Sub
    
ErrorHandler:
    HandleError True, _
                "CalcTaxlotValues " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  AddLeadingZeros
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Add leading zeros if necessary
'Called From:   basUtilites.CalcTaxlotValues
'               basUtilities.FormatOMMapNum
'Description:   Given a string, asCurString, and a length, lWidth.  Creates
'               a string of lWidth characters padded on the left with zeros.
'Methods:       None
'Inputs:        asCurString - The string to pad with zeros
'               lWidth - The final length of the string
'Parameters:    None
'Outputs:       None
'Returns:       A string of length lWidth characters
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  asCurString was being passed by Reference now
'                           passed by value
'***************************************************************************

Public Function AddLeadingZeros( _
  ByVal asCurString As String, _
  ByVal lWidth As Long) As String
On Error GoTo ErrorHandler

    If Len(asCurString) < lWidth Then
        asCurString = String(lWidth - Len(asCurString), "0") & asCurString
    End If
    AddLeadingZeros = asCurString

  Exit Function
ErrorHandler:
    HandleError True, _
                "AddLeadingZeros " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  CT_GetCenterOfEnvelope
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Find the center of a an envelope
'Called From:   basUtilites.SetAnnoSize
'Description:   Given an envelope, pEnv.
'               Determine the x- and y-coordinates of the center of pEnv,
'               and return them as a Point object
'Methods:       None
'Inputs:        pEnv - An envelope
'Parameters:    None
'Outputs:       None
'Returns:       A Point object that represents the center of the envelope,
'               pEnv
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************

Function CT_GetCenterOfEnvelope( _
  ByRef pEnv As esriGeometry.IEnvelope) As esriGeometry.IPoint
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pCenter As esriGeometry.IPoint
    '++ END JWalton 2/7/2007
    
    ' Initialize objects
    Set pCenter = New esriGeometry.Point
    
    ' Set the coordinates of the center of the envelope in the point
    pCenter.X = pEnv.XMin + (pEnv.XMax - pEnv.XMin) / 2
    pCenter.Y = pEnv.YMin + (pEnv.YMax - pEnv.YMin) / 2
    
    ' Returns the center of the envelope to the function
    Set CT_GetCenterOfEnvelope = pCenter

    Exit Function
ErrorHandler:
    HandleError True, _
                "CT_GetCenterOfEnvelope " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  GetRelatedObjects
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Retrieve a feature-linked annotation feature
'Called From:   cmdAutoUpdate.m_pEditorEvents_OnChangeFeature
'               cmdAutoUpdate.m_pEditorEvents_OnCreateFeature
'Description:   Given a feature, pObj.
'               Finds all related objects to the feature through the first
'               found relationship class, and returns the first related
'               object as the return value.
'               This is optimized for annotation because there is a single
'               relationship class.
'Methods:       None
'Inputs:        pObj - An initialized geodatabase object
'Parameters:    None
'Outputs:       None
'Returns:       An object that supports the IFeature object.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point
'JWalton        2/7/2007    Removed ESRI Error Handler in favor or returning
'                           Nothing
'***************************************************************************

Public Function GetRelatedObjects( _
  pObj As esriGeoDatabase.IObject) As esriGeoDatabase.IFeature
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pEnumRelClass As esriGeoDatabase.IEnumRelationshipClass
    Dim pParentFeat As esriGeoDatabase.IFeature
    Dim pRelClass As esriGeoDatabase.IRelationshipClass
    Dim pParentSet As esriSystem.ISet
    '++ END JWalton 2/7/2007
    
    ' Retrieves the objects related to the passed object
    Set pEnumRelClass = pObj.Class.RelationshipClasses(esriRelRoleAny)
    If Not pEnumRelClass Is Nothing Then
      Set pRelClass = pEnumRelClass.Next
      If Not pRelClass Is Nothing Then
          Set pParentSet = pRelClass.GetObjectsRelatedToObject(pObj)
      End If
    Else
        GoTo Process_Exit
    End If
    
    ' Returns the first related feature
    If Not pParentSet Is Nothing Then
        Set pParentFeat = pParentSet.Next
        If Not pParentFeat Is Nothing Then
            Set GetRelatedObjects = pParentFeat
        End If
    End If

Process_Exit:
    Exit Function
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Returns Nothing in the case of an error
    Set GetRelatedObjects = Nothing
    '++ END JWalton
End Function

'***************************************************************************
'Name:                  GetAnnoSizeByScale
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Determine the annotation size based on scale
'Called From:   basUtilities.SetAnnoSize
'Description:   Given a feature class, sFCName, and a scale, lScale.
'               Determines the proper size for the text in sFCName.
'               Defaults at size 5 is the scale is invalid, and size 10 if
'               sFCName is not Taxlot Acreage Annotation or Taxlot Number
'               Annotation
'Methods:       None
'Inputs:        sFCName - Feature class to find the proper annotation size
'                         for
'               lScale - The scale of the feature class
'Parameters:    None
'Outputs:       None
'Returns:       A double that represents the proper scale factor.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Using Strcomp function to compare strings
'***************************************************************************

Public Function GetAnnoSizeByScale( _
  sFCName As String, _
  lScale As Long) As Double
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim dSize As Double
    '++ END JWalton 2/7/2007
    
    '++ New coded added 10/21/05
    With g_pFldnames
        If StrComp(sFCName, .FCTLAcrAnno, vbTextCompare) = 0 Then
            '++ START JWalton 2/7/2007 Replaced repeated if statements with Select Case
            'Determine anno size based on scale
            Select Case lScale
              Case 120
                dSize = .AnnoSizeTLAcr120
              Case 240
                dSize = .AnnoSizeTLAcr240
              Case 360
                dSize = .AnnoSizeTLAcr360
              Case 480
                dSize = .AnnoSizeTLAcr480
              Case 600
                dSize = .AnnoSizeTLAcr600
              Case 1200
                dSize = .AnnoSizeTLAcr1200
              Case 2400
                dSize = .AnnoSizeTLAcr2400
              Case 4800
                dSize = .AnnoSizeTLAcr4800
              Case 9600
                dSize = .AnnoSizeTLAcr9600
              Case 24000
                dSize = .AnnoSizeTLAcr24000
              Case Else
                ' Default size
                dSize = 5
            End Select
            '++ END JWalton 2/7/2007
          ElseIf StrComp(sFCName, .FCTLNumAnno, vbTextCompare) = 0 Then
            '++ START JWalton 2/7/2007 Replaced repeated if statements with Select Case
            'Determine anno size based on scale
            Select Case lScale
              Case 120
                dSize = .AnnoSizeTLNum120
              Case 240
                dSize = .AnnoSizeTLNum240
              Case 360
                dSize = .AnnoSizeTLNum360
              Case 480
                dSize = .AnnoSizeTLNum480
              Case 600
                dSize = .AnnoSizeTLNum600
              Case 1200
                dSize = .AnnoSizeTLNum1200
              Case 2400
                dSize = .AnnoSizeTLNum2400
              Case 4800
                dSize = .AnnoSizeTLNum4800
              Case 9600
                dSize = .AnnoSizeTLNum9600
              Case 24000
                dSize = .AnnoSizeTLNum24000
              Case Else
                ' Default size
                dSize = 5
            End Select
          Else
            'Something not being trapped
            dSize = 10
        End If
        '++ END JWalton 2/7/2007
    End With
    '++end new code
    
    '++ START JWalton 2/7/2007
        ' Removed default to Case Else in Select Case blocks
    '++ END JWalton 2/7/2007
    
    ' Returns the funtion's new value
    GetAnnoSizeByScale = dSize

    Exit Function
ErrorHandler:
    HandleError True, _
                "GetAnnoSizeByScale " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  FileExists
'Initial Author:        <<Unknown>>
'Subsequent Author:     JWalton
'Created:               2/6/2007
'Purpose:       Determine file existence
'Called From:   frmArrows.cmdHelp_Click
'               frmCombine.cmdHelp_Click
'               frmLocate.cmdHelp_Click
'               frmMapIndex.cmdHelp_Click
'               frmTaxlotAssignment.cmdHelp_Click
'Methods:       None
'Inputs:        sPath: A string that represents the file to check
'Parameters:    None
'Outputs:       None
'Returns:       A boolean value the indicates if the file exists
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'JWalton        2/6/2007    Revised using Scripting Runtime Objects
'***************************************************************************

Public Function FileExists( _
  ByVal sPath As String) As Boolean
On Error GoTo ErrorHandler
    ' Variable declarations
    Dim pFileSystemObj As Scripting.FileSystemObject

    ' Intialize objects
    Set pFileSystemObj = New Scripting.FileSystemObject
    
    ' Returns the existence state of the file
    FileExists = pFileSystemObj.FileExists(sPath)
    
    ' Cleans up and exits
    Set pFileSystemObj = Nothing
    Exit Function
    
ErrorHandler:
    ' Returns a negative value in case of error
    FileExists = False
End Function

'***************************************************************************
'Name:                  LoadFCIntoMap
'Initial Author:        <<Unknown>>
'Subsequent Author:     JWalton
'Created:               <<Unknown>>
'Purpose:       Loads a feature class into the current map
'Called From:   basUtilities.CalcTaxlotValues
'               cmdAutoUpdate.ICommand_OnClick
'               cmdTaxlotAssignment.ICommand_OnClick
'               frmMapIndex.Form_Load
'Description:   Given a feature class, sFCName, and an alternate title,
'               sTitle.
'               Show a dialog box with title sTitle that allows the user to
'               select the personal geodatabase that sFCName resides in.
'               The feature class sFCName is then loaded from the chosen
'               personal geodatabase.
'Methods:       None
'Inputs:        sFCName - The feature class to find
'               sTitle - An alternate title for the file dialog box
'Parameters:    <<None>>
'Outputs:       <<None>>
'Returns:       Boolean -- True for loaded, False for not loaded
'Errors:        This routine raises no known errors.
'Assumptions:   <<None>>
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'JWalton        1/31/2007   Revised from Sub to Function
'***************************************************************************

Public Function LoadFCIntoMap( _
  ByVal sFCName As String, _
  Optional ByVal sTitle As String = "") As Boolean
On Error GoTo ErrorHandler
    ' Variable declarations
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    '++ START JWalton 1/31/2007 Additional variable declarations
    Dim pFileDlg As clsCatalogFileDlg
    '++ END JWalton 1/31/2007
    Dim pMXDoc As esriArcMapUI.IMxDocument
    Dim pFeatLayer As esriCarto.IFeatureLayer
    Dim pMap As esriCarto.IMap
    Dim pDataset As esriGeoDatabase.IDataset
    Dim pFC As esriGeoDatabase.IFeatureClass
    Dim pFWS As esriGeoDatabase.IFeatureWorkspace
    Dim pWS As esriGeoDatabase.IWorkspace
    '++ START JWalton 1/31/2007 Additional variable declarations
    Dim pWSFact As esriGeoDatabase.IWorkspaceFactory2
    '++ END JWalton 1/31/2007
    '++ END JWalton 2/7/2007
    
    '++ START JWalton 1/31/2007 Forces the user to choose the geodatabase where the feature class resides
    ' Initialize objects
    Set pFileDlg = New clsCatalogFileDlg
    With pFileDlg
        .AllowMultiSelect = False
        .ButtonCaption = "Select"
        If Len(sTitle) Then
            .Title = sTitle
          Else
             .Title = "Find feature class " & sFCName & "..."
        End If
        .SetFilter New esriCatalog.GxFilterPersonalGeodatabases, True, True
        .ShowOpen
    End With
    
    ' Exit if there is no selection
    If Len(pFileDlg.SelectedObject(1)) = 0 Then GoTo Process_Exit
        
    ' Initialize a workspace with the chosen personal geodatabase
    Set pWSFact = New esriDataSourcesGDB.AccessWorkspaceFactory
    Set pWS = pWSFact.OpenFromFile(pFileDlg.SelectedObject(1), 0)
    '++ END JWalton 1/31/2007
    
    ' Adds the passed feature class from the selected workspace to the map
    Set pFWS = pWS
    Set pFC = pFWS.OpenFeatureClass(sFCName)
    Set pFeatLayer = New esriCarto.FeatureLayer
    Set pFeatLayer.FeatureClass = pFC
    Set pDataset = pFC
    pFeatLayer.Name = pDataset.Name
    Set pMXDoc = g_pApp.Document
    Set pMap = pMXDoc.FocusMap
    pMap.AddLayer pFeatLayer
    pMXDoc.CurrentContentsView.Refresh 0
    
'++ START JWalton 1/31/2007
    ' Returns the value of the function and exits
Process_Exit:
    LoadFCIntoMap = True
    Exit Function
    
ErrorHandler:
    ' Returns the value of the function and exits
    LoadFCIntoMap = False
'++ END JWalton 1/31/2007
End Function

'***************************************************************************
'Name:                  IsOrMapFeature
'Initial Author:        <<Unknown>>
'Subsequent Author:     James Moore
'Created:               <<Unknown>>
'Purpose:       Determines if a feature class is part of the ORMAP design.
'Called From:   cmdAutoUpdate.m_pEditorEvents_OnDeleteFeature
'Description:   Given a geodatabase object, pObj.
'               Determine if the object is part of the ORMAP Data Design by
'               comparing the name of the dataset of pObj's feature class
'               to the name of the ORMAP Data Design feature classes.
'Methods:       None
'Inputs:        pObj - A valid initialized geodatabase object
'Parameters:    None
'Outputs:       None
'Returns:       True or False
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    1/10/2007   This function would never return true, because a
'                           return value of true had never been assigned to
'                           the function.
'                           Modified to use StrComp function instead of
'                           LCase and Trim
'John Walton    2/7/2007    Renamed pName to sName to comply with variable
'                           naming conventions
'                           Removed ESRI Error Handler in favor of returning
'                           False in the event of an error
'***************************************************************************

Public Function IsOrMapFeature( _
  pObj As esriGeoDatabase.IObject) As Boolean
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pDSet As esriGeoDatabase.IDataset
    Dim pOC As esriGeoDatabase.IObjectClass
    Dim sName As String
    '++ END JWalton 2/7/2007
    
    Set pOC = pObj.Class
    Set pDSet = pOC
'++ START JWM 01/10/2007
    sName = pDSet.Name
    
    '++ START JWalton 2/7/2007
    IsOrMapFeature = StrComp(sName, g_pFldnames.FCAnno10, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCAnno100, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCAnno20, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCAnno200, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCAnno2000, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCAnno30, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCAnno40, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCAnno400, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCAnno50, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCAnno800, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCCartoLines, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCLotsAnno, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCMapIndex, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCPlats, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCReferenceLines, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCTaxCode, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCTaxCodeAnno, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCTaxlot, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCTaxlotLines, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCTLAcrAnno, vbTextCompare) = 0 Or _
                     StrComp(sName, g_pFldnames.FCTLNumAnno, vbTextCompare) = 0
    '++ END JWalton 2/7/2007
    
'++ END JWM 01/10/2007


    Exit Function
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Return False in the event of an error
    IsOrMapFeature = False
    '++ END JWalton 2/7/2007
End Function

'***************************************************************************
'Name:                  SetAnnoSize
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Update/Initialize feature linked annotation size
'Called From:   cmdAutoUpdate.m_pEditorEvents_OnChangeFeature
'               cmdAutoUpdate.m_pEditorEvents_OnCreateFeature
'Description:   Given an object, pObj, and a feature, IFeature.
'               Determines if pObj is an annotation feature, derives the
'               map number for the Map Index polygon overlaying pFeat, and
'               resets the annotation feature size in pObj.
'Methods:       None
'Inputs:        pObj - A valid initialized geodatabase object
'               pFeat - The feature associated with the annotation
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point
'John Walton    2/7/2007    Deleted variable u as it appears to be dead
'***************************************************************************

Public Sub SetAnnoSize( _
  pObj As esriGeoDatabase.IObject, _
  pFeat As esriGeoDatabase.IFeature)
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pAnnotationElement As esriCarto.IAnnotationElement
    Dim pAnnotationFeature As esriCarto.IAnnotationFeature
    Dim pElement As esriCarto.IElement
    Dim pMIFlayer As esriCarto.IFeatureLayer
    Dim pTextElement As esriCarto.ITextElement
    Dim pTextSym As esriDisplay.ITextSymbol
    Dim pAnnoDset As esriGeoDatabase.IDataset
    Dim pAnnoFeat As esriGeoDatabase.IFeature
    Dim pMIFclass As esriGeoDatabase.IFeatureClass
    Dim pAnnoClass As esriGeoDatabase.IObjectClass
    Dim pAOC As esriGeoDatabase.IObjectClass
    Dim pEnv As esriGeometry.IEnvelope
    Dim pGeometry As esriGeometry.IGeometry
    Dim pCenter As esriGeometry.IPoint
    Dim dSize As Double
    Dim lAnnoMapNumFld As Long
    Dim lFld As Long
    Dim sMapNum As String
    Dim sMapScale As String
    Dim vVal As Variant
    '++ END JWalton 2/7/2007
    
    ' Initialize objects
    Set pAOC = pObj.Class
    Set pAnnoFeat = pObj
    
    'Capture MapNumber for each anno feature created
    lAnnoMapNumFld = LocateFields(pObj.Class, g_pFldnames.MIMapNumberFN)
    If lAnnoMapNumFld = -1 Then GoTo Process_Exit
    
    'If new anno feature with no text, determine if it has a shape
    lFld = pAnnoFeat.Fields.FindField("TextString")
    If lFld = -1 Then
        '++ START JWalton 2/8/2007 Reformatted message
        MsgBox "Unable to locate text string field in annotation class." & vbCrLf & _
               Space(20) & "Cannot set size", vbCritical
        '++ END JWalton 2/8/2007
        GoTo Process_Exit
    End If
    vVal = pAnnoFeat.Value(lFld)
    If IsNull(vVal) Then GoTo Process_Exit
        
    ' Gets the map number from the overlaying MapIndex feature class
    Set pFeat = pObj
    Set pGeometry = pFeat.Shape
    If pGeometry.IsEmpty Then GoTo Process_Exit
    Set pEnv = pGeometry.Envelope
    Set pCenter = CT_GetCenterOfEnvelope(pEnv)
    Set pMIFlayer = basUtilities.FindFeatureLayerByDS(g_pFldnames.FCMapIndex)
    If pMIFlayer Is Nothing Then GoTo Process_Exit
    Set pMIFclass = pMIFlayer.FeatureClass
    sMapNum = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapNumberFN)
    
    ' Allow existing anno to be moved without changing MapNumber
    ' Some anno will reside in another Taxlot, but labels the neighboring taxlot
    If sMapNum = pObj.Value(lAnnoMapNumFld) Then
        ' Sets the value of the annotation map number field
        pObj.Value(lAnnoMapNumFld) = sMapNum
    
        ' Update the size to reflect current mapscale
        sMapScale = GetValueViaOverlay(pCenter, pMIFclass, g_pFldnames.MIMapScaleFN)
        If IsNull(sMapScale) Then GoTo Process_Exit
        
        ' Determine which annotation class this is
        Set pAnnoClass = pObj.Class
        Set pAnnoDset = pAnnoClass
        
        'If other anno, don't continue
        If LCase(pAnnoDset.Name) <> LCase(g_pFldnames.FCTLAcrAnno) And _
           LCase(pAnnoDset.Name) <> LCase(g_pFldnames.FCTLNumAnno) Then
            GoTo Process_Exit
        End If
        
        ' Gets the size of the annotation from the scale of the annotation dataset
        dSize = basUtilities.GetAnnoSizeByScale(pAnnoDset.Name, CLng(sMapScale))
        
        ' Get the anno feature, its symbol, set the appropriate size
        Set pAnnotationFeature = pObj
        Set pAnnotationElement = pAnnotationFeature.Annotation
        Set pElement = pAnnotationElement
        Set pTextElement = pElement
        Set pTextSym = pTextElement.Symbol
        pTextSym.Size = dSize
        pTextElement.Symbol = pTextSym
        Set pElement = pTextElement
        Set pAnnotationElement = pElement
        pAnnotationFeature.Annotation = pAnnotationElement
    End If
    
Process_Exit:
    Exit Sub
  
ErrorHandler:
    HandleError True, _
                "SetAnnoSize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  LocateFields
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Find the index of a field in a feature class
'Called From:   basUtilities.CalcTaxlotValues
'               basUtilities.GetMapSufNum
'               basUtilities.GetMapSufType
'               basUtilities.GetSpecialInterests
'               basUtilities.SetAnnoSize
'               cmdArrows.GenerateHooks
'               cmdArrows.ITool_OnMouseDown
'               cmdAutoUpdate.m_pEditorEvents_OnCreateFeature
'               cmdCombine.cmdApply_Click
'               cmdTaxlotAssignment.ICommand_OnClick
'Description:   Given a feature class, pFClass, and a field name, sFldName.
'               Find the index of the field, and return either it, or -1,
'               in the case that it is not found, to.
'Methods:       This function may return zero because that is a valid index,
'               but -1 is not.
'               The return value of -1 means the field was not found.
'Inputs:        pFClass - The feature class to locate a field in
'               sFldName - The name of the field to find
'Parameters:    None
'Outputs:       None
'Returns:       Index of field or -1
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'james moore    12-19-06    Will now return the result of the FindField call
'                           so that the user can make a decision base on the
'                           result.
'John Walton    1/7/2007    Removed message to user regarding field not
'                           found
'                           Removed ESRI Error Handler in favor of returning
'                           -1 in the event of an error
'***************************************************************************

Public Function LocateFields( _
  pFClass As esriGeoDatabase.IFeatureClass, _
  sFldName As String) As Long
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007
    LocateFields = pFClass.Fields.FindField(sFldName)
    '++ END JWalton 2/7/2007

    Exit Function
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Return -1 in the event of an error
    LocateFields = -1
    '++ START JWalton 2/7/2007
End Function

'***************************************************************************
'Name:                  UpdateAutoFields
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Update environment constant fields
'Called From:   cmdAutoUpdate.m_pEditorEvents_OnChangeFeature
'               cmdAutoUpdate.m_pEditorEvents_OnCreateFeature
'Description:   Give a feature, pFeat.
'               Update the AutoWho and the AutoDate fields with the current
'               username and date/time, respectively, of pFeat.
'Methods:       None
'Inputs:        pFeat - A feature
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

Public Sub UpdateAutoFields( _
  pFeat As esriGeoDatabase.IFeature)
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim lAutoDateFld As Long
    Dim lAutoWhoFld As Long
    '++ END JWalton 2/7/2007
    
    '++ START JWalton 2/8/2007
    If pFeat Is Nothing Then GoTo Process_Exit
    '++ END JWalton 2/8/2007
    
    ' Populate the AutoDate field
    lAutoDateFld = pFeat.Fields.FindField(g_pFldnames.AutoDateFN)
    If lAutoDateFld > -1 Then
        pFeat.Value(lAutoDateFld) = Now
    End If
    
    ' Populate the AutoWho field
    lAutoWhoFld = pFeat.Fields.FindField(g_pFldnames.AutoWhoFN)
    If lAutoWhoFld > -1 Then
'++ START JWalton 1/26/2007 Replaced Environ, which may err in Windows XP with a Windows call that will not
        pFeat.Value(lAutoWhoFld) = UserName
'++ END JWalton 1/26/2007
    End If

Process_Exit:
    Exit Sub
    
ErrorHandler:
    HandleError True, _
                "UpdateAutoFields " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Sub

'***************************************************************************
'Name:                  GetMapSufType
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               10/11/2006
'Purpose:       Validate and format a map suffix type
'Called From:   basUtilities.CalcOMTLNum
'               cmdTaxlotAssignment.ITool_OnMouseDown
'Description:   Given a feature, pFeature.
'               Retrieves the map suffix type from pFeature, validates it,
'               and formats it by the mask '0'.
'Methods:       None
'Inputs:        pFeature - An object that supports the IFeature interface
'Parameters:    None
'Outputs:       None
'Returns:       A string the represents a properly formatted Map Suffix
'               Type
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Initial creation of this comment section
'JWalton        2/7/2007    Removed ESRI Error Handler in favor or returning
'                           the default string '0'.
'***************************************************************************

Public Function GetMapSufType( _
  pFeature As esriGeoDatabase.IFeature) As String
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim lTLMapSufTypeFld As Long
    Dim sTLMapSufTypeVal As String
    '++ END JWalton 2/7/2007
    
    ' Finds the map suffix type field and formats the entry
    lTLMapSufTypeFld = LocateFields(pFeature.Class, g_pFldnames.TLSufTypeFN)
    If lTLMapSufTypeFld = -1 Then
        sTLMapSufTypeVal = "0"
      Else
        If Not IsNull(pFeature.Value(lTLMapSufTypeFld)) Then
            sTLMapSufTypeVal = pFeature.Value(lTLMapSufTypeFld)
          Else
            sTLMapSufTypeVal = "0"
        End If
        'Verify that it is 1 digit
        If Len(sTLMapSufTypeVal) < 1 Then
            Do Until Len(sTLMapSufTypeVal) = 1
                sTLMapSufTypeVal = "0" & sTLMapSufTypeVal
            Loop
        End If
        
        'Verify that it isn't more than 1 digit
        If Len(sTLMapSufTypeVal) > 1 Then
            sTLMapSufTypeVal = Left(sTLMapSufTypeVal, 1)
        End If
    End If
    
    ' Return's the function's value
    GetMapSufType = sTLMapSufTypeVal

    Exit Function
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Returns the function's value
    GetMapSufType = "0"
    '++ END JWalton 2/7/2007
End Function
'++ END, Laura Gordon, November 29, 2005

'***************************************************************************
'Name:                  GetMapSufNum
'Initial Author:        Laura Gordon
'Subsequent Author:     <<Type your name here>>
'Created:               November 29, 2005
'Purpose:       Validate and format a map suffix number
'Called From:   cmdTaxlotAssignment.ITool_OnMouseDown
'Description:   Given a feature, pFeature.
'               Retrieves the map suffix number from pFeature, validates it,
'               and formats it by the mask '000'.
'Methods:       None
'Inputs:        pFeature - An object that supports the IFeature interface
'Parameters:    None
'Outputs:       None
'Returns:       A string the represents a properly formatted Map Suffix
'               Number
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'JWalton        2/7/2007    Removed ESRI Error Handler in favor or returning
'                           the default string '000'.
'***************************************************************************

Public Function GetMapSufNum( _
  pFeature As esriGeoDatabase.IFeature) As String
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim lTLMapSufNumFld As Long
    Dim sTLMapSufNumVal As String
    '++ END JWalton 2/7/2007
    
    ' Finds the map suffix number field and formats the entry
    lTLMapSufNumFld = LocateFields(pFeature.Class, g_pFldnames.TLSufNumFN)
    If lTLMapSufNumFld = -1 Then
        sTLMapSufNumVal = "000"
      Else
        If Not IsNull(pFeature.Value(lTLMapSufNumFld)) Then
            sTLMapSufNumVal = pFeature.Value(lTLMapSufNumFld)
          Else
            sTLMapSufNumVal = "000"
        End If
        'Verify that it is 3 digit
        If Len(sTLMapSufNumVal) < 3 Then
            Do Until Len(sTLMapSufNumVal) = 3
                sTLMapSufNumVal = "0" & sTLMapSufNumVal
            Loop
        End If
        
        'Verify that it isn't more than 3 digits
        If Len(sTLMapSufNumVal) > 3 Then
            sTLMapSufNumVal = Left(sTLMapSufNumVal, 3)
        End If
    End If
    
    ' Returns the function's value
    GetMapSufNum = sTLMapSufNumVal

    Exit Function
ErrorHandler:
    '++ START JWalton 2/7/2007
    ' Returns the function's value
    GetMapSufNum = "000"
    '++ END JWalton 2/7/2007
End Function

'***************************************************************************
'Name:                  CalcOMTLNum
'Initial Author:        <<Unknown>>
'Subsequent Author:     Type your name here.
'Created:               <<Unknown>>
'Purpose:       Calculate ORMAP Taxlot Number when one if its components has
'               changed
'Called From:   basUtilities.CalcTaxlotValues
'Description    Given an ORMAP Number, sExistOMNum, and feature, pFeat, and
'               a taxlot value, sTLVal.
'               Remove the existing map suffix type and number from
'               sExistOMNum and replace them with the new values in pFeat and
'               append sTLVal to form the return value.
'Methods:       None
'Inputs:        sExistOMNum - An ORMAP Number
'               pFeat - An object that supports the IFeature interface
'               sTLVal - A taxlot number
'Parameters:    None
'Outputs:       None
'Returns:       A string that represents an ORMAP number updated with the
'               value from pFeature and sTLVal.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/7/2007    Removed variable sOMTLNval as it is no longer
'                           necessary
'***************************************************************************

Public Function CalcOMTLNum( _
  ByVal sExistOMNum As String, _
  ByRef pFeat As IFeature, _
  ByVal sTLVal As String) As String
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim sShortOMNum As String
    '++ END JWalton 2/7/2007
    
    '++ START Laura Gordon, November 29, 2005
    Dim sTLMapSufNumVal As String
    Dim sTLMapSufTypeVal As String
    '++ END Laura Gordon, November 29, 2005
    

    ' Gets the Ormap number with the map suffix type or number
    sShortOMNum = ShortenOMMapNum(sExistOMNum)
    
    '++ START Laura Gordon, November 29, 2005
    ' Gets the new map suffix type and number
    sTLMapSufTypeVal = GetMapSufType(pFeat)
    sTLMapSufNumVal = GetMapSufNum(pFeat)
    '++ END Laura Gordon, Novemeber 29, 2005
    
    '++ START JWalton 2/7/2007
    ' Recreate and return the ORMAP Taxlot number
    CalcOMTLNum = sShortOMNum & sTLMapSufTypeVal & sTLMapSufNumVal & sTLVal
    '++ END JWalton 2/7/2007

    Exit Function
ErrorHandler:
    HandleError True, _
                "CalcOMTLNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  ShortenOMTLNum
'Initial Author:        <<Unknown>>
'Subsequent Author:     Type your name here.
'Created:               <<Unknown>>
'Purpose:       Return a short ORMAP Number that is minus the final two
'               component -- Map Suffix Type, and Number
'Called From:   basUtilities.CalcOMTLNum
'               cmdTaxlotAssignment.ITool_OnMouseDown
'Description:   Given an ORMAP Number, sOMVal.
'               Truncate sOMVal to 20 characters to obtain the short number
'Methods:       None
'Inputs:        sOMVal - An ORMAP Number
'Parameters:    None
'Outputs:       None
'Returns:       A string representing the short number
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************

Public Function ShortenOMMapNum( _
  sOMVal As String) As String
On Error GoTo ErrorHandler

    'Remove two values from the ORMAPMap number for the purpose of populating ORMAPTaxlot
    ShortenOMMapNum = Left(sOMVal, 20)

    Exit Function
ErrorHandler:
    HandleError True, _
                "ShortenOMMapNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  ZoomToExtent
'Initial Author:        <<Unknown>>
'Subsequent Author:     Type your name here.
'Created:               <<Unknown>>
'Purpose:       Zooms the current extent to the passed in envelope (i.e.
'               zoom to feature).
'               Works for Layout and Data view
'Called From:   frmLocate.cmdApply_Click
'Description:   Given an envelope, pEnv, and an ArcMap Document, pMXDoc.
'               Update the extent of the active view of the active map in
'               pMXDoc
'Methods:       None
'Inputs:        pEnv - An envelope
'               pMXDoc - An ArcMap Document
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

Public Sub ZoomToExtent( _
  ByRef pEnv As IEnvelope, _
  ByRef pMXDoc As IMxDocument)
     '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pMap As esriCarto.IMap
    Dim pActiveView As esriCarto.IActiveView
    '++ END JWalton 2/7/2007
    
    ' Gets a reference to the current view window
    Set pMap = pMXDoc.FocusMap
    Set pActiveView = pMap

    ' Updates the view's extent
    pActiveView.Extent = pEnv
    pActiveView.Refresh
End Sub


'***************************************************************************
'Name:                  gsb_StartDoc
'Initial Author:        James Moore
'Subsequent Author:     <<Type your name here>>
'Created:               10/16/2006
'Purpose:       Opens a document with its associated application.
'Called From:   frmArrows.cmdHelp_Click
'               frmCombine.cmdHelp_Click
'               frmLocate.cmdHelp_Click
'               frmMapIndex.cmdHelp_Click
'               frmTaxlotAssignment.cmdHelp_Click
'Description:   You can use the Windows API ShellExecute() function to start
'               the application associated with a given document extension
'               without knowing the name of the associated application.
'Methods:
'Parameters:    alWindowHandle: handle to calling form.
'               asDocname: fully qualified path to document (including file name)
'Outputs:       None
'Returns:       None
'Errors:        This routine use the return value from the API call to see
'               if there was an error, if so an appropriate message will be
'               displayed in a message box.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/16/2006  Initial creation of this comment section
'***************************************************************************

Public Sub gsb_StartDoc( _
  ByRef alWindowHandle As Long, _
  ByRef asDocname As String)
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim lResult As Long
    Dim sMsg As String
    '++ END JWalton 2/7/2007

    lResult = ShellExecute(alWindowHandle, "Open", asDocname, "", "C:\", SW_SHOWNORMAL)
    If lResult <= 32 Then
        'There was an error
        Select Case lResult
            Case SE_ERR_FNF
                sMsg = "File not found"
            Case SE_ERR_PNF
               sMsg = "Path not found"
            Case SE_ERR_ACCESSDENIED
                sMsg = "Access denied"
            Case SE_ERR_OOM
                sMsg = "Out of memory"
            Case SE_ERR_DLLNOTFOUND
                sMsg = "DLL not found"
            Case SE_ERR_SHARE
                sMsg = "A sharing violation occurred"
            Case SE_ERR_ASSOCINCOMPLETE
                sMsg = "Incomplete or invalid file association"
            Case SE_ERR_DDETIMEOUT
                sMsg = "DDE Time out"
            Case SE_ERR_DDEFAIL
                sMsg = "DDE transaction failed"
            Case SE_ERR_DDEBUSY
                sMsg = "DDE busy"
            Case SE_ERR_NOASSOC
                sMsg = "No association for file extension"
            Case ERROR_BAD_FORMAT
                sMsg = "Invalid EXE file or error in EXE image"
            Case Else
                sMsg = "Unknown error"
        End Select
        MsgBox sMsg, vbOKOnly + vbExclamation, "Error opening file"
    End If
End Sub

'***************************************************************************
'Name:                  gfn_s_CreateMapTaxlotValue
'Initial Author:        James Moore
'Subsequent Author:     <<Type your name here>>
'Created:               9-23-2005
'Purpose:       Use the ORMapTaxlot value to create a MapTaxlot value based
'               on the mask from the ini file.
'Called From:   basUtilities.CalcTaxlotValues
'               basUtilities.gsb_StartDoc
'               cmdTaxlotAssignment.ITool_OnMouseDown
'               frmMapIndex.UpdateTaxlots
'Methods:       The string parsing procedures depends on a valid ORTaxlot
'               string 29 characters long as defined in version 1.3 of the
'               ORMAP data structure.
'               Extensive use of the Mid function causes heavy reliance on
'               the position of values in the string
'               Need to find a better way to handle half townships and
'               ranges
'               May want to change this function to use Regular Expressions
'Inputs:        as_ORMapTaxlotString: The ORTaxlot value
'               as_MaskFormatString: the formatting string
'Parameters:    None
'Outputs:       None
'Returns:       A formatted string that can be used as parcel ID or/and as
'               a MapTaxlot value
'Errors:        This routine raises no known errors.
'Assumptions:   A valid ORTaxlot string and Mask value is passed in.
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    9-23-2005   Initial creation of this routine
'James Moore    11-8-06     Adding special case for County that uses a Q to
'                           store half ranges
'John Walton    2/7/2007    Renamed variables to conform to variable naming
'                           conventions
'***************************************************************************

Public Function gfn_s_CreateMapTaxlotValue( _
  ByVal as_ORMapTaxlotString As String, _
  ByRef as_MaskFormatString As String) As String
On Error GoTo gfn_s_CreateMapTaxlotValue_Error
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim bProcessedParcelID As Boolean             ' Flag
    Dim bProcessedRangeFractional As Boolean
    Dim bHasAlphaQtr As Boolean                   ' Flag for processing the mask
    Dim bHasAlphaQtrQtr As Boolean                ' Flag for processing the mask
    Dim bHasTownPart As Boolean                   ' Flag for processing the mask
    Dim bHasRangePart As Boolean                  ' Flag for processing the mask
    Dim bProcessedTownFractional As Boolean
    Dim i As Integer
    Dim iCharCode As Integer
    Dim iPosCharMaskForward As Integer            'Marks the current postion in the mask array
    Dim iMaskLength As Integer
    Dim iCountyCode As Integer
    Dim lMaskTokenCount As Long                   ' How many characters in the mask
    Dim sArr_MaskValues() As String
    Dim sCurrORMapNumValue As String              ' To hold a char from ORMAP string
    Dim sFormattedString As String                ' The result of our work
    Dim sMaskToApply As String
    Dim sPrevCharInMaskArray As String            ' To use as check for character position
    Dim sTemp As String
    '++ END JWalton 2/7/2007
    
    If Len(as_ORMapTaxlotString) = 0 Or Len(as_MaskFormatString) = 0 Then
        GoTo ProcessExit
    End If
    
    iCountyCode = CInt(Left$(as_ORMapTaxlotString, 2))
    
    ' flag for half townships,ranges
    bHasTownPart = (Val(Mid$(as_ORMapTaxlotString, 5, 3)) > 0)
    bHasRangePart = (Val(Mid$(as_ORMapTaxlotString, 11, 3)) > 0)
    
    'set flags for section qtrs
    Select Case iCountyCode
        Case 1 To 19, 21 To 36
            bHasAlphaQtr = Not IsNumeric(Mid$(as_ORMapTaxlotString, 17, 1))
            bHasAlphaQtrQtr = Not IsNumeric(Mid$(as_ORMapTaxlotString, 18, 1))
        Case 20 'lane county uses a totally numeric identifier for qtrs of sections with zeros as placeholders
            bHasAlphaQtr = False
            bHasAlphaQtrQtr = False
    End Select
    
    'We must adjust the mask for clackamas county if there are no  half ranges in the current string
    If InStr(Mid$(as_MaskFormatString, 2, 6), "^") > 0 Then
        If bHasRangePart = False Then
            sMaskToApply = Replace(as_MaskFormatString, "^", vbNullString) 'remove this character
        Else 'if there is a range part the letter Q will be  placed in the position where D sits
            sMaskToApply = Replace(as_MaskFormatString, "D", vbNullString)
        End If
    Else
        sMaskToApply = as_MaskFormatString
    End If
    
    iMaskLength = Len(sMaskToApply)
    
    'Dimension the mask array and fill each position with a character from the mask
    ' I am using an array that begins at dimension one for ease of use
    ReDim sArr_MaskValues(1 To iMaskLength) As String
    
    For i = 1 To iMaskLength
        sArr_MaskValues(i) = UCase$(Mid$(sMaskToApply, i, 1))
    Next i
    
    ' Create a string of spaces to place our results in. This helps a speed up string manipulation a little.
    sFormattedString = Space(iMaskLength)
    
    For i = 1 To UBound(sArr_MaskValues)
        ' Increment our position in the mask
        iPosCharMaskForward = InStr(i, sMaskToApply, sArr_MaskValues(i), vbTextCompare)
        iCharCode = Asc(sArr_MaskValues(i)) 'the ascii value of the character
       
        ' Returns how many of these characters appear in the mask, AND when used in
        ' Mid function gets/sets that many chars
        lMaskTokenCount = gfn_l_CountTokens(UCase$(sMaskToApply), sArr_MaskValues(i))
        
        Select Case iCharCode
            Case 68 '"D"
                If StrComp(sPrevCharInMaskArray, "^", vbTextCompare) = 0 Then 'for clackamas county which uses a Q for halfs
                    If StrComp(Mid$(sMaskToApply, iPosCharMaskForward - 2, 1), "T", vbTextCompare) = 0 Then ' TOWNSHIP DIRECTION
                        Mid$(sFormattedString, iPosCharMaskForward, 1) = Mid$(as_ORMapTaxlotString, 8, 1)
                    ElseIf StrComp(Mid$(sMaskToApply, iPosCharMaskForward - 2, 1), "R", vbTextCompare) = 0 Then 'RANGE DIRECTION
                        Mid$(sFormattedString, iPosCharMaskForward, 1) = Mid$(as_ORMapTaxlotString, 14, 1)
                    End If
                Else
                    If StrComp(sPrevCharInMaskArray, "T", vbTextCompare) = 0 Then ' TOWNSHIP DIRECTION
                        Mid$(sFormattedString, iPosCharMaskForward, 1) = Mid$(as_ORMapTaxlotString, 8, 1)
                    ElseIf StrComp(sPrevCharInMaskArray, "R", vbTextCompare) = 0 Then 'RANGE DIRECTION
                        Mid$(sFormattedString, iPosCharMaskForward, 1) = Mid$(as_ORMapTaxlotString, 14, 1)
                    End If
                End If
            'Formats for the parcel id
            Case 64 '"@"
                If Not bProcessedParcelID Then
                    Mid$(sFormattedString, iPosCharMaskForward) = ffn_s_CreateParcelID(Mid$(as_ORMapTaxlotString, 25, 5), Mid$(sMaskToApply, iPosCharMaskForward, lMaskTokenCount))
                    bProcessedParcelID = True
                End If
            Case 38 '"&" 'Using these characters in mask will strip leading zeros from parcel id
                If Not bProcessedParcelID Then
                    sTemp = ffn_s_CreateParcelID(Mid$(as_ORMapTaxlotString, 25, 5), Mid$(sMaskToApply, iPosCharMaskForward, lMaskTokenCount))
                    Mid$(sFormattedString, iPosCharMaskForward) = ffn_s_StripLeadingZeros(sTemp)
                    bProcessedParcelID = True
                End If
            'QUARTER and QUARTER QUARTER
            Case 81 '"Q"
                If StrComp(sPrevCharInMaskArray, "Q", vbTextCompare) = 0 Then ' Quarter Quarter
                    If bHasAlphaQtrQtr Then
                        Mid$(sFormattedString, iPosCharMaskForward, 1) = Mid$(as_ORMapTaxlotString, 18, 1)
                    Else
                        sCurrORMapNumValue = UCase$(Mid$(as_ORMapTaxlotString, 18, 1))
                        If sCurrORMapNumValue Like "[A-D]" Then
                            Mid$(sFormattedString, iPosCharMaskForward, 1) = Switch(sCurrORMapNumValue = "A", 1, sCurrORMapNumValue = "B", 2, sCurrORMapNumValue = "C", 3, sCurrORMapNumValue = "D", 4)
                        Else
                            Mid$(sFormattedString, iPosCharMaskForward, 1) = Chr$(48) 'ZERO
                        End If
                    End If
                Else ' Quarter
                    If bHasAlphaQtr Then
                        Mid$(sFormattedString, iPosCharMaskForward, 1) = Mid$(as_ORMapTaxlotString, 17, 1)
                    Else
                        sCurrORMapNumValue = UCase$(Mid$(as_ORMapTaxlotString, 17, 1))
                        If sCurrORMapNumValue Like "[A-D]" Then
                            Mid$(sFormattedString, iPosCharMaskForward, 1) = Switch(sCurrORMapNumValue = "A", 1, sCurrORMapNumValue = "B", 2, sCurrORMapNumValue = "C", 3, sCurrORMapNumValue = "D", 4)
                        Else
                            Mid$(sFormattedString, iPosCharMaskForward, 1) = Chr$(48) 'ZERO
                        End If
                    End If
                End If
            'Range
            Case 82 '"R"
                If StrComp(sPrevCharInMaskArray, "R", vbTextCompare) <> 0 Then
                    If lMaskTokenCount > 1 Then
                        Mid$(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid$(as_ORMapTaxlotString, 9, lMaskTokenCount)
                    Else
                        Mid$(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid$(as_ORMapTaxlotString, 10, lMaskTokenCount)
                    End If
                End If
            'SECTION
            Case 83 '"S"
                If StrComp(sPrevCharInMaskArray, "S", vbTextCompare) = 0 Then 'SECOND pos
                    Mid$(sFormattedString, iPosCharMaskForward, 1) = Mid$(as_ORMapTaxlotString, 16, lMaskTokenCount)
                Else 'FIRST POS
                    Mid$(sFormattedString, iPosCharMaskForward, 1) = Mid$(as_ORMapTaxlotString, 15, lMaskTokenCount)
                End If
            'Township
            Case 84 '"T"
                If StrComp(sPrevCharInMaskArray, "T", vbTextCompare) <> 0 Then
                    If lMaskTokenCount > 1 Then
                        Mid$(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid$(as_ORMapTaxlotString, 3, lMaskTokenCount)
                    Else
                        Mid$(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid$(as_ORMapTaxlotString, 4, lMaskTokenCount)
                    End If
                End If
            ' Fractional parts of a township
            Case 80 '"P"
                If StrComp(sPrevCharInMaskArray, "T", vbTextCompare) = 0 Then
                    If Not bProcessedRangeFractional Then
                        Mid$(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid$(as_ORMapTaxlotString, 11, lMaskTokenCount)
                        bProcessedRangeFractional = True
                    End If
                ElseIf StrComp(sPrevCharInMaskArray, "R", vbTextCompare) = 0 Then
                    If Not bProcessedTownFractional Then
                        Mid$(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid$(as_ORMapTaxlotString, 5, lMaskTokenCount)
                        bProcessedTownFractional = True
                    End If
                End If
            Case 94 '^ special case for Clackamas county
                If StrComp(sPrevCharInMaskArray, "R", vbTextCompare) = 0 Then
                    If bHasRangePart Then
                        Mid$(sFormattedString, iPosCharMaskForward, 1) = Chr$(81) 'Q
                    End If
                ElseIf StrComp(sPrevCharInMaskArray, "T", vbTextCompare) = 0 Then 'fractional part of township
                    If bHasTownPart Then
                        Mid$(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Chr$(81)
                    End If
                End If
        End Select
        
        sPrevCharInMaskArray = sArr_MaskValues(i)
        
    Next i
    
    ' Returns the value of the function
    gfn_s_CreateMapTaxlotValue = Trim$(sFormattedString)

ProcessExit:
    Exit Function

gfn_s_CreateMapTaxlotValue_Error:
    HandleError True, _
                "gfn_s_CreateMapTaxlotValue " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  gfn_l_CountTokens
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Given a string of token characters and a single character
'               tokento search for, the number of tokens in the string will
'               be returned. This function is useful for dimensioning an
'               array to store the delimited items.
'Called From:   gfn_s_CreateMapTaxlotValue
'Description:   Given a string to search in, as_Source, and a character to
'               count, as_Token.
'               Counts and return the number of times that as_Token occurs
'               in as_Source
'Method:        This function uses Unicode representation of characters
'Inputs:        as_Source: A list of tokens
'               as_Token:  The character token to search for.
'Parameters:    None
'Outputs:       None
'Returns:       The number of tokens in sSource. If PsSource is empty or
'               there is no token to count, 0 is returned
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Developer:     Date:           Comments:
'----------     ------      ---------
'James Moore    9/23/2005   Initial creation of this routine
'John Walton    2/7/2007    Renamed variables to conform to variable naming
'                           conventions
'                           Added error handler
'***************************************************************************

 Public Function gfn_l_CountTokens( _
   ByVal as_Source As String, _
   ByRef as_Token As String) As Long
On Error GoTo Err_Handler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim byCharArray() As Byte
    Dim i As Long
    Dim lCount As Long
    Dim lUnicodeValue As Long
    '++ END JWalton 2/7/2007
    
    '++ START JWalton 2/7/2007
        ' Rewrote If...Else...End If as If...End If using initialized value
        ' lCount = 0 if no source or token is specified
    '++ END JWalton 2/7/2007
    If Len(as_Source) > 0 And Len(as_Token) > 0 Then
        byCharArray() = as_Source 'this assignment creates a unicode character array
        lUnicodeValue = AscW(as_Token) 'The AscW() function returns the Unicode character code
        For i = 0 To UBound(byCharArray()) Step 2 'this is a Unicode byte array so we must step by 2
            ' If this is the char, increase the counter
            If byCharArray()(i) = lUnicodeValue Then lCount = lCount + 1
        Next i
    End If
    
'++ START JWalton 2/7/2007
    ' Returns the value of the function
    gfn_l_CountTokens = lCount

    
    Exit Function
Err_Handler:
    ' Return 0 in the case of an error
    gfn_l_CountTokens = 0
'++ END JWalton 2/7/2007
End Function
   
'***************************************************************************
'Name:                  ffn_s_CreateParcelID
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Create a parcel ID from a mask.
'Called From:   basUtilities.gfn_s_CreateMapTaxlotValue
'Description:
'Method:        I use the Format function with user-defined string formats
'               which consist of either all (@) characters or all ampersands
'               (&)
'Inputs:        the value to mask and a mask
'Parameters:
'Outputs:
'Returns:       If a value is passed in that is not numeric then just pass
'               it straight through else return a parcel id with or without
'               leading zeros
'Errors:        This routine raises no known errors
'Assumptions:   That the mask will be either all ampersands or @ characters
'Developer:     Date:           Comments:
'----------     ------      ---------
'James Moore    9/23/2005   Initial creation of this routine
'John Walton    2/7/2007    Renamed variables to conform to variable naming
'                           conventions
'***************************************************************************

Private Function ffn_s_CreateParcelID( _
  ByRef as_ValueToMask As String, _
  ByRef as_MaskToApply As String) As String
On Error GoTo ffn_s_CreateParcelID_Error
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim sTemp As String
    '++ END JWalton 2/7/2007

    If Len(as_MaskToApply) = 0 Or Len(as_ValueToMask) = 0 Then
        GoTo ProcessExit
    End If

    sTemp = Space(Len(as_MaskToApply))
    'add exclamation point to mask so that the string will be formatted left to right
    as_MaskToApply = "!" & as_MaskToApply
    If IsNumeric(as_ValueToMask) Then
        sTemp = Format$(as_ValueToMask, as_MaskToApply)
    Else
        sTemp = as_ValueToMask
    End If
    ffn_s_CreateParcelID = sTemp
    
ProcessExit:
    Exit Function

ffn_s_CreateParcelID_Error:
    HandleError True, _
                "ffn_s_CreateParcelID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  ffn_s_StripLeadingZeros
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Remove leading zeros from a string
'Called From:   basUtilities.gfn_s_CreateMapTaxlotValue
'Description:
'Methods:
'Inputs:        sStringToParse: A string that may have leading zeros
'Parameters:
'Outputs:
'Returns:       A string of same length with blank spaces instead of
'               leading zeros.
'Errors:        This routine raises no known errors
'Assumptions:   A string with leading zeros may not be passed in.
'               In that case the whole string will be returned
'Developer:     Date:           Comments:
'----------     ------      ---------
'James Moore    9/23/2005   Initial creation of this routine
'John Walton    2/7/2007    Renamed variables to conform to variable naming
'                           conventions
'***************************************************************************

Private Function ffn_s_StripLeadingZeros( _
  ByRef sStringToParse As String) As String
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim lInputCharCount As Long
    Dim lCounter As Long
    Dim sChar As String, sTemp As String
    '++ END JWalton 2/7/2007
    
    lInputCharCount = Len(sStringToParse)
    sTemp = Space(lInputCharCount) 'create string of same length
    
    For lCounter = 1 To lInputCharCount
        sChar = Mid$(sStringToParse, lCounter, 1)
        If InStr(1, "0", sChar, vbTextCompare) < 1 Then 'go past all leading zeros
           Mid$(sTemp, lCounter) = Mid$(sStringToParse, lCounter) 'get all remaing chars
           Exit For 'and exit
        End If
    Next lCounter
    ffn_s_StripLeadingZeros = sTemp ' do not trim off leading spaces
End Function

'***************************************************************************
'Name:                 UserName
'Initial Author:        John Walton
'Subsequent Author:     <<Type your name here>>
'Created:               2/7/2007
'Purpose:       Return the current logged on user's UserName
'Called From:   basUtilities.UpdateAutoFields
'Description:   Calls the GetUserName function from Windows to get the
'               currently logged on user's name
'Methods:       This function replaces the user of Environ$ in this module.
'               Environ$ sometimes causes errors in Windows XP, but this
'               function does not.
'Inputs:        <<None>>
'Parameters:    None
'Outputs:       None
'Returns:       User's username as a string
'Errors:        This routine raises no known errors
'Assumptions:   <<None>>
'Developer:     Date:           Comments:
'----------     ----------      ----------
' JWalton        1/26/2007       Initial creation of this function
'***************************************************************************

Public Function UserName() As String
    ' Variable declarations
    Dim strBuffer As String * UNLEN_MAX
    Dim nBuffer As Long
    Dim bResult As Boolean
        
    ' Initialize buffer
    nBuffer = UNLEN
    
    ' Retrieves the user's username from Windows
    bResult = GetUserName(strBuffer, nBuffer)
    
    ' Returns the result
    If bResult Then
        UserName = Left(strBuffer, nBuffer - 1)
    End If
End Function

Public Function gfn_s_GetWindowsTempPath() As String
'************************************************************
'Name: gfn_s_GetWindowsTempPath
'Purpose: returns the string value for the windoze temp path
'Method: Some users may define their temp directory in the environment variable.
'In WinXP each user has their own temp directory to write to because they may
'not have permission to write to C:\temp or c:\windoze\temp.
'i.e. C:\Documents and Settings\[USERNAME]\Local Settings\Temp.
'This function DOES NOT add a slash at the end of the string
'Inputs: None
'Outputs: The windows temp path as a string
'Assumptions: Every system has a temp directory
'Errors:None known
'Developer: James Moore
'Date: 02/08/2000
'Revisions:
'************************************************************
    Dim lStrLen     As Long
    Dim sOutPath  As String
    
    sOutPath = String$(UNLEN, vbNullChar)
    
    lStrLen = GetTempPath(UNLEN, sOutPath)
    
    gfn_s_GetWindowsTempPath = Left$(sOutPath, lStrLen)
    
End Function
'#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#
'#                                                                         #
'# All functions after this point are not used in the DLL; that is nothing #
'# call any one of them.                                                   #
'#                                                                         #
'#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#

 '***************************************************************************
'Name:                  CompareAndSaveValue
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Unknown>>
'Created:               <<Unknown>>
'Description:   Compare the descriptive value in the GUI to the original
'               descriptive value
'Purpose:       Return an object that indicates the status
'               (changed/unchanged) of this row
'Called From:   No calls are made to this method
'Description:   Compares a value, vValNew, in a field, sFldName, in a row,
'               pRow, and changes the value in pRow if they are not equal.
'Methods:
'Inputs:        pRow - A table row object.
'               sFldName - A field in the table row object.
'               vValNew - The new object.
'               pRowChange - An object that indicates a change.
'Parameters:    None
'Outputs:       pRowChange
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    11-1-06     Appears to be a dead routine
'John Walton    2/7/06      Interchanged names for sgValNew and sValNew to
'                           conform to variable naming conventions
'                           Renamed iValNew to lValNew to conform to
'                           variable naming conventions
'                           Renamed pFldName to sFieldName in accordance
'                           with variable naming conventions
'***************************************************************************

Public Sub CompareAndSaveValue( _
  ByRef pRow As esriGeoDatabase.IRow, _
  ByVal sFldName As String, _
  ByVal vValNew As Variant, _
  ByRef pRowChanged As clsRowChanged)
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pFldType As esriGeoDatabase.esriFieldType
    Dim dtValNew As Date
    Dim dValNew As Double
    Dim lValNew As Long
    Dim lFld As Long
    Dim sgValNew As Single
    Dim sValNew As String
    Dim vValOrg As Variant
    '++ END JWalton 2/7/2007

    ' Gets the original value of the field
    vValOrg = ReadValue(pRow, sFldName)
    
    ' Tests for a difference between what is and what was
    If vValNew <> vValOrg Then
        'Get the Code value that is to be stored in the db
        vValNew = ConvertCode(pRow.Fields, sFldName, vValNew)
        
        'If the value is changed, update the row
        lFld = pRow.Fields.FindField(sFldName)
        If lFld > -1 Then
            pFldType = pRow.Fields.Field(lFld).Type
            If pFldType = esriGeoDatabase.esriFieldType.esriFieldTypeDouble Then
                ' Double Field Type
                If IsNumeric(vValNew) Then dValNew = CDbl(vValNew)
                If dValNew <> vValOrg Then
                    pRow.Value(lFld) = dValNew
                    pRowChanged.RowChanged = True
                End If
              ElseIf pFldType = esriGeoDatabase.esriFieldType.esriFieldTypeInteger Or _
               pFldType = esriGeoDatabase.esriFieldType.esriFieldTypeSmallInteger Then
                ' Integer or Long Field Type
                If IsNumeric(vValNew) Then lValNew = CLng(vValNew)
                If lValNew <> vValOrg Then
                    pRow.Value(lFld) = lValNew
                    pRowChanged.RowChanged = True
                End If
              ElseIf pFldType = esriGeoDatabase.esriFieldType.esriFieldTypeSingle Then
                ' Single Field Type
                If IsNumeric(vValNew) Then sgValNew = CSng(vValNew)
                If sgValNew <> vValOrg Then
                    pRow.Value(lFld) = sgValNew
                    pRowChanged.RowChanged = True
                End If
              ElseIf pFldType = esriGeoDatabase.esriFieldType.esriFieldTypeDate Then
                ' Date Field Type
                If IsDate(vValNew) Then dtValNew = CDate(vValNew)
                If dtValNew <> vValOrg Then
                    pRow.Value(lFld) = dtValNew
                    pRowChanged.RowChanged = True
                End If
              ElseIf pFldType = esriGeoDatabase.esriFieldType.esriFieldTypeString Then
                ' String Field Type
                sValNew = vValNew
                If sValNew <> vValOrg Then
                    pRow.Value(lFld) = sValNew
                    pRowChanged.RowChanged = True
                End If
              Else
                ' Do Nothing - Unknown field type
            End If
        End If
    End If

Process_Exit:
  Exit Sub
ErrorHandler:
  HandleError True, _
              "CompareAndSaveValue " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
              Err.Number, _
              Err.Source, _
              Err.Description, _
              4
End Sub

'***************************************************************************
'Name:                  FormatOMMapNum
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Return properly formatted part of OMMapNum string
'Called From:   No calls are made to this function
'Methods:       This function uses string portions to return a specific
'               portion of an ORMAP number.
'               This functionality has been replaced by the ORMAPNumber
'               class.
'Inputs:        asVal - The ORMAP Number
'               asPartName  - The part of the ORMAP Number to return
'Parameters:    None
'Outputs:       None
'Returns:       A formatted portion of an ORMAP Map Number string
'Errors:        This routine raises no known errors.
'Assumptions:   The passed in ORMAP Number is valid
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  Single exit point using goto
'***************************************************************************

Public Function FormatOMMapNum( _
  ByRef asVal As String, _
  ByRef asPartName As String) As String
On Error GoTo ErrorHandler

    FormatOMMapNum = asVal
    Select Case LCase(asPartName)
        Case "county"
            If Len(FormatOMMapNum) <> 2 Then
                FormatOMMapNum = AddLeadingZeros(FormatOMMapNum, 2)
            End If
        Case "town"
            If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "00"
        Case "townpart"
            FormatOMMapNum = Replace(FormatOMMapNum, "0.", ".")
            'If Len(FormatOMMapNum) <> 3 Then FormatOMMapNum = "000"
        Case "towndir"
            If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "N"
        Case "range"
            If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "01"
        Case "rangepart"
            FormatOMMapNum = Replace(FormatOMMapNum, "0.", ".")
            'If Len(FormatOMMapNum) <> 3 Then FormatOMMapNum = "000"
            'If Len(sVal) <> 3 Then FormatOMMapNum = "000"
        Case "rangedir"
            If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "W"
        Case "section"
            If Len(FormatOMMapNum) <> 2 Then FormatOMMapNum = "00"
        Case "qtr"
            If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "qtrqtr"
            If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "suffixtype"
            If Len(FormatOMMapNum) <> 1 Then FormatOMMapNum = "0"
        Case "suffixnum"
            If Len(FormatOMMapNum) <> 0 And Len(FormatOMMapNum) > 3 Then
                FormatOMMapNum = "000"
                GoTo Process_Exit
            ElseIf Len(FormatOMMapNum) = 1 Then
                FormatOMMapNum = "00" & FormatOMMapNum
                GoTo Process_Exit
            ElseIf Len(FormatOMMapNum) = 2 Then
                FormatOMMapNum = "0" & FormatOMMapNum
                GoTo Process_Exit
            End If
        Case "anomaly"
            If Len(FormatOMMapNum) > 2 Or Len(FormatOMMapNum) = 0 Then
                FormatOMMapNum = "00"
                GoTo Process_Exit
            ElseIf Len(FormatOMMapNum) = 2 Then
            
            ElseIf Len(FormatOMMapNum) = 1 Then
                FormatOMMapNum = "0" & FormatOMMapNum
                GoTo Process_Exit
            End If
        Case Else
            'Nothing to implement now
    End Select

Process_Exit:
  Exit Function
ErrorHandler:
  HandleError True, _
              "FormatOMMapNum " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
              Err.Number, _
              Err.Source, _
              Err.Description, _
              4
End Function

'***************************************************************************
'Name:                  GetAppRef
'Initial Author:        James Moore
'Subsequent Author:     <<Type your name here>>
'Created:               10/11/2006
'Description:   Used to obtain a reference the the Application, which is
'               used throughout the code.
'               This is a more complex process with VB code because the code
'               does not live in the MXD.
'Called From:   No calls are made to this function
'Methods:       Gets a reference to the application through the application
'               running object table.
'               This method is dangerous in that it can potentially return
'               ArcCatalog.  The tools in this DLL are designed for ArcMap.
'               To derive the ArcMap application, use the
'               global application object instead, g_pApp.
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       An object that supports the IApplication interface
'Errors:        This routine raises no known errors.
'Assumptions:
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    11-1-06     removed dead variables
'***************************************************************************

Public Function GetAppRef() As esriFramework.IApplication
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim rot As esriFramework.AppROT
    Dim app As esriFramework.IApplication
    Dim pobjectFactory As esriFramework.IObjectFactory
    '++ END JWalton 2/7/2007
    
    Set rot = New esriFramework.AppROT
    If rot.Count = 1 Then
        Set app = rot.Item(0) 'ArcCatalog
    Else
        Set app = rot.Item(1) 'ArcMap
    End If
    Set pobjectFactory = app
    
    Set GetAppRef = app


    Exit Function
ErrorHandler:
    HandleError True, _
                "GetAppRef " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  GetCentroid
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Return the centroid for annotation and dimension features
'Called From:   No calls are made to this function
'Description:   Determines if this feature is annotation feature class then
'               gets the centroid.
'Methods:       None
'Inputs:        pFeat - An object that support IFeature
'Parameters:    None
'Outputs:       None
'Returns:       A Point object that represents the centroid of the feature,
'               or Nothing if the object is not an annotation or dimesion
'               feature
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    11-1-06     Appears to be a dead routine
'***************************************************************************

Public Function GetCentroid( _
  ByRef pFeat As esriGeoDatabase.IFeature) As esriGeometry.IPoint
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pArea As esriGeometry.IArea
    '++ END JWalton 2/7/2007
    
    If pFeat.FeatureType = esriGeoDatabase.esriFeatureType.esriFTAnnotation Or _
       pFeat.FeatureType = esriGeoDatabase.esriFeatureType.esriFTDimension Then
        Set pArea = pFeat.Shape
        Set GetCentroid = pArea.Centroid
    End If

  Exit Function
ErrorHandler:
    HandleError True, _
                "GetCentroid " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  GetDomainDefaultValue
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Returns the default value if this is a domain field with a
'               default
'Called From:   No calls are made to this function
'Description:   Extracts the default value of field if the field exists,
'               has a domain, and the domain has a default.
'Methods:       None
'Inputs:        pTable -- The table the field resides in
'               sFldName -- The field to find the default domain value for
'Parameters:    None
'Outputs:       None
'Returns:       A variant value representing the default value for the
'               domain in the field
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  single exit point using goto
'James Moore    11-13-06    removed dead variable
'***************************************************************************

Public Function GetDomainDefaultValue( _
  ByRef pTable As esriGeoDatabase.ITable, _
  ByVal sFldName As String) As Variant
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pField As esriGeoDatabase.IField
    Dim pCVDomain As esriGeoDatabase.ICodedValueDomain
    Dim pDomain As esriGeoDatabase.IDomain
    Dim i As Integer
    Dim lFld As Long
    Dim vDomainVal As Variant
    '++ END JWalton 2/7/2007
    
    ' Validate the existence of the domain field
    lFld = pTable.FindField(sFldName)
    If lFld > -1 Then
        Set pField = pTable.Fields.Field(lFld)
      Else
        GetDomainDefaultValue = ""
        GoTo Process_Exit
    End If
    
    ' Determine the default value of the domain
    Set pDomain = pField.Domain
    If pDomain Is Nothing Then
        GetDomainDefaultValue = ""
        GoTo Process_Exit
      Else
        'Determine type of domain  -If Coded Value, get the description
        If TypeOf pDomain Is esriGeoDatabase.ICodedValueDomain Then
            Set pCVDomain = pDomain
            vDomainVal = pField.DefaultValue
            'Search the domain for the code
            For i = 0 To pCVDomain.CodeCount - 1
                If pCVDomain.Value(i) = vDomainVal Then
                    'return the description
                    GetDomainDefaultValue = pCVDomain.Name(i)
                    GoTo Process_Exit
                End If
            Next i
          Else ' If range domain, return the numeric value
            GetDomainDefaultValue = pField.DefaultValue
            GoTo Process_Exit
        End If
    End If  'If pDomain is nothing/Else

Process_Exit:
    Exit Function
ErrorHandler:
    HandleError True, _
                "GetDomainDefaultValue " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  GetFWorkspace
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Extract a workspace from an object
'Called From:   No calls are made to this function
'Description:   Uses query interface to obtain the dataset from the object's
'               feature class, and returns the workspace from the dataset
'Methods:       None
'Inputs:        pObj - A valid initialized geodatabase object
'Parameters:    None
'Outputs:       None
'Returns:       A IFeatureWorkspace interface on a Workspace object
'Errors:        This routine raises no known errors.
'Assumptions:   pObj is a valid initialized geodatabase object
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************

Public Function GetFWorkspace( _
  ByRef pObj As esriGeoDatabase.IObject) As esriGeoDatabase.IFeatureWorkspace
On Error GoTo ErrorHandler
'jwm this procedure is not called by any other process
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pFWS As esriGeoDatabase.IFeatureWorkspace
    Dim pObjClass As esriGeoDatabase.IObjectClass
    Dim pDataset As esriGeoDatabase.IDataset
    '++ END JWalton 2/7/2007
  
    ' QI for the dataset from the feature class
    Set pObjClass = pObj.Class
    Set pDataset = pObjClass
    
    ' Extract the workspace from the dataset
    Set pFWS = pDataset.Workspace
    Set GetFWorkspace = pFWS

Process_Exit:
    Exit Function
ErrorHandler:
    HandleError True, _
                "GetFWorkspace " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  GetMXDocRef
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Get a reference to the current map document
'Called From:   No calls are made to this function
'Methods:       Gets a reference to the application through the application
'               running object table.
'               This method is dangerous in that it can potentially return
'               ArcCatalog.  The tools in this DLL are designed for ArcMap.
'               To derive the MxDoc of the ArcMap application, use the
'               global application object instead, g_pApp.
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       An ESRI ArcMap Document
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    11-1-06     removed dead variables
'***************************************************************************

Public Function GetMXDocRef() As esriArcMapUI.IMxDocument
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim pMXDoc As esriArcMapUI.IMxDocument
    Dim rot As esriFramework.AppROT
    Dim app As esriFramework.IApplication
    Dim pobjectFactory As esriFramework.IObjectFactory
    '++ END JWalton 2/7/2007
    
    ' Gets a reference to the currently running application, preferably ArcMap
    Set rot = New esriFramework.AppROT
    If rot.Count = 1 Then
        Set app = rot.Item(0) 'ArcCatalog
    Else
        Set app = rot.Item(1) 'ArcMap
    End If
    Set pobjectFactory = app
    Set pMXDoc = app.Document
    
    ' Returns the function's value
    Set GetMXDocRef = pMXDoc


    Exit Function
ErrorHandler:
    HandleError True, _
                "GetMXDocRef " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function

'***************************************************************************
'Name:                  GetSpecialInterests
'Initial Author:        <<Unknown>>
'Subsequent Author:     <<Type your name here>>
'Created:               10/11/2006
'Purpose:       Return a properly formatted special interest number given
'               a feature
'Called From:   No calls are made to this function
'Description:   Accepts a feature, extracts the special interest number,
'               and formats it in the form '00000'
'Methods:       None
'Inputs:        pFeature - A feature
'Parameters:    None
'Outputs:       None
'Returns:       A properly formatted string.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  dead routine
'***************************************************************************

Public Function GetSpecialInterests( _
  ByRef pFeature As esriGeoDatabase.IFeature) As String
On Error GoTo ErrorHandler
    '++ START JWalton 2/7/2007 Centralized Variable Declarations
    Dim lTLSpecInterestFld As Long
    Dim sTLSpecVal As String
    '++ END JWalton 2/7/2007

    lTLSpecInterestFld = LocateFields(pFeature.Class, g_pFldnames.TLSpecInterestFN)
    If lTLSpecInterestFld = -1 Then
        sTLSpecVal = "00000"
      Else
        If Not IsNull(pFeature.Value(lTLSpecInterestFld)) Then
            sTLSpecVal = pFeature.Value(lTLSpecInterestFld)
          Else
            sTLSpecVal = "00000"
        End If
        'Verify that it is 5 digits
        If Len(sTLSpecVal) < 5 Then
            Do Until Len(sTLSpecVal) = 5
                sTLSpecVal = "0" & sTLSpecVal
            Loop
        End If
    End If
    GetSpecialInterests = sTLSpecVal

    Exit Function
ErrorHandler:
    HandleError True, _
                "GetSpecialInterests " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function


'***************************************************************************
'Name:                  Validate5Digits
'Initial Author:        <<Unknown>>
'Subsequent Author:     Type your name here.
'Created:               10/11/2006
'Purpose:       String formatting
'Called From:   No calls are made to this function
'Description:   Accepts a string and pads it to the left with zeros up to
'               five characters
'Methods:       None
'Inputs:        sString - The string to pad
'Parameters:    None
'Outputs:       None
'Returns:       A properly formatted string
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'James Moore    10/11/2006  THIS FUNCTION IS DEAD. IT IS NOT CALLED BY ANY _
                            OTHER PROCESS.
'***************************************************************************

Public Function Validate5Digits( _
  ByVal sString As String) As String
On Error GoTo ErrorHandler

    'Make sure taxlot number is 5 characters
    If Len(sString) < 5 Then
        Do Until Len(sString) = 5
            sString = "0" & sString
        Loop
    End If
    Validate5Digits = sString

    Exit Function
ErrorHandler:
    HandleError True, _
                "Validate5Digits " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
                Err.Number, _
                Err.Source, _
                Err.Description, _
                4
End Function
