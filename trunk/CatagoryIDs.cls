VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryIDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
' SCC Revision number: $Revision: 88 $
' Date of last change: $Date: 2007-02-21 09:27:42 -0800 (Wed, 21 Feb 2007) $
'
'
' File name:            CategoryIDs
'
' Initial Author:       JWalton
'
' Date Created:         4/11/2007
'
' Description: Definitions of system GUIDs for ArcGIS component categories.
'
' Entry points:
'       Multiple
'
' Dependencies:
'       File References
'           esriSystem
'
' Issues:
'       <<None>>
'
' Method:
'       Allows use of the CatInstall.exe utility located in the ArcGIS\Bin folder
'       to register the DLL with the operating, and components with ArcGIS.
'
' Updates:
'       4/11/2007 -- Initial implementation. (JWalton)


Option Explicit
'******************************
' Private Definitions
'------------------------------
'------------------------------
' Private Variables
'------------------------------
Private m_colComponentGuids As Collection
Private m_colComponentNames As Collection

'***************************************************************************
'Name:                  Class_Initialize
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Initialize class variables and objects
'Called From:   Nowhere
'Description:   Class initialization routine.
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
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Private Sub Class_Initialize()
    ' Initialize class catalog collections
    Set m_colComponentGuids = New Collection
    Set m_colComponentNames = New Collection
    
    ' Create a catalog of component guids
    CatalogComponentGuids
    
    ' Create a catalog of component names
    CatalogComponentNames
End Sub

'***************************************************************************
'Name:                  Class_Initialize
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Initialize class variables and objects
'Called From:   Nowhere
'Description:   Class initialization routine.
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
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Private Sub Class_Terminate()
    ' Destroy the class catalog collections
    Set m_colComponentGuids = Nothing
    Set m_colComponentNames = Nothing
End Sub

'***************************************************************************
'Name:                  CategoryNames
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Return all component categories as a collection of strings
'Called From:   Nowhere
'Description:   None
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       Collection
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Public Function CategoryNames() As Collection
    Set CategoryNames = m_colComponentNames
End Function

'***************************************************************************
'Name:                  CategoryGUIDs
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Return all component categories GUIDs as a collection of
'               strings
'Called From:   Nowhere
'Description:   None
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       Collection
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Public Function CategoryGUIDs() As Collection
    Set CategoryGUIDs = m_colComponentGuids
End Function

'***************************************************************************
'Name:                  CategoryByName
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Return a program identifier of a component category given
'               its name
'Called From:   Nowhere
'Description:   None
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       esriSystem.UID as esriSystem.IUID
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Public Function CategoryByName(ByVal strCategory As String) As esriSystem.IUID
    Set CategoryByName = New esriSystem.UID
    CategoryByName.Value = strCategory
End Function

'***************************************************************************
'Name:                  MakeUID
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Return a program identifier of a component category given
'               its program ID.
'Called From:   Nowhere
'Description:   None
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       esriSystem.UID as esriSystem.IUID
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Public Function MakeUID(ByVal ProgID_or_IID As String) As esriSystem.IUID
  ' Return a UID with its Value property set to the ProgID_or_IID
  Set MakeUID = New esriSystem.UID
  MakeUID.Value = ProgID_or_IID
End Function

'***************************************************************************
'Name:                  <<Properties>>
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Exposure of program identifiers to external callers
'Called From:   Nowhere
'Description:   None
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       esriSystem.UID as esriSystem.IUID
'Errors:        Routines raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Public Property Get DDDExtensions() As esriSystem.IUID
  Set DDDExtensions = MakeUID("{A1523ADD-FB91-11D4-80FC-00C04FA0ADF8}")
End Property

Public Property Get AdjustmentMethodPropertyPages() As esriSystem.IUID
  Set AdjustmentMethodPropertyPages = MakeUID("{0772EAEC-F5B4-4F2C-BFE8-019861BD3A43}")
End Property

Public Property Get AdjustmentPropertyPages() As esriSystem.IUID
  Set AdjustmentPropertyPages = MakeUID("{3E15409F-572B-47DE-91C3-7F310CE04F5B}")
End Property

Public Property Get AnimationPropertyPages() As esriSystem.IUID
  Set AnimationPropertyPages = MakeUID("{E8778391-8124-11D5-B29B-00508BCDDE28}")
End Property

Public Property Get AnnotateMapPropertyPages() As esriSystem.IUID
  Set AnnotateMapPropertyPages = MakeUID("{1532016F-7C3E-46EA-8753-F3A0E05EA015}")
End Property

Public Property Get AnnotationEngines() As esriSystem.IUID
  Set AnnotationEngines = MakeUID("{6ED8A829-5036-11D3-9FB3-00C04F6BC6A5}")
End Property

Public Property Get AnnotationExpresionPages() As esriSystem.IUID
  Set AnnotationExpresionPages = MakeUID("{F3E803A1-6EC5-11D2-A2C6-080009B6F22B}")
End Property

Public Property Get AnnotationExpressionParsers() As esriSystem.IUID
  Set AnnotationExpressionParsers = MakeUID("{02D588E1-3D42-11D2-A2A2-080009B6F22B}")
End Property

Public Property Get AnnotationPlacementPages() As esriSystem.IUID
  Set AnnotationPlacementPages = MakeUID("{2B5310E1-6EC9-11D2-A2C6-080009B6F22B}")
End Property

Public Property Get AnnotationPropertiesConverters() As esriSystem.IUID
  Set AnnotationPropertiesConverters = MakeUID("{6ED8B756-1BAA-12D3-9FB3-00C04F6BB7A6}")
End Property

Public Property Get ArcToolboxTools() As esriSystem.IUID
  Set ArcToolboxTools = MakeUID("{B5E9A20C-0D77-11D4-9FE0-00C04F6BC61D}")
End Property

Public Property Get AreaPatches() As esriSystem.IUID
  Set AreaPatches = MakeUID("{F21E3E7B-E3BA-11D2-B868-00600802E603}")
End Property

Public Property Get AttributeTransferTypes() As esriSystem.IUID
  Set AttributeTransferTypes = MakeUID("{51C6C6AB-DACE-49DB-AF63-32F241C35E49}")
End Property

Public Property Get AVFSrcImporters() As esriSystem.IUID
  Set AVFSrcImporters = MakeUID("{56BBC8A4-E583-11D3-9FD8-00C04F6BC78E}")
End Property

Public Property Get AVThemeImporters() As esriSystem.IUID
  Set AVThemeImporters = MakeUID("{56BBC8A1-E583-11D3-9FD8-00C04F6BC78E}")
End Property

Public Property Get Backgrounds() As esriSystem.IUID
  Set Backgrounds = MakeUID("{EB5A6D3A-DE07-11D2-B868-00600802E603}")
End Property

Public Property Get Borders() As esriSystem.IUID
  Set Borders = MakeUID("{EB5A6D39-DE07-11D2-B868-00600802E603}")
End Property

Public Property Get CadDrawingPropertyPages() As esriSystem.IUID
  Set CadDrawingPropertyPages = MakeUID("{83EEE9C7-0CB1-11D3-9B35-00C04FA33299}")
End Property

Public Property Get CadFeatureClassPropertyPages() As esriSystem.IUID
  Set CadFeatureClassPropertyPages = MakeUID("{83EEE9C6-0CB1-11D3-9B35-00C04FA33299}")
End Property

Public Property Get CadFeaturePropertyPages() As esriSystem.IUID
  Set CadFeaturePropertyPages = MakeUID("{83EEE9C5-0CB1-11D3-9B35-00C04FA33299}")
End Property

Public Property Get ClassificationMethods() As esriSystem.IUID
  Set ClassificationMethods = MakeUID("{F5B68AD5-EB4E-11D1-AAB9-00C04FA334B3}")
End Property

Public Property Get ClipboardFormats() As esriSystem.IUID
  Set ClipboardFormats = MakeUID("{52BB5364-947E-11D2-ACFF-0000F87808EE}")
End Property

Public Property Get ColorCorrectionPages() As esriSystem.IUID
  Set ColorCorrectionPages = MakeUID("{3745A6D7-50EF-11D3-9FF4-00C04F6BC8DD}")
End Property

Public Property Get ColorPropertyPages() As esriSystem.IUID
  Set ColorPropertyPages = MakeUID("{31CB4BD9-2FDC-11D3-9F58-00C04F6BC709}")
End Property

Public Property Get ColorRamp() As esriSystem.IUID
  Set ColorRamp = MakeUID("{2411EC71-8BCC-11D2-A2EE-080009B6F22B}")
End Property

Public Property Get ColorRampPropetyPages() As esriSystem.IUID
  Set ColorRampPropetyPages = MakeUID("{2411EC72-8BCC-11D2-A2EE-080009B6F22B}")
End Property

Public Property Get ContentsViews() As esriSystem.IUID
  Set ContentsViews = MakeUID("{089874FC-CC18-11D2-9F39-00C04F6BC78E}")
End Property

Public Property Get ContextAnalyzers() As esriSystem.IUID
  Set ContextAnalyzers = MakeUID("{0C452011-84CB-11D2-AE68-080009EC732A}")
End Property

Public Property Get CoverageFeatureClassPropertyPages() As esriSystem.IUID
  Set CoverageFeatureClassPropertyPages = MakeUID("{5F7BBE84-E6EC-11D2-9F30-00C04F79927C}")
End Property

Public Property Get CoveragePropertyPages() As esriSystem.IUID
  Set CoveragePropertyPages = MakeUID("{5F7BBE83-E6EC-11D2-9F30-00C04F79927C}")
End Property

Public Property Get DataExclusionPropertyPages() As esriSystem.IUID
  Set DataExclusionPropertyPages = MakeUID("{75C2A63E-A98A-11D2-AB27-00C04FA334B3}")
End Property

Public Property Get DataGraphPropertyPages() As esriSystem.IUID
  Set DataGraphPropertyPages = MakeUID("{977D75CF-E00C-11D3-A664-0008C7DF88DB}")
End Property

Public Property Get DataSamplingPropertyPages() As esriSystem.IUID
  Set DataSamplingPropertyPages = MakeUID("{5C15FD91-2DB2-11D3-9F52-00C04F6BC709}")
End Property

Public Property Get DataWindowFactory() As esriSystem.IUID
  Set DataWindowFactory = MakeUID("{B56A7C49-83D4-11D2-A2E9-080009B6F22B}")
End Property

Public Property Get DDDAnimationTypes() As esriSystem.IUID
  Set DDDAnimationTypes = MakeUID("{5980E69C-A95E-11D5-B2A0-00508BCDDE28}")
End Property

Public Property Get DimensionConstructors() As esriSystem.IUID
  Set DimensionConstructors = MakeUID("{0864F69B-E975-11D3-9F79-00C04F6BDDD9}")
End Property

Public Property Get EditorExtensions() As esriSystem.IUID
  Set EditorExtensions = MakeUID("{58122E51-D89C-11D2-9F25-00C04F6BC979}")
End Property

Public Property Get EditorPropertyPages() As esriSystem.IUID
  Set EditorPropertyPages = MakeUID("{58122E52-D89C-11D2-9F25-00C04F6BC979}")
End Property

Public Property Get EditTasks() As esriSystem.IUID
  Set EditTasks = MakeUID("{58122E50-D89C-11D2-9F25-00C04F6BC979}")
End Property

Public Property Get ElementPropertyPages() As esriSystem.IUID
  Set ElementPropertyPages = MakeUID("{43B21395-F1E2-11D2-B870-00600802E603}")
End Property

Public Property Get ESDAEngineMethods() As esriSystem.IUID
  Set ESDAEngineMethods = MakeUID("{DA805B10-54D7-11D3-8D39-00C04F5B87B2}")
End Property

Public Property Get ExporterDrivers() As esriSystem.IUID
  Set ExporterDrivers = MakeUID("{5FDE0991-7B12-11D2-ACF9-0000F87808EE}")
End Property

Public Property Get ExporterPropertyPages() As esriSystem.IUID
  Set ExporterPropertyPages = MakeUID("{5FDE099A-7B12-11D2-ACF9-0000F87808EE}")
End Property

Public Property Get FeatureClassPropertyPages() As esriSystem.IUID
  Set FeatureClassPropertyPages = MakeUID("{750D818E-E0B5-11D2-99BF-0000F80372B4}")
End Property

Public Property Get FeatureDatasetPropertyPages() As esriSystem.IUID
  Set FeatureDatasetPropertyPages = MakeUID("{6EDC31E1-E3AD-11D2-99C1-0000F80372B4}")
End Property

Public Property Get FeatureLayerSymbology() As esriSystem.IUID
  Set FeatureLayerSymbology = MakeUID("{7480BE41-EEA1-4E16-858C-B2AB6C92F720}")
End Property

Public Property Get FillSymbol() As esriSystem.IUID
  Set FillSymbol = MakeUID("{5424FB02-5713-11D2-97E5-0080C7E04196}")
End Property

Public Property Get FontMappingPropertyPages() As esriSystem.IUID
  Set FontMappingPropertyPages = MakeUID("{FB5216EF-6AC6-11D3-9F4E-00C04F6BC8DD}")
End Property

Public Property Get FrameElementPropertyPages() As esriSystem.IUID
  Set FrameElementPropertyPages = MakeUID("{43B21396-F1E2-11D2-B870-00600802E603}")
End Property

Public Property Get GeoAnalysis() As esriSystem.IUID
  Set GeoAnalysis = MakeUID("{3297E9C1-93A1-11D2-9F33-00C04F8ED1D7}")
End Property

Public Property Get GeodatabaseWorkspaceExtensions() As esriSystem.IUID
  Set GeodatabaseWorkspaceExtensions = MakeUID("{DA10C1C7-09A3-11D4-9FAA-00C04F6BDF0C}")
End Property

Public Property Get GeoElementTools() As esriSystem.IUID
  Set GeoElementTools = MakeUID("{B56A7C44-83D4-11D2-A2E9-080009B6F22B}")
End Property

Public Property Get GeographicCoordSysPropertyPages() As esriSystem.IUID
  Set GeographicCoordSysPropertyPages = MakeUID("{7E9913E1-E22E-11D2-99C0-0000F80372B4}")
End Property

Public Property Get GeometricNetworkPropertyPages() As esriSystem.IUID
  Set GeometricNetworkPropertyPages = MakeUID("{750D818F-E0B5-11D2-99BF-0000F80372B4}")
End Property

Public Property Get GeoObjectClassDescriptions() As esriSystem.IUID
  Set GeoObjectClassDescriptions = MakeUID("{0813623A-A72E-11D2-8924-0000F877762D}")
End Property

Public Property Get GeoObjectClassExtensions() As esriSystem.IUID
  Set GeoObjectClassExtensions = MakeUID("{D4E2A322-5D59-11D2-89FD-006097AFF44E}")
End Property

Public Property Get GeoObjects() As esriSystem.IUID
  Set GeoObjects = MakeUID("{D4E2A321-5D59-11D2-89FD-006097AFF44E}")
End Property

Public Property Get GeoRefSaveAsDlg() As esriSystem.IUID
  Set GeoRefSaveAsDlg = MakeUID("{9FB18846-5930-4D7A-AFBE-0FF2E4B6DA8C}")
End Property

Public Property Get GeostatExtensions() As esriSystem.IUID
  Set GeostatExtensions = MakeUID("{ACCCD296-9D8E-4F6B-97F4-39D65330D891}")
End Property

Public Property Get GNEditorPropertyPages() As esriSystem.IUID
  Set GNEditorPropertyPages = MakeUID("{583A720C-FC86-46A1-A6DD-E5C315076BEB}")
End Property

Public Property Get GraphicSnapAgents() As esriSystem.IUID
  Set GraphicSnapAgents = MakeUID("{D618307B-AA29-11D2-87F2-0000F8751720}")
End Property

Public Property Get GridLabels() As esriSystem.IUID
  Set GridLabels = MakeUID("{8AF09A6C-A0DD-11D2-AE7F-080009EC732A}")
End Property

Public Property Get GSEngineMethods() As esriSystem.IUID
  Set GSEngineMethods = MakeUID("{23B9BCD6-6F6F-11D2-AAC3-00C04FA375FB}")
End Property

Public Property Get GSRenderers() As esriSystem.IUID
  Set GSRenderers = MakeUID("{60B20514-C449-11D3-A63F-0008C7BF3347}")
End Property

Public Property Get GxArcViewExportFeatureClassFilters() As esriSystem.IUID
  Set GxArcViewExportFeatureClassFilters = MakeUID("{9EB1D578-740B-423E-8CF4-C77A1775A5D6}")
End Property

Public Property Get GxArcViewExportTableFilters() As esriSystem.IUID
  Set GxArcViewExportTableFilters = MakeUID("{593C6231-9B13-4229-A1DF-A7A36F1AE10D}")
End Property

Public Property Get GxCadDataContextMenuCommands() As esriSystem.IUID
  Set GxCadDataContextMenuCommands = MakeUID("{83EEE9CA-0CB1-11D3-9B35-00C04FA33299}")
End Property

Public Property Get GxCadDrawingContextMenuCommands() As esriSystem.IUID
  Set GxCadDrawingContextMenuCommands = MakeUID("{83EEE9C1-0CB1-11D3-9B35-00C04FA33299}")
End Property

Public Property Get GxCadDrawingExportMenuCommands() As esriSystem.IUID
  Set GxCadDrawingExportMenuCommands = MakeUID("{83EEE9C4-0CB1-11D3-9B35-00C04FA33299}")
End Property

Public Property Get GxCadFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxCadFeatureClassContextMenuCommands = MakeUID("{83EEE9C0-0CB1-11D3-9B35-00C04FA33299}")
End Property

Public Property Get GxCadFeatureClassExportMenuCommands() As esriSystem.IUID
  Set GxCadFeatureClassExportMenuCommands = MakeUID("{83EEE9C3-0CB1-11D3-9B35-00C04FA33299}")
End Property

Public Property Get GxCadFeatureContextMenuCommands() As esriSystem.IUID
  Set GxCadFeatureContextMenuCommands = MakeUID("{83EEE9BF-0CB1-11D3-9B35-00C04FA33299}")
End Property

Public Property Get GxCadFeatureExportMenuCommands() As esriSystem.IUID
  Set GxCadFeatureExportMenuCommands = MakeUID("{83EEE9C2-0CB1-11D3-9B35-00C04FA33299}")
End Property

Public Property Get GxCatalogContextMenuCommands() As esriSystem.IUID
  Set GxCatalogContextMenuCommands = MakeUID("{25C0E6F0-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxCatalogPropertyPages() As esriSystem.IUID
  Set GxCatalogPropertyPages = MakeUID("{BD1DF56D-E644-11D2-99C1-0000F80372B4}")
End Property

Public Property Get GxCommandBars() As esriSystem.IUID
  Set GxCommandBars = MakeUID("{56C205F9-E53A-11D1-9496-080009EEBECB}")
End Property

Public Property Get GxCommands() As esriSystem.IUID
  Set GxCommands = MakeUID("{5F08CBCA-E91F-11D1-AEE8-080009EC734B}")
End Property

Public Property Get GxCoverageContextMenuCommands() As esriSystem.IUID
  Set GxCoverageContextMenuCommands = MakeUID("{25B9A22F-E6C7-11D2-9F30-00C04F79927C}")
End Property

Public Property Get GxCoverageExportMenuCommands() As esriSystem.IUID
  Set GxCoverageExportMenuCommands = MakeUID("{A6BF6E97-EB70-11D2-9F31-00C04F79927C}")
End Property

Public Property Get GxCoverageFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxCoverageFeatureClassContextMenuCommands = MakeUID("{25B9A230-E6C7-11D2-9F30-00C04F79927C}")
End Property

Public Property Get GxCoverageFeatureClassExportMenuCommands() As esriSystem.IUID
  Set GxCoverageFeatureClassExportMenuCommands = MakeUID("{A6BF6E98-EB70-11D2-9F31-00C04F79927C}")
End Property

Public Property Get GxCoverageNewMenuCommands() As esriSystem.IUID
  Set GxCoverageNewMenuCommands = MakeUID("{1449254B-319B-11D3-9F33-00C04F79927C}")
End Property

Public Property Get GxDatabaseContextMenuCommands() As esriSystem.IUID
  Set GxDatabaseContextMenuCommands = MakeUID("{25C0E6F5-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxDatabaseExtensions() As esriSystem.IUID
  Set GxDatabaseExtensions = MakeUID("{6CB2F1D8-4E9D-11D3-9F52-00C04F6BDF06}")
End Property

Public Property Get GxDatabaseImportMenuCommands() As esriSystem.IUID
  Set GxDatabaseImportMenuCommands = MakeUID("{7513EB97-D7EE-11D2-9F50-00C04F6BC626}")
End Property

Public Property Get GxDatabaseNewMenuCommands() As esriSystem.IUID
  Set GxDatabaseNewMenuCommands = MakeUID("{25C0E706-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxDataGraphContextMenuCommands() As esriSystem.IUID
  Set GxDataGraphContextMenuCommands = MakeUID("{B14315A5-098A-11D4-A676-0008C7DF88DB}")
End Property

Public Property Get GxDatasetContextMenuCommands() As esriSystem.IUID
  Set GxDatasetContextMenuCommands = MakeUID("{25C0E6F7-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxDialogObjectFilters() As esriSystem.IUID
  Set GxDialogObjectFilters = MakeUID("{BB8F3046-1D26-11D3-9F55-00C04F6BC69E}")
End Property

Public Property Get GxDiskConnectionContextMenuCommands() As esriSystem.IUID
  Set GxDiskConnectionContextMenuCommands = MakeUID("{25C0E6F1-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxDockableWindows() As esriSystem.IUID
  Set GxDockableWindows = MakeUID("{117623B4-F9D1-11D3-A67F-0008C7DF97B9}")
End Property

Public Property Get GxEnabledWorkspaceFactories() As esriSystem.IUID
  Set GxEnabledWorkspaceFactories = MakeUID("{6E208C99-DBD3-11D2-9F2F-00C04F6BC69E}")
End Property

Public Property Get GxEnterpriseFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxEnterpriseFeatureClassContextMenuCommands = MakeUID("{F4EDFE54-E856-4D72-AEC6-7BD59A3E7B52}")
End Property

Public Property Get GxEnterpriseFeatureClassContextMenuCommandsAV() As esriSystem.IUID
  Set GxEnterpriseFeatureClassContextMenuCommandsAV = MakeUID("{9648D913-4475-4165-B746-D2058776F9E8}")
End Property

Public Property Get GxEnterpriseFeatureDatasetContextMenuCommands() As esriSystem.IUID
  Set GxEnterpriseFeatureDatasetContextMenuCommands = MakeUID("{DF03366D-F30C-490B-A353-B05BAA18A3AF}")
End Property

Public Property Get GxEnterpriseRasterBandContextMenuCommands() As esriSystem.IUID
  Set GxEnterpriseRasterBandContextMenuCommands = MakeUID("{16685968-2C9E-45A5-9D47-F2DC62942DDE}")
End Property

Public Property Get GxEnterpriseRasterDatasetContextMenuCommands() As esriSystem.IUID
  Set GxEnterpriseRasterDatasetContextMenuCommands = MakeUID("{F2889296-A1EC-4BDF-AEA9-D08D8E0837D9}")
End Property

Public Property Get GxEnterpriseRelationshipClassContextMenuCommands() As esriSystem.IUID
  Set GxEnterpriseRelationshipClassContextMenuCommands = MakeUID("{2004CEC9-E2BD-4D2D-8F3F-9E791E7AB106}")
End Property

Public Property Get GxEnterpriseStandaloneFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxEnterpriseStandaloneFeatureClassContextMenuCommands = MakeUID("{DABA683A-7888-40BE-8E70-49BC0D4F8C30}")
End Property

Public Property Get GxEnterpriseStandaloneFeatureClassContextMenuCommandsAV() As esriSystem.IUID
  Set GxEnterpriseStandaloneFeatureClassContextMenuCommandsAV = MakeUID("{39C1320B-01A5-4D61-B286-2880337B53C4}")
End Property

Public Property Get GxEnterpriseTableContextMenuCommands() As esriSystem.IUID
  Set GxEnterpriseTableContextMenuCommands = MakeUID("{9B9ACF28-60FE-4ABA-9DC4-5B060A903DF2}")
End Property

Public Property Get GxEnterpriseTableContextMenuCommandsAV() As esriSystem.IUID
  Set GxEnterpriseTableContextMenuCommandsAV = MakeUID("{0273158A-0EA1-4627-B60A-9D3DF568F8F6}")
End Property

Public Property Get GxExportFeatureClassFilters() As esriSystem.IUID
  Set GxExportFeatureClassFilters = MakeUID("{8E415006-F854-4ECB-98C3-EFBD68BAF635}")
End Property

Public Property Get GxExportTableFilters() As esriSystem.IUID
  Set GxExportTableFilters = MakeUID("{260A3656-96D0-417B-AC05-3469A444015A}")
End Property

Public Property Get GxExtensions() As esriSystem.IUID
  Set GxExtensions = MakeUID("{4531C69D-DC07-11D2-9F2F-00C04F6BC69E}")
End Property

Public Property Get GxExtensionToolbar() As esriSystem.IUID
  Set GxExtensionToolbar = MakeUID("{58374DA9-3A52-11D3-9F70-00C04F6BC69E}")
End Property

Public Property Get GxFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxFeatureClassContextMenuCommands = MakeUID("{25C0E6F9-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxFeatureClassExportMenuCommands() As esriSystem.IUID
  Set GxFeatureClassExportMenuCommands = MakeUID("{7513EB9B-D7EE-11D2-9F50-00C04F6BC626}")
End Property

Public Property Get GxFeatureDatasetContextMenuCommands() As esriSystem.IUID
  Set GxFeatureDatasetContextMenuCommands = MakeUID("{25C0E6F8-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxFeatureDatasetExportMenuCommands() As esriSystem.IUID
  Set GxFeatureDatasetExportMenuCommands = MakeUID("{7513EB9A-D7EE-11D2-9F50-00C04F6BC626}")
End Property

Public Property Get GxFeatureDatasetImportMenuCommands() As esriSystem.IUID
  Set GxFeatureDatasetImportMenuCommands = MakeUID("{7513EB99-D7EE-11D2-9F50-00C04F6BC626}")
End Property

Public Property Get GxFeatureDatasetNewMenuCommands() As esriSystem.IUID
  Set GxFeatureDatasetNewMenuCommands = MakeUID("{25C0E707-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxFileContextMenuCommands() As esriSystem.IUID
  Set GxFileContextMenuCommands = MakeUID("{25C0E700-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxFilterFeatureClasses() As esriSystem.IUID
  Set GxFilterFeatureClasses = MakeUID("{FBFDAC87-FE57-4ABA-B5AA-571C78B6399D}")
End Property

Public Property Get GxFilterFeatureDatasets() As esriSystem.IUID
  Set GxFilterFeatureDatasets = MakeUID("{3EB92610-DB62-465E-A326-8E0E71C24C20}")
End Property

Public Property Get GxFilterTables() As esriSystem.IUID
  Set GxFilterTables = MakeUID("{5FC6B144-22A8-487F-A208-67C34877F120}")
End Property

Public Property Get GxFolderContextMenuCommands() As esriSystem.IUID
  Set GxFolderContextMenuCommands = MakeUID("{25C0E6F4-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxFolderNewMenuCommands() As esriSystem.IUID
  Set GxFolderNewMenuCommands = MakeUID("{25C0E705-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxFolderNewMenuCommandsAV() As esriSystem.IUID
  Set GxFolderNewMenuCommandsAV = MakeUID("{1A9A50A4-F123-4CEC-93DC-8AF1CE0CDDA0}")
End Property

Public Property Get GxGeometricNetworkContextMenuCommands() As esriSystem.IUID
  Set GxGeometricNetworkContextMenuCommands = MakeUID("{25C0E6FB-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxGridContextMenuCommands() As esriSystem.IUID
  Set GxGridContextMenuCommands = MakeUID("{25C0E6FE-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxInfoTableContextMenuCommands() As esriSystem.IUID
  Set GxInfoTableContextMenuCommands = MakeUID("{25B9A231-E6C7-11D2-9F30-00C04F79927C}")
End Property

Public Property Get GxInfoTableExportMenuCommands() As esriSystem.IUID
  Set GxInfoTableExportMenuCommands = MakeUID("{A6BF6E99-EB70-11D2-9F31-00C04F79927C}")
End Property

Public Property Get GxInfoTableToFeatureClassMenuCommands() As esriSystem.IUID
  Set GxInfoTableToFeatureClassMenuCommands = MakeUID("{506187AA-301B-11D4-9FC8-00C04F6BDF06}")
End Property

Public Property Get GxLayerContextMenuCommands() As esriSystem.IUID
  Set GxLayerContextMenuCommands = MakeUID("{25C0E702-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxLayerFilters() As esriSystem.IUID
  Set GxLayerFilters = MakeUID("{C2C372C2-1EB4-4101-9A38-992536A5C0D0}")
End Property

Public Property Get GxLocatorContextMenuCommands() As esriSystem.IUID
  Set GxLocatorContextMenuCommands = MakeUID("{940C2D60-4AAD-11D3-9F51-00C04F6BDF06}")
End Property

Public Property Get GxLocatorFolderContextMenuCommands() As esriSystem.IUID
  Set GxLocatorFolderContextMenuCommands = MakeUID("{E62D6CB4-4AB0-11D3-9F51-00C04F6BDF06}")
End Property

Public Property Get GxLocatorMenus() As esriSystem.IUID
  Set GxLocatorMenus = MakeUID("{55EFFA76-622C-11D3-9F60-00C04F8ED1C4}")
End Property

Public Property Get GxMapContextMenuCommands() As esriSystem.IUID
  Set GxMapContextMenuCommands = MakeUID("{25C0E703-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxMapPropertyPages() As esriSystem.IUID
  Set GxMapPropertyPages = MakeUID("{2044F8C9-D8D7-11D2-99BE-0000F80372B4}")
End Property

Public Property Get GxMetadataExporter() As esriSystem.IUID
  Set GxMetadataExporter = MakeUID("{E5FB4CF2-660C-11D3-A68B-0008C7D3AE50}")
End Property

Public Property Get GxMetadataImporter() As esriSystem.IUID
  Set GxMetadataImporter = MakeUID("{E5FB4CF1-660C-11D3-A68B-0008C7D3AE50}")
End Property

Public Property Get GxMultipleSelectionContextMenuCommands() As esriSystem.IUID
  Set GxMultipleSelectionContextMenuCommands = MakeUID("{35DB0ACD-1DAF-11D3-9FAE-00C04F6BC626}")
End Property

Public Property Get GxObjectFactory() As esriSystem.IUID
  Set GxObjectFactory = MakeUID("{65D77504-895C-11D2-AF5D-080009EC734B}")
End Property

Public Property Get GxObjectFilters() As esriSystem.IUID
  Set GxObjectFilters = MakeUID("{BB8F3045-1D26-11D3-9F55-00C04F6BC69E}")
End Property

Public Property Get GxOleDBDatabaseContextMenuCommands() As esriSystem.IUID
  Set GxOleDBDatabaseContextMenuCommands = MakeUID("{8F4CFA6B-B38E-4E4B-9AB1-E1AFB3CBB569}")
End Property

Public Property Get GxOleDBTableContextMenuCommands() As esriSystem.IUID
  Set GxOleDBTableContextMenuCommands = MakeUID("{44C34C54-E41B-40EF-AC63-8D25A9FC01FA}")
End Property

Public Property Get GxPalettes() As esriSystem.IUID
  Set GxPalettes = MakeUID("{5F08CBCB-E91F-11D1-AEE8-080009EC734B}")
End Property

Public Property Get GxPCCoverageContextMenuCommands() As esriSystem.IUID
  Set GxPCCoverageContextMenuCommands = MakeUID("{4B6C687E-66D2-11D3-9F36-00C04F79927C}")
End Property

Public Property Get GxPCCoverageFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxPCCoverageFeatureClassContextMenuCommands = MakeUID("{4B6C687F-66D2-11D3-9F36-00C04F79927C}")
End Property

Public Property Get GxPMFMapContextMenuCommands() As esriSystem.IUID
  Set GxPMFMapContextMenuCommands = MakeUID("{F7F3ABF6-FB3C-4E7F-9F38-D45DCDF5F30C}")
End Property

Public Property Get GxPMFMapPropertyPages() As esriSystem.IUID
  Set GxPMFMapPropertyPages = MakeUID("{C00B1F2C-52D1-404F-8DD5-E3694FD60E00}")
End Property

Public Property Get GxPre70CoverageContextMenuCommands() As esriSystem.IUID
  Set GxPre70CoverageContextMenuCommands = MakeUID("{C65A2BB4-32ED-11D3-9F33-00C04F79927C}")
End Property

Public Property Get GxPreviews() As esriSystem.IUID
  Set GxPreviews = MakeUID("{3EEA2CB0-A730-11D2-AF6D-080009EC734B}")
End Property

Public Property Get GxPrjFileContextMenuCommands() As esriSystem.IUID
  Set GxPrjFileContextMenuCommands = MakeUID("{25C0E701-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxRasterBandContextMenuCommands() As esriSystem.IUID
  Set GxRasterBandContextMenuCommands = MakeUID("{25C0E6FD-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxRasterDatasetContextMenuCommands() As esriSystem.IUID
  Set GxRasterDatasetContextMenuCommands = MakeUID("{25C0E6FC-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxReadOnlyFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxReadOnlyFeatureClassContextMenuCommands = MakeUID("{A44C92BC-79D6-4F89-A0A8-40CDCC279EDC}")
End Property

Public Property Get GxReadOnlyFeatureDatasetContextMenuCommands() As esriSystem.IUID
  Set GxReadOnlyFeatureDatasetContextMenuCommands = MakeUID("{5484E3DE-86FB-489C-8EB1-73EC4A0E4485}")
End Property

Public Property Get GxReadOnlyStandaloneFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxReadOnlyStandaloneFeatureClassContextMenuCommands = MakeUID("{D1E6735C-4464-4875-8673-59D94A603C3C}")
End Property

Public Property Get GxReadOnlyTableContextMenuCommands() As esriSystem.IUID
  Set GxReadOnlyTableContextMenuCommands = MakeUID("{3BF096DF-B346-46AF-99E2-AFB52ED90F4E}")
End Property

Public Property Get GxRelationshipClassContextMenuCommands() As esriSystem.IUID
  Set GxRelationshipClassContextMenuCommands = MakeUID("{6C9CB46D-DCA2-11D2-9F54-00C04F6BC626}")
End Property

Public Property Get GxRemoteDatabaseContextMenuCommands() As esriSystem.IUID
  Set GxRemoteDatabaseContextMenuCommands = MakeUID("{25C0E6F6-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxRemoteDatabaseFolderContextMenuCommands() As esriSystem.IUID
  Set GxRemoteDatabaseFolderContextMenuCommands = MakeUID("{25C0E6F2-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxRootObjects() As esriSystem.IUID
  Set GxRootObjects = MakeUID("{6E208C9A-DBD3-11D2-9F2F-00C04F6BC69E}")
End Property

Public Property Get GxSceneContextMenuCommands() As esriSystem.IUID
  Set GxSceneContextMenuCommands = MakeUID("{3E804E72-8FE3-11D4-AB88-0008C73FD50C}")
End Property

Public Property Get GxScenePropertyPages() As esriSystem.IUID
  Set GxScenePropertyPages = MakeUID("{A72131C2-90FA-11D4-AB88-0008C73FD50C}")
End Property

Public Property Get GxSDCFeatureClassExportMenuCommands() As esriSystem.IUID
  Set GxSDCFeatureClassExportMenuCommands = MakeUID("{95BCD5A1-8C27-4ACE-9675-BFF5391CE339}")
End Property

Public Property Get GxSDCTableExportMenuCommands() As esriSystem.IUID
  Set GxSDCTableExportMenuCommands = MakeUID("{291B2148-0495-43AC-9108-F9891917A0D1}")
End Property

Public Property Get GxSearchEngines() As esriSystem.IUID
  Set GxSearchEngines = MakeUID("{1236649F-CD2B-11D3-A6F0-0008C7D3AE50}")
End Property

Public Property Get GxShapefileFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxShapefileFeatureClassContextMenuCommands = MakeUID("{E10D6EC0-06E6-11D3-9F87-00C04F6BC626}")
End Property

Public Property Get GxShapefileFeatureClassExportMenuCommands() As esriSystem.IUID
  Set GxShapefileFeatureClassExportMenuCommands = MakeUID("{E10D6EC2-06E6-11D3-9F87-00C04F6BC626}")
End Property

Public Property Get GxShapefileTableContextMenuCommands() As esriSystem.IUID
  Set GxShapefileTableContextMenuCommands = MakeUID("{E10D6EC1-06E6-11D3-9F87-00C04F6BC626}")
End Property

Public Property Get GxShapefileTableExportMenuCommands() As esriSystem.IUID
  Set GxShapefileTableExportMenuCommands = MakeUID("{E10D6EC3-06E6-11D3-9F87-00C04F6BC626}")
End Property

Public Property Get GxShapefileTableToFeatureClassMenuCommands() As esriSystem.IUID
  Set GxShapefileTableToFeatureClassMenuCommands = MakeUID("{124C8C26-301B-11D4-9FC8-00C04F6BDF06}")
End Property

Public Property Get GxShortcutContextMenuCommands() As esriSystem.IUID
  Set GxShortcutContextMenuCommands = MakeUID("{AB6C1071-13FC-4274-8287-A8AA35A1011A}")
End Property

Public Property Get GxSpatialReferencesFolderContextMenuCommands() As esriSystem.IUID
  Set GxSpatialReferencesFolderContextMenuCommands = MakeUID("{25C0E6F3-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxSpatialReferencesFolderNewMenuCommands() As esriSystem.IUID
  Set GxSpatialReferencesFolderNewMenuCommands = MakeUID("{9828BC65-3AD7-11D3-A62F-0008C7D3AE8D}")
End Property

Public Property Get GxStandaloneFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxStandaloneFeatureClassContextMenuCommands = MakeUID("{22211295-7778-11D3-A00D-00C04F6BC626}")
End Property

Public Property Get GxStreetMapFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxStreetMapFeatureClassContextMenuCommands = MakeUID("{3B5C1D75-C440-4EBF-B629-ACCAD41BF396}")
End Property

Public Property Get GxStreetMapFeatureClassExportMenuCommands() As esriSystem.IUID
  Set GxStreetMapFeatureClassExportMenuCommands = MakeUID("{5DBC42A8-3F19-4138-BDFE-C8EB60456A0F}")
End Property

Public Property Get GxStreetMapFeatureClassPropertyPages() As esriSystem.IUID
  Set GxStreetMapFeatureClassPropertyPages = MakeUID("{F1B979F7-AE98-478A-8A80-C731B8C8F3E9}")
End Property

Public Property Get GxTableContextMenuCommands() As esriSystem.IUID
  Set GxTableContextMenuCommands = MakeUID("{25C0E6FA-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxTableExportMenuCommands() As esriSystem.IUID
  Set GxTableExportMenuCommands = MakeUID("{16FD2B1D-1556-11D3-9FA7-00C04F6BC626}")
End Property

Public Property Get GxTableToFeatureClassMenuCommands() As esriSystem.IUID
  Set GxTableToFeatureClassMenuCommands = MakeUID("{20580FBF-2DDE-11D4-9FC8-00C04F6BDF06}")
End Property

Public Property Get GxTabViews() As esriSystem.IUID
  Set GxTabViews = MakeUID("{3EEA2CAF-A730-11D2-AF6D-080009EC734B}")
End Property

Public Property Get GxTinContextMenuCommands() As esriSystem.IUID
  Set GxTinContextMenuCommands = MakeUID("{25C0E6FF-CD06-11D2-9F40-00C04F6BC626}")
End Property

Public Property Get GxViews() As esriSystem.IUID
  Set GxViews = MakeUID("{3EEA2CB1-A730-11D2-AF6D-080009EC734B}")
End Property

Public Property Get GxVpfCoverageContextMenuCommands() As esriSystem.IUID
  Set GxVpfCoverageContextMenuCommands = MakeUID("{832A7CA1-F47F-11D3-9B56-00C04FA33299}")
End Property

Public Property Get GxVpfCoverageExportMenuCommands() As esriSystem.IUID
  Set GxVpfCoverageExportMenuCommands = MakeUID("{832A7CA3-F47F-11D3-9B56-00C04FA33299}")
End Property

Public Property Get GxVpfCoveragePropertyPages() As esriSystem.IUID
  Set GxVpfCoveragePropertyPages = MakeUID("{832A7CA5-F47F-11D3-9B56-00C04FA33299}")
End Property

Public Property Get GxVpfFeatureClassContextMenuCommands() As esriSystem.IUID
  Set GxVpfFeatureClassContextMenuCommands = MakeUID("{832A7CA2-F47F-11D3-9B56-00C04FA33299}")
End Property

Public Property Get GxVpfFeatureClassExportMenuCommands() As esriSystem.IUID
  Set GxVpfFeatureClassExportMenuCommands = MakeUID("{832A7CA4-F47F-11D3-9B56-00C04FA33299}")
End Property

Public Property Get GxVpfFeatureClassPropertyPages() As esriSystem.IUID
  Set GxVpfFeatureClassPropertyPages = MakeUID("{832A7CA6-F47F-11D3-9B56-00C04FA33299}")
End Property

Public Property Get GxVpfTableContextMenuCommands() As esriSystem.IUID
  Set GxVpfTableContextMenuCommands = MakeUID("{1CB448E4-F485-11D3-9B56-00C04FA33299}")
End Property

Public Property Get GxVpfTableExportMenuCommands() As esriSystem.IUID
  Set GxVpfTableExportMenuCommands = MakeUID("{1CB448E5-F485-11D3-9B56-00C04FA33299}")
End Property

Public Property Get GxVpfTablePropertyPages() As esriSystem.IUID
  Set GxVpfTablePropertyPages = MakeUID("{1CB448E6-F485-11D3-9B56-00C04FA33299}")
End Property

Public Property Get IdentifyObj() As esriSystem.IUID
  Set IdentifyObj = MakeUID("{997E7AC0-D353-11D2-8CD0-00C04F5B951E}")
End Property

Public Property Get ImportPathPropertyPages() As esriSystem.IUID
  Set ImportPathPropertyPages = MakeUID("{FEE7FBA1-C40E-11D5-B2A2-00508BCDDE28}")
End Property

Public Property Get IMSMetadataServicePropertyPages() As esriSystem.IUID
  Set IMSMetadataServicePropertyPages = MakeUID("{079E9B6C-DA59-44EB-B129-E013139D3812}")
End Property

Public Property Get InfoTablePropertyPages() As esriSystem.IUID
  Set InfoTablePropertyPages = MakeUID("{5F7BBE85-E6EC-11D2-9F30-00C04F79927C}")
End Property

Public Property Get LabelStylePages() As esriSystem.IUID
  Set LabelStylePages = MakeUID("{4C90DE7D-CB77-11D2-9F34-00C04F6BC6A5}")
End Property

Public Property Get LayerFactory() As esriSystem.IUID
  Set LayerFactory = MakeUID("{34C20001-4D3C-11D0-92D8-00805F7C28B0}")
End Property

Public Property Get LayerLabelsPropertyPages() As esriSystem.IUID
  Set LayerLabelsPropertyPages = MakeUID("{D80B7114-739C-4ECB-AD06-4DF67C0E6596}")
End Property

Public Property Get LayerPropertyPages() As esriSystem.IUID
  Set LayerPropertyPages = MakeUID("{1476C782-6F57-11D2-A2C6-080009B6F22B}")
End Property

Public Property Get LayerPropertySheet() As esriSystem.IUID
  Set LayerPropertySheet = MakeUID("{DC70783B-5CE1-11D0-92DA-00805F7C28B0}")
End Property

Public Property Get LegendItemPropSheets() As esriSystem.IUID
  Set LegendItemPropSheets = MakeUID("{4E91E870-4C86-11D1-880B-080009EC732A}")
End Property

Public Property Get LegendItems() As esriSystem.IUID
  Set LegendItems = MakeUID("{4E91E86F-4C86-11D1-880B-080009EC732A}")
End Property

Public Property Get LegendSymbolEditors() As esriSystem.IUID
  Set LegendSymbolEditors = MakeUID("{9C275BE0-2A66-11D3-80E3-0080C79F0371}")
End Property

Public Property Get LineDecorationElement() As esriSystem.IUID
  Set LineDecorationElement = MakeUID("{5673B810-2B1F-11D3-80E5-0080C79F0371}")
End Property

Public Property Get LinePatches() As esriSystem.IUID
  Set LinePatches = MakeUID("{F21E3E7C-E3BA-11D2-B868-00600802E603}")
End Property

Public Property Get LineSymbol() As esriSystem.IUID
  Set LineSymbol = MakeUID("{5424FB00-5713-11D2-97E5-0080C7E04196}")
End Property

Public Property Get MapGridBorders() As esriSystem.IUID
  Set MapGridBorders = MakeUID("{8AF09A6E-A0DD-11D2-AE7F-080009EC732A}")
End Property

Public Property Get MapGridFactory() As esriSystem.IUID
  Set MapGridFactory = MakeUID("{A0C55AF7-F9D5-11D1-AA8E-000000000000}")
End Property

Public Property Get MapGridPropertyPages() As esriSystem.IUID
  Set MapGridPropertyPages = MakeUID("{452A7917-1B8E-11D3-9F92-00C04F6BC78E}")
End Property

Public Property Get MapPropertyPages() As esriSystem.IUID
  Set MapPropertyPages = MakeUID("{AE76C8D2-A58A-11D2-A2FE-080009B6F22B}")
End Property

Public Property Get MarkerSymbol() As esriSystem.IUID
  Set MarkerSymbol = MakeUID("{5424FB01-5713-11D2-97E5-0080C7E04196}")
End Property

Public Property Get MBDerivedGeoDatasetProp() As esriSystem.IUID
  Set MBDerivedGeoDatasetProp = MakeUID("{A7386B03-775E-11D3-9F54-00C04F8ED1D7}")
End Property

Public Property Get MBGeoDatasetProp() As esriSystem.IUID
  Set MBGeoDatasetProp = MakeUID("{A7386B01-775E-11D3-9F54-00C04F8ED1D7}")
End Property

Public Property Get MBGeoOperationProp() As esriSystem.IUID
  Set MBGeoOperationProp = MakeUID("{A7386B05-775E-11D3-9F54-00C04F8ED1D7}")
End Property

Public Property Get MBLayerProp() As esriSystem.IUID
  Set MBLayerProp = MakeUID("{A7386B02-775E-11D3-9F54-00C04F8ED1D7}")
End Property

Public Property Get MBMethodProp() As esriSystem.IUID
  Set MBMethodProp = MakeUID("{A7386B06-775E-11D3-9F54-00C04F8ED1D7}")
End Property

Public Property Get MBProcessProp() As esriSystem.IUID
  Set MBProcessProp = MakeUID("{A7386B07-775E-11D3-9F54-00C04F8ED1D7}")
End Property

Public Property Get MBTableProp() As esriSystem.IUID
  Set MBTableProp = MakeUID("{A7386B04-775E-11D3-9F54-00C04F8ED1D7}")
End Property

Public Property Get MetadataEditor() As esriSystem.IUID
  Set MetadataEditor = MakeUID("{7AD0DA0A-0D4A-11D3-A626-0008C7D3AE50}")
End Property

Public Property Get MetadataEditorPropertyPages() As esriSystem.IUID
  Set MetadataEditorPropertyPages = MakeUID("{5907D721-04C9-11D3-83FF-00C04F8ED211}")
End Property

Public Property Get MetadataSynchronizers() As esriSystem.IUID
  Set MetadataSynchronizers = MakeUID("{A802F1A4-E795-4169-997C-6B8FE2F53F98}")
End Property

Public Property Get MxCommandBars() As esriSystem.IUID
  Set MxCommandBars = MakeUID("{B56A7C4A-83D4-11D2-A2E9-080009B6F22B}")
End Property

Public Property Get MxCommands() As esriSystem.IUID
  Set MxCommands = MakeUID("{B56A7C42-83D4-11D2-A2E9-080009B6F22B}")
End Property

Public Property Get MxDDECommandHandler() As esriSystem.IUID
  Set MxDDECommandHandler = MakeUID("{F440A933-565C-447C-8C56-28BB0404E5B9}")
End Property

Public Property Get MxDockableWindows() As esriSystem.IUID
  Set MxDockableWindows = MakeUID("{117623B5-F9D1-11D3-A67F-0008C7DF97B9}")
End Property

Public Property Get MxDocumentDropTarget() As esriSystem.IUID
  Set MxDocumentDropTarget = MakeUID("{9168AC07-5DBE-11D4-A689-0008C7DF88DB}")
End Property

Public Property Get MxExtension() As esriSystem.IUID
  Set MxExtension = MakeUID("{B56A7C45-83D4-11D2-A2E9-080009B6F22B}")
End Property

Public Property Get MxFileMenuCommands() As esriSystem.IUID
  Set MxFileMenuCommands = MakeUID("{F8FC396D-5BA0-44A8-98C4-96C70C5312DE}")
End Property

Public Property Get MxFileMenuDefs() As esriSystem.IUID
  Set MxFileMenuDefs = MakeUID("{6F400F1B-1B8F-4469-9581-A0A3826B9FD5}")
End Property

Public Property Get MxFinders() As esriSystem.IUID
  Set MxFinders = MakeUID("{8803C8EF-75FE-11D3-A6A6-0008C7D3AE50}")
End Property

Public Property Get MxInsetWindowPages() As esriSystem.IUID
  Set MxInsetWindowPages = MakeUID("{67713804-1368-4C15-A064-639396316EC8}")
End Property

Public Property Get MxPalettes() As esriSystem.IUID
  Set MxPalettes = MakeUID("{B56A7C48-83D4-11D2-A2E9-080009B6F22B}")
End Property

Public Property Get MxToolMenuCommands() As esriSystem.IUID
  Set MxToolMenuCommands = MakeUID("{6B1DF87B-DEC8-49B0-884F-345FE2EB1274}")
End Property

Public Property Get MxViewCommands() As esriSystem.IUID
  Set MxViewCommands = MakeUID("{B56A7C43-83D4-11D2-A2E9-080009B6F22B}")
End Property

Public Property Get NameStringParser() As esriSystem.IUID
  Set NameStringParser = MakeUID("{7BF26B1F-C776-11D2-9F24-00C04F6BC69E}")
End Property

Public Property Get NativeTypes() As esriSystem.IUID
  Set NativeTypes = MakeUID("{D80D64ED-0B0F-4C45-B68C-F71F1C310353}")
End Property

Public Property Get NumberFormatPropertyPages() As esriSystem.IUID
  Set NumberFormatPropertyPages = MakeUID("{88002C0B-939A-11D2-AE73-080009EC732A}")
End Property

Public Property Get OptionsPropertyPages() As esriSystem.IUID
  Set OptionsPropertyPages = MakeUID("{FB8C2199-9D0A-11D3-A645-0008C7DF8DE1}")
End Property

Public Property Get Pictures() As esriSystem.IUID
  Set Pictures = MakeUID("{289C9956-C063-11D2-9F22-00C04F6BC8DD}")
End Property

Public Property Get PlugInWorkspaceFactoryHelpers() As esriSystem.IUID
  Set PlugInWorkspaceFactoryHelpers = MakeUID("{6871A79C-CD65-46F5-A7BB-88F97D0E2CDC}")
End Property

Public Property Get Pre70CoveragePropertyPages() As esriSystem.IUID
  Set Pre70CoveragePropertyPages = MakeUID("{C65A2BB7-32ED-11D3-9F33-00C04F79927C}")
End Property

Public Property Get PrinterDrivers() As esriSystem.IUID
  Set PrinterDrivers = MakeUID("{DF933391-7A7A-11D2-ACF9-0000F87808EE}")
End Property

Public Property Get PrinterPropertyPages() As esriSystem.IUID
  Set PrinterPropertyPages = MakeUID("{9A797CC1-5E01-11D2-9181-0000F87808EE}")
End Property

Public Property Get ProjectedCoordSysPropertyPages() As esriSystem.IUID
  Set ProjectedCoordSysPropertyPages = MakeUID("{7E9913E2-E22E-11D2-99C0-0000F80372B4}")
End Property

Public Property Get PropertyPages() As esriSystem.IUID
  Set PropertyPages = MakeUID("{C899AE27-E515-11D1-877C-0000F8751720}")
End Property

Public Property Get RasterAnalysis() As esriSystem.IUID
  Set RasterAnalysis = MakeUID("{59CA6922-47FB-11D2-8D29-0000F8780535}")
End Property

Public Property Get RasterBandPropertyPages() As esriSystem.IUID
  Set RasterBandPropertyPages = MakeUID("{DC942D39-E3A8-11D2-9F52-00C04F8ED21A}")
End Property

Public Property Get RasterConditionalOp() As esriSystem.IUID
  Set RasterConditionalOp = MakeUID("{09FF656C-F398-11D3-9F9B-00C04F8EF00E}")
End Property

Public Property Get RasterConversionOp() As esriSystem.IUID
  Set RasterConversionOp = MakeUID("{CD568B5A-CA31-11D2-9F3C-00C04F8ED1D7}")
End Property

Public Property Get RasterDataExclusionPropSheet() As esriSystem.IUID
  Set RasterDataExclusionPropSheet = MakeUID("{9A90E0A4-6265-11D3-9FD1-00C04F68E699}")
End Property

Public Property Get RasterDatasetPropertyPages() As esriSystem.IUID
  Set RasterDatasetPropertyPages = MakeUID("{74F98CC2-E2E7-11D2-9F51-00C04F8ED21A}")
End Property

Public Property Get RasterDensityOp() As esriSystem.IUID
  Set RasterDensityOp = MakeUID("{EB3F5D32-1306-11D3-9F45-00C04F8ED1D7}")
End Property

Public Property Get RasterDistanceOp() As esriSystem.IUID
  Set RasterDistanceOp = MakeUID("{4606F178-B47D-11D2-9F3B-00C04F8ED1D7}")
End Property

Public Property Get RasterExportMenuCommands() As esriSystem.IUID
  Set RasterExportMenuCommands = MakeUID("{837FAC81-38C9-11D4-8D94-00C04F5B87B2}")
End Property

Public Property Get RasterExtractionOp() As esriSystem.IUID
  Set RasterExtractionOp = MakeUID("{09FF656D-F398-11D3-9F9B-00C04F8EF00E}")
End Property

Public Property Get RasterGeneralizeOp() As esriSystem.IUID
  Set RasterGeneralizeOp = MakeUID("{EB3F5D31-1306-11D3-9F45-00C04F8ED1D7}")
End Property

Public Property Get RasterGxExportFilters() As esriSystem.IUID
  Set RasterGxExportFilters = MakeUID("{8837B6E0-7F9E-11D4-B284-00508BCDC7C8}")
End Property

Public Property Get RasterHydrologyOp() As esriSystem.IUID
  Set RasterHydrologyOp = MakeUID("{EB3F5D2C-1306-11D3-9F45-00C04F8ED1D7}")
End Property

Public Property Get RasterInterpolationOp() As esriSystem.IUID
  Set RasterInterpolationOp = MakeUID("{3297E9D4-93A1-11D2-9F33-00C04F8ED1D7}")
End Property

Public Property Get RasterLayerPropertyPages() As esriSystem.IUID
  Set RasterLayerPropertyPages = MakeUID("{2AD25BF1-3C02-11D3-A3FA-0004AC1B1D86}")
End Property

Public Property Get RasterLocalOp() As esriSystem.IUID
  Set RasterLocalOp = MakeUID("{83370E8B-9F6B-11D2-9F35-00C04F8ED1D7}")
End Property

Public Property Get RasterMapAlgebraOp() As esriSystem.IUID
  Set RasterMapAlgebraOp = MakeUID("{3297E9C3-93A1-11D2-9F33-00C04F8ED1D7}")
End Property

Public Property Get RasterMathOps() As esriSystem.IUID
  Set RasterMathOps = MakeUID("{4606F177-B47D-11D2-9F3B-00C04F8ED1D7}")
End Property

Public Property Get RasterNeighborhoodOp() As esriSystem.IUID
  Set RasterNeighborhoodOp = MakeUID("{4606F17A-B47D-11D2-9F3B-00C04F8ED1D7}")
End Property

Public Property Get RasterOverlayOp() As esriSystem.IUID
  Set RasterOverlayOp = MakeUID("{EB3F5D2E-1306-11D3-9F45-00C04F8ED1D7}")
End Property

Public Property Get RasterRasterMakerOp() As esriSystem.IUID
  Set RasterRasterMakerOp = MakeUID("{CC0B70D8-EA20-11D3-9F63-00C04F8ED1D7}")
End Property

Public Property Get RasterReclassOp() As esriSystem.IUID
  Set RasterReclassOp = MakeUID("{EB3F5D2D-1306-11D3-9F45-00C04F8ED1D7}")
End Property

Public Property Get RasterRendererMakers() As esriSystem.IUID
  Set RasterRendererMakers = MakeUID("{F38F6DB0-C73E-11D4-B290-00508BCDC7C8}")
End Property

Public Property Get RasterStatEnvDialog() As esriSystem.IUID
  Set RasterStatEnvDialog = MakeUID("{80F65394-FD12-11D2-9F63-00C04F68E699}")
End Property

Public Property Get RasterStatisticsOp() As esriSystem.IUID
  Set RasterStatisticsOp = MakeUID("{EB3F5D2F-1306-11D3-9F45-00C04F8ED1D7}")
End Property

Public Property Get RasterSurfaceOp() As esriSystem.IUID
  Set RasterSurfaceOp = MakeUID("{4606F179-B47D-11D2-9F3B-00C04F8ED1D7}")
End Property

Public Property Get RasterTransformationOp() As esriSystem.IUID
  Set RasterTransformationOp = MakeUID("{EB3F5D30-1306-11D3-9F45-00C04F8ED1D7}")
End Property

Public Property Get RasterZonalOp() As esriSystem.IUID
  Set RasterZonalOp = MakeUID("{4606F17B-B47D-11D2-9F3B-00C04F8ED1D7}")
End Property

Public Property Get RelationshipClassPropertyPages() As esriSystem.IUID
  Set RelationshipClassPropertyPages = MakeUID("{05318B65-DE31-11D2-9F54-00C04F6BC626}")
End Property

Public Property Get RendererPropertyPages() As esriSystem.IUID
  Set RendererPropertyPages = MakeUID("{962BD9A9-1EC8-11D3-9F4D-00C04F6BC709}")
End Property

Public Property Get SAExtensions() As esriSystem.IUID
  Set SAExtensions = MakeUID("{BA648A2A-B79E-433C-999A-AA3C7B075E17}")
End Property

Public Property Get SAOperation() As esriSystem.IUID
  Set SAOperation = MakeUID("{FAF0CFBB-CD00-11D3-A051-00C04F68E699}")
End Property

Public Property Get ScaleBars() As esriSystem.IUID
  Set ScaleBars = MakeUID("{34DC96D1-186D-11D3-B8A3-00600802E603}")
End Property

Public Property Get SceneExporter3dDrivers() As esriSystem.IUID
  Set SceneExporter3dDrivers = MakeUID("{D6AC48F8-37F4-11D4-A168-444553547777}")
End Property

Public Property Get SceneExporter3dPropertyPages() As esriSystem.IUID
  Set SceneExporter3dPropertyPages = MakeUID("{D6AC48FC-37F4-11D4-A168-444553547777}")
End Property

Public Property Get SceneExporterVideoDrivers() As esriSystem.IUID
  Set SceneExporterVideoDrivers = MakeUID("{4251BA09-9106-11D5-B29F-00508BCDDE28}")
End Property

Public Property Get SceneExporterVideoPropertyPages() As esriSystem.IUID
  Set SceneExporterVideoPropertyPages = MakeUID("{4251BA0A-9106-11D5-B29F-00508BCDDE28}")
End Property

Public Property Get ScenePropertyPages() As esriSystem.IUID
  Set ScenePropertyPages = MakeUID("{CF228351-10A9-11D4-9F7F-00C04F6BC619}")
End Property

Public Property Get SearchResultsContextMenuCommands() As esriSystem.IUID
  Set SearchResultsContextMenuCommands = MakeUID("{13A17004-CA63-46D3-B8CF-031F4C186B3C}")
End Property

Public Property Get Shadows() As esriSystem.IUID
  Set Shadows = MakeUID("{4DE36362-E872-4159-B2DB-3ECEA4F57F0E}")
End Property

Public Property Get ShapefileFeatureClassPropertyPages() As esriSystem.IUID
  Set ShapefileFeatureClassPropertyPages = MakeUID("{E10D6EC4-06E6-11D3-9F87-00C04F6BC626}")
End Property

Public Property Get ShapefileTablePropertyPages() As esriSystem.IUID
  Set ShapefileTablePropertyPages = MakeUID("{E10D6EC5-06E6-11D3-9F87-00C04F6BC626}")
End Property

Public Property Get SketchMenuCommands() As esriSystem.IUID
  Set SketchMenuCommands = MakeUID("{3DD0D4DD-01E6-41AC-BB50-108D11C763A7}")
End Property

Public Property Get SketchToolMenuCommands() As esriSystem.IUID
  Set SketchToolMenuCommands = MakeUID("{3DD0D4DC-01E6-41AC-BB50-108D11C763A7}")
End Property

Public Property Get SnapAgents() As esriSystem.IUID
  Set SnapAgents = MakeUID("{5844F9F6-DC03-11D2-9F26-00C04F6BC979}")
End Property

Public Property Get StandaloneTablePropertyPages() As esriSystem.IUID
  Set StandaloneTablePropertyPages = MakeUID("{DA7C157D-E58B-11D3-ADEA-00C04FA33A15}")
End Property

Public Property Get StyleGalleryClasses() As esriSystem.IUID
  Set StyleGalleryClasses = MakeUID("{AC0E982B-91CB-11D1-8813-080009EC732A}")
End Property

Public Property Get StyleImporter() As esriSystem.IUID
  Set StyleImporter = MakeUID("{17049F82-8E15-11D2-983E-0080C7E04196}")
End Property

Public Property Get StylePropertyPages() As esriSystem.IUID
  Set StylePropertyPages = MakeUID("{B6DD2405-29D6-11D3-B8AD-00600802E603}")
End Property

Public Property Get SxCommandBars() As esriSystem.IUID
  Set SxCommandBars = MakeUID("{F27D8291-A383-11D3-8206-0080C7597E71}")
End Property

Public Property Get SxCommands() As esriSystem.IUID
  Set SxCommands = MakeUID("{F27D8292-A383-11D3-8206-0080C7597E71}")
End Property

Public Property Get SxContentsViews() As esriSystem.IUID
  Set SxContentsViews = MakeUID("{BB07DB51-E57B-11D3-9F74-00C04F6BC619}")
End Property

Public Property Get SxContextAnalyzers() As esriSystem.IUID
  Set SxContextAnalyzers = MakeUID("{3BB42A1D-05A0-11D4-9F79-00C04F6BC619}")
End Property

Public Property Get SxDockableWindows() As esriSystem.IUID
  Set SxDockableWindows = MakeUID("{B9789374-FADC-11D3-A67F-0008C7DF97B9}")
End Property

Public Property Get SxExtensions() As esriSystem.IUID
  Set SxExtensions = MakeUID("{369448D7-A6BD-11D3-ABB8-0008C73FCA1C}")
End Property

Public Property Get SxFinders() As esriSystem.IUID
  Set SxFinders = MakeUID("{B3E00B2A-25D3-11D4-A383-00C04F6BC619}")
End Property

Public Property Get SxLayerPropertyPages() As esriSystem.IUID
  Set SxLayerPropertyPages = MakeUID("{3F82C99B-1C5F-11D4-A381-00C04F6BC619}")
End Property

Public Property Get SxOptionsPropertyPages() As esriSystem.IUID
  Set SxOptionsPropertyPages = MakeUID("{F58B83FF-31AE-11D4-A385-00C04F6BC619}")
End Property

Public Property Get SxPalettes() As esriSystem.IUID
  Set SxPalettes = MakeUID("{A51B87D1-A383-11D3-8206-0080C7597E71}")
End Property

Public Property Get SxPropertyPages() As esriSystem.IUID
  Set SxPropertyPages = MakeUID("{779E31F8-25E1-11D4-9B80-00C04FA33299}")
End Property

Public Property Get SxRootObjects() As esriSystem.IUID
  Set SxRootObjects = MakeUID("{5A10D191-A385-11D3-8206-0080C7597E71}")
End Property

Public Property Get SymbolPropertyPages() As esriSystem.IUID
  Set SymbolPropertyPages = MakeUID("{818B37C0-F34E-11D2-BC8F-0080C7E04196}")
End Property

Public Property Get TablePropertyPages() As esriSystem.IUID
  Set TablePropertyPages = MakeUID("{750D818D-E0B5-11D2-99BF-0000F80372B4}")
End Property

Public Property Get TableWindowPropertyPages() As esriSystem.IUID
  Set TableWindowPropertyPages = MakeUID("{DA51F911-EA46-4E57-9B9E-2CF443D34CAB}")
End Property

Public Property Get TextBackground() As esriSystem.IUID
  Set TextBackground = MakeUID("{F8AEACE0-25D1-11D3-80DA-0080C79F0371}")
End Property

Public Property Get TextBackgroundPropertyPages() As esriSystem.IUID
  Set TextBackgroundPropertyPages = MakeUID("{F1CAF8B0-25C3-11D3-80DA-0080C79F0371}")
End Property

Public Property Get TextFilePropertyPages() As esriSystem.IUID
  Set TextFilePropertyPages = MakeUID("{090BE115-10AB-11D4-AE05-00C04FA33A15}")
End Property

Public Property Get TextSymbol() As esriSystem.IUID
  Set TextSymbol = MakeUID("{5424FB03-5713-11D2-97E5-0080C7E04196}")
End Property

Public Property Get TinDatasetPropertyPages() As esriSystem.IUID
  Set TinDatasetPropertyPages = MakeUID("{2A3354C9-24E1-11D3-9F4C-00C04F6BC619}")
End Property

Public Property Get TinRendererPropertyPages() As esriSystem.IUID
  Set TinRendererPropertyPages = MakeUID("{19283BDF-656F-11D3-9F57-00C04F6BC619}")
End Property

Public Property Get TraceTasks() As esriSystem.IUID
  Set TraceTasks = MakeUID("{5BF495FA-95D1-11D2-AAEB-00C04FA37D59}")
End Property

Public Property Get TransformationMethods() As esriSystem.IUID
  Set TransformationMethods = MakeUID("{D9246646-BEDB-4D2B-88B0-5B94D1BBD5A7}")
End Property

Public Property Get UtilityNetworkAnalysisGNPropertyPages() As esriSystem.IUID
  Set UtilityNetworkAnalysisGNPropertyPages = MakeUID("{69BCACD9-019C-11D3-9F7D-00C04F6BDF0E}")
End Property

Public Property Get UtilityNetworkAnalysisPropertyPages() As esriSystem.IUID
  Set UtilityNetworkAnalysisPropertyPages = MakeUID("{97499635-F279-11D2-9F6A-00C04F6BDF0E}")
End Property

Public Property Get UtilityNetworkTasks() As esriSystem.IUID
  Set UtilityNetworkTasks = MakeUID("{39DC1B4F-BE34-11D2-BABE-00C04FA33C20}")
End Property

Public Property Get VersionPropertyPages() As esriSystem.IUID
  Set VersionPropertyPages = MakeUID("{0224F7BC-F736-11D2-9F2E-00C04F6BC979}")
End Property

Public Property Get WorkspaceDomainsPropertyPages() As esriSystem.IUID
  Set WorkspaceDomainsPropertyPages = MakeUID("{C255D347-CB5C-11D2-9F3C-00C04F6BC626}")
End Property

Public Property Get WorkspaceFactory() As esriSystem.IUID
  Set WorkspaceFactory = MakeUID("{07BC056E-DB8A-11D1-AA81-00C04FA33A15}")
End Property

Public Property Get WorkspacePropertyPages() As esriSystem.IUID
  Set WorkspacePropertyPages = MakeUID("{C255D343-CB5C-11D2-9F3C-00C04F6BC626}")
End Property

'***************************************************************************
'Name:                  CatalogComponentNames
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Initialize collection of component names
'Called From:   Class_Initialize
'Description:   None
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
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Private Sub CatalogComponentNames()
    With m_colComponentNames
        .Add "AdjustmentMethodPropertyPages", "AdjustmentMethodPropertyPages"
        .Add "AdjustmentPropertyPages", "AdjustmentPropertyPages"
        .Add "AnimationPropertyPages", "AnimationPropertyPages"
        .Add "AnnotateMapPropertyPages", "AnnotateMapPropertyPages"
        .Add "AnnotationEngines", "AnnotationEngines"
        .Add "AnnotationExpresionPages", "AnnotationExpresionPages"
        .Add "AnnotationExpressionParsers", "AnnotationExpressionParsers"
        .Add "AnnotationPlacementPages", "AnnotationPlacementPages"
        .Add "AnnotationPropertiesConverters", "AnnotationPropertiesConverters"
        .Add "ArcToolboxTools", "ArcToolboxTools"
        .Add "AreaPatches", "AreaPatches"
        .Add "AttributeTransferTypes", "AttributeTransferTypes"
        .Add "AVFSrcImporters", "AVFSrcImporters"
        .Add "AVThemeImporters", "AVThemeImporters"
        .Add "Backgrounds", "Backgrounds"
        .Add "Borders", "Borders"
        .Add "CadDrawingPropertyPages", "CadDrawingPropertyPages"
        .Add "CadFeatureClassPropertyPages", "CadFeatureClassPropertyPages"
        .Add "CadFeaturePropertyPages", "CadFeaturePropertyPages"
        .Add "ClassificationMethods", "ClassificationMethods"
        .Add "ClipboardFormats", "ClipboardFormats"
        .Add "ColorCorrectionPages", "ColorCorrectionPages"
        .Add "ColorPropertyPages", "ColorPropertyPages"
        .Add "ColorRamp", "ColorRamp"
        .Add "ColorRampPropetyPages", "ColorRampPropetyPages"
        .Add "ContentsViews", "ContentsViews"
        .Add "ContextAnalyzers", "ContextAnalyzers"
        .Add "CoverageFeatureClassPropertyPages", "CoverageFeatureClassPropertyPages"
        .Add "CoveragePropertyPages", "CoveragePropertyPages"
        .Add "DataExclusionPropertyPages", "DataExclusionPropertyPages"
        .Add "DataGraphPropertyPages", "DataGraphPropertyPages"
        .Add "DataSamplingPropertyPages", "DataSamplingPropertyPages"
        .Add "DataWindowFactory", "DataWindowFactory"
        .Add "DDDAnimationTypes", "DDDAnimationTypes"
        .Add "DDDExtensions", "DDDExtensions"
        .Add "DimensionConstructors", "DimensionConstructors"
        .Add "EditorExtensions", "EditorExtensions"
        .Add "EditorPropertyPages", "EditorPropertyPages"
        .Add "EditTasks", "EditTasks"
        .Add "ElementPropertyPages", "ElementPropertyPages"
        .Add "ESDAEngineMethods", "ESDAEngineMethods"
        .Add "ExporterDrivers", "ExporterDrivers"
        .Add "ExporterPropertyPages", "ExporterPropertyPages"
        .Add "FeatureClassPropertyPages", "FeatureClassPropertyPages"
        .Add "FeatureDatasetPropertyPages", "FeatureDatasetPropertyPages"
        .Add "FeatureLayerSymbology", "FeatureLayerSymbology"
        .Add "FillSymbol", "FillSymbol"
        .Add "FontMappingPropertyPages", "FontMappingPropertyPages"
        .Add "FrameElementPropertyPages", "FrameElementPropertyPages"
        .Add "GeoAnalysis", "GeoAnalysis"
        .Add "GeodatabaseWorkspaceExtensions", "GeodatabaseWorkspaceExtensions"
        .Add "GeoElementTools", "GeoElementTools"
        .Add "GeographicCoordSysPropertyPages", "GeographicCoordSysPropertyPages"
        .Add "GeometricNetworkPropertyPages", "GeometricNetworkPropertyPages"
        .Add "GeoObjectClassDescriptions", "GeoObjectClassDescriptions"
        .Add "GeoObjectClassExtensions", "GeoObjectClassExtensions"
        .Add "GeoObjects", "GeoObjects"
        .Add "GeoRefSaveAsDlg", "GeoRefSaveAsDlg"
        .Add "GeostatExtensions", "GeostatExtensions"
        .Add "GNEditorPropertyPages", "GNEditorPropertyPages"
        .Add "GraphicSnapAgents", "GraphicSnapAgents"
        .Add "GridLabels", "GridLabels"
        .Add "GSEngineMethods", "GSEngineMethods"
        .Add "GSRenderers", "GSRenderers"
        .Add "GxArcViewExportFeatureClassFilters", "GxArcViewExportFeatureClassFilters"
        .Add "GxArcViewExportTableFilters", "GxArcViewExportTableFilters"
        .Add "GxCadDataContextMenuCommands", "GxCadDataContextMenuCommands"
        .Add "GxCadDrawingContextMenuCommands", "GxCadDrawingContextMenuCommands"
        .Add "GxCadDrawingExportMenuCommands", "GxCadDrawingExportMenuCommands"
        .Add "GxCadFeatureClassContextMenuCommands", "GxCadFeatureClassContextMenuCommands"
        .Add "GxCadFeatureClassExportMenuCommands", "GxCadFeatureClassExportMenuCommands"
        .Add "GxCadFeatureContextMenuCommands", "GxCadFeatureContextMenuCommands"
        .Add "GxCadFeatureExportMenuCommands", "GxCadFeatureExportMenuCommands"
        .Add "GxCatalogContextMenuCommands", "GxCatalogContextMenuCommands"
        .Add "GxCatalogPropertyPages", "GxCatalogPropertyPages"
        .Add "GxCommandBars", "GxCommandBars"
        .Add "GxCommands", "GxCommands"
        .Add "GxCoverageContextMenuCommands", "GxCoverageContextMenuCommands"
        .Add "GxCoverageExportMenuCommands", "GxCoverageExportMenuCommands"
        .Add "GxCoverageFeatureClassContextMenuCommands", "GxCoverageFeatureClassContextMenuCommands"
        .Add "GxCoverageFeatureClassExportMenuCommands", "GxCoverageFeatureClassExportMenuCommands"
        .Add "GxCoverageNewMenuCommands", "GxCoverageNewMenuCommands"
        .Add "GxDatabaseContextMenuCommands", "GxDatabaseContextMenuCommands"
        .Add "GxDatabaseExtensions", "GxDatabaseExtensions"
        .Add "GxDatabaseImportMenuCommands", "GxDatabaseImportMenuCommands"
        .Add "GxDatabaseNewMenuCommands", "GxDatabaseNewMenuCommands"
        .Add "GxDataGraphContextMenuCommands", "GxDataGraphContextMenuCommands"
        .Add "GxDatasetContextMenuCommands", "GxDatasetContextMenuCommands"
        .Add "GxDialogObjectFilters", "GxDialogObjectFilters"
        .Add "GxDiskConnectionContextMenuCommands", "GxDiskConnectionContextMenuCommands"
        .Add "GxDockableWindows", "GxDockableWindows"
        .Add "GxEnabledWorkspaceFactories", "GxEnabledWorkspaceFactories"
        .Add "GxEnterpriseFeatureClassContextMenuCommands", "GxEnterpriseFeatureClassContextMenuCommands"
        .Add "GxEnterpriseFeatureClassContextMenuCommandsAV", "GxEnterpriseFeatureClassContextMenuCommandsAV"
        .Add "GxEnterpriseFeatureDatasetContextMenuCommands", "GxEnterpriseFeatureDatasetContextMenuCommands"
        .Add "GxEnterpriseRasterBandContextMenuCommands", "GxEnterpriseRasterBandContextMenuCommands"
        .Add "GxEnterpriseRasterDatasetContextMenuCommands", "GxEnterpriseRasterDatasetContextMenuCommands"
        .Add "GxEnterpriseRelationshipClassContextMenuCommands", "GxEnterpriseRelationshipClassContextMenuCommands"
        .Add "GxEnterpriseStandaloneFeatureClassContextMenuCommands", "GxEnterpriseStandaloneFeatureClassContextMenuCommands"
        .Add "GxEnterpriseStandaloneFeatureClassContextMenuCommandsAV", "GxEnterpriseStandaloneFeatureClassContextMenuCommandsAV"
        .Add "GxEnterpriseTableContextMenuCommands", "GxEnterpriseTableContextMenuCommands"
        .Add "GxEnterpriseTableContextMenuCommandsAV", "GxEnterpriseTableContextMenuCommandsAV"
        .Add "GxExportFeatureClassFilters", "GxExportFeatureClassFilters"
        .Add "GxExportTableFilters", "GxExportTableFilters"
        .Add "GxExtensions", "GxExtensions"
        .Add "GxExtensionToolbar", "GxExtensionToolbar"
        .Add "GxFeatureClassContextMenuCommands", "GxFeatureClassContextMenuCommands"
        .Add "GxFeatureClassExportMenuCommands", "GxFeatureClassExportMenuCommands"
        .Add "GxFeatureDatasetContextMenuCommands", "GxFeatureDatasetContextMenuCommands"
        .Add "GxFeatureDatasetExportMenuCommands", "GxFeatureDatasetExportMenuCommands"
        .Add "GxFeatureDatasetImportMenuCommands", "GxFeatureDatasetImportMenuCommands"
        .Add "GxFeatureDatasetNewMenuCommands", "GxFeatureDatasetNewMenuCommands"
        .Add "GxFileContextMenuCommands", "GxFileContextMenuCommands"
        .Add "GxFilterFeatureClasses", "GxFilterFeatureClasses"
        .Add "GxFilterFeatureDatasets", "GxFilterFeatureDatasets"
        .Add "GxFilterTables", "GxFilterTables"
        .Add "GxFolderContextMenuCommands", "GxFolderContextMenuCommands"
        .Add "GxFolderNewMenuCommands", "GxFolderNewMenuCommands"
        .Add "GxFolderNewMenuCommandsAV", "GxFolderNewMenuCommandsAV"
        .Add "GxGeometricNetworkContextMenuCommands", "GxGeometricNetworkContextMenuCommands"
        .Add "GxGridContextMenuCommands", "GxGridContextMenuCommands"
        .Add "GxInfoTableContextMenuCommands", "GxInfoTableContextMenuCommands"
        .Add "GxInfoTableExportMenuCommands", "GxInfoTableExportMenuCommands"
        .Add "GxInfoTableToFeatureClassMenuCommands", "GxInfoTableToFeatureClassMenuCommands"
        .Add "GxLayerContextMenuCommands", "GxLayerContextMenuCommands"
        .Add "GxLayerFilters", "GxLayerFilters"
        .Add "GxLocatorContextMenuCommands", "GxLocatorContextMenuCommands"
        .Add "GxLocatorFolderContextMenuCommands", "GxLocatorFolderContextMenuCommands"
        .Add "GxLocatorMenus", "GxLocatorMenus"
        .Add "GxMapContextMenuCommands", "GxMapContextMenuCommands"
        .Add "GxMapPropertyPages", "GxMapPropertyPages"
        .Add "GxMetadataExporter", "GxMetadataExporter"
        .Add "GxMetadataImporter", "GxMetadataImporter"
        .Add "GxMultipleSelectionContextMenuCommands", "GxMultipleSelectionContextMenuCommands"
        .Add "GxObjectFactory", "GxObjectFactory"
        .Add "GxObjectFilters", "GxObjectFilters"
        .Add "GxOleDBDatabaseContextMenuCommands", "GxOleDBDatabaseContextMenuCommands"
        .Add "GxOleDBTableContextMenuCommands", "GxOleDBTableContextMenuCommands"
        .Add "GxPalettes", "GxPalettes"
        .Add "GxPCCoverageContextMenuCommands", "GxPCCoverageContextMenuCommands"
        .Add "GxPCCoverageFeatureClassContextMenuCommands", "GxPCCoverageFeatureClassContextMenuCommands"
        .Add "GxPMFMapContextMenuCommands", "GxPMFMapContextMenuCommands"
        .Add "GxPMFMapPropertyPages", "GxPMFMapPropertyPages"
        .Add "GxPre70CoverageContextMenuCommands", "GxPre70CoverageContextMenuCommands"
        .Add "GxPreviews", "GxPreviews"
        .Add "GxPrjFileContextMenuCommands", "GxPrjFileContextMenuCommands"
        .Add "GxRasterBandContextMenuCommands", "GxRasterBandContextMenuCommands"
        .Add "GxRasterDatasetContextMenuCommands", "GxRasterDatasetContextMenuCommands"
        .Add "GxReadOnlyFeatureClassContextMenuCommands", "GxReadOnlyFeatureClassContextMenuCommands"
        .Add "GxReadOnlyFeatureDatasetContextMenuCommands", "GxReadOnlyFeatureDatasetContextMenuCommands"
        .Add "GxReadOnlyStandaloneFeatureClassContextMenuCommands", "GxReadOnlyStandaloneFeatureClassContextMenuCommands"
        .Add "GxReadOnlyTableContextMenuCommands", "GxReadOnlyTableContextMenuCommands"
        .Add "GxRelationshipClassContextMenuCommands", "GxRelationshipClassContextMenuCommands"
        .Add "GxRemoteDatabaseContextMenuCommands", "GxRemoteDatabaseContextMenuCommands"
        .Add "GxRemoteDatabaseFolderContextMenuCommands", "GxRemoteDatabaseFolderContextMenuCommands"
        .Add "GxRootObjects", "GxRootObjects"
        .Add "GxSceneContextMenuCommands", "GxSceneContextMenuCommands"
        .Add "GxScenePropertyPages", "GxScenePropertyPages"
        .Add "GxSDCFeatureClassExportMenuCommands", "GxSDCFeatureClassExportMenuCommands"
        .Add "GxSDCTableExportMenuCommands", "GxSDCTableExportMenuCommands"
        .Add "GxSearchEngines", "GxSearchEngines"
        .Add "GxShapefileFeatureClassContextMenuCommands", "GxShapefileFeatureClassContextMenuCommands"
        .Add "GxShapefileFeatureClassExportMenuCommands", "GxShapefileFeatureClassExportMenuCommands"
        .Add "GxShapefileTableContextMenuCommands", "GxShapefileTableContextMenuCommands"
        .Add "GxShapefileTableExportMenuCommands", "GxShapefileTableExportMenuCommands"
        .Add "GxShapefileTableToFeatureClassMenuCommands", "GxShapefileTableToFeatureClassMenuCommands"
        .Add "GxShortcutContextMenuCommands", "GxShortcutContextMenuCommands"
        .Add "GxSpatialReferencesFolderContextMenuCommands", "GxSpatialReferencesFolderContextMenuCommands"
        .Add "GxSpatialReferencesFolderNewMenuCommands", "GxSpatialReferencesFolderNewMenuCommands"
        .Add "GxStandaloneFeatureClassContextMenuCommands", "GxStandaloneFeatureClassContextMenuCommands"
        .Add "GxStreetMapFeatureClassContextMenuCommands", "GxStreetMapFeatureClassContextMenuCommands"
        .Add "GxStreetMapFeatureClassExportMenuCommands", "GxStreetMapFeatureClassExportMenuCommands"
        .Add "GxStreetMapFeatureClassPropertyPages", "GxStreetMapFeatureClassPropertyPages"
        .Add "GxTableContextMenuCommands", "GxTableContextMenuCommands"
        .Add "GxTableExportMenuCommands", "GxTableExportMenuCommands"
        .Add "GxTableToFeatureClassMenuCommands", "GxTableToFeatureClassMenuCommands"
        .Add "GxTabViews", "GxTabViews"
        .Add "GxTinContextMenuCommands", "GxTinContextMenuCommands"
        .Add "GxViews", "GxViews"
        .Add "GxVpfCoverageContextMenuCommands", "GxVpfCoverageContextMenuCommands"
        .Add "GxVpfCoverageExportMenuCommands", "GxVpfCoverageExportMenuCommands"
        .Add "GxVpfCoveragePropertyPages", "GxVpfCoveragePropertyPages"
        .Add "GxVpfFeatureClassContextMenuCommands", "GxVpfFeatureClassContextMenuCommands"
        .Add "GxVpfFeatureClassExportMenuCommands", "GxVpfFeatureClassExportMenuCommands"
        .Add "GxVpfFeatureClassPropertyPages", "GxVpfFeatureClassPropertyPages"
        .Add "GxVpfTableContextMenuCommands", "GxVpfTableContextMenuCommands"
        .Add "GxVpfTableExportMenuCommands", "GxVpfTableExportMenuCommands"
        .Add "GxVpfTablePropertyPages", "GxVpfTablePropertyPages"
        .Add "IdentifyObj", "IdentifyObj"
        .Add "ImportPathPropertyPages", "ImportPathPropertyPages"
        .Add "IMSMetadataServicePropertyPages", "IMSMetadataServicePropertyPages"
        .Add "InfoTablePropertyPages", "InfoTablePropertyPages"
        .Add "LabelStylePages", "LabelStylePages"
        .Add "LayerFactory", "LayerFactory"
        .Add "LayerLabelsPropertyPages", "LayerLabelsPropertyPages"
        .Add "LayerPropertyPages", "LayerPropertyPages"
        .Add "LayerPropertySheet", "LayerPropertySheet"
        .Add "LegendItemPropSheets", "LegendItemPropSheets"
        .Add "LegendItems", "LegendItems"
        .Add "LegendSymbolEditors", "LegendSymbolEditors"
        .Add "LineDecorationElement", "LineDecorationElement"
        .Add "LinePatches", "LinePatches"
        .Add "LineSymbol", "LineSymbol"
        .Add "MapGridBorders", "MapGridBorders"
        .Add "MapGridFactory", "MapGridFactory"
        .Add "MapGridPropertyPages", "MapGridPropertyPages"
        .Add "MapPropertyPages", "MapPropertyPages"
        .Add "MarkerSymbol", "MarkerSymbol"
        .Add "MBDerivedGeoDatasetProp", "MBDerivedGeoDatasetProp"
        .Add "MBGeoDatasetProp", "MBGeoDatasetProp"
        .Add "MBGeoOperationProp", "MBGeoOperationProp"
        .Add "MBLayerProp", "MBLayerProp"
        .Add "MBMethodProp", "MBMethodProp"
        .Add "MBProcessProp", "MBProcessProp"
        .Add "MBTableProp", "MBTableProp"
        .Add "MetadataEditor", "MetadataEditor"
        .Add "MetadataEditorPropertyPages", "MetadataEditorPropertyPages"
        .Add "MetadataSynchronizers", "MetadataSynchronizers"
        .Add "MxCommandBars", "MxCommandBars"
        .Add "MxCommands", "MxCommands"
        .Add "MxDDECommandHandler", "MxDDECommandHandler"
        .Add "MxDockableWindows", "MxDockableWindows"
        .Add "MxDocumentDropTarget", "MxDocumentDropTarget"
        .Add "MxExtension", "MxExtension"
        .Add "MxFileMenuCommands", "MxFileMenuCommands"
        .Add "MxFileMenuDefs", "MxFileMenuDefs"
        .Add "MxFinders", "MxFinders"
        .Add "MxInsetWindowPages", "MxInsetWindowPages"
        .Add "MxPalettes", "MxPalettes"
        .Add "MxToolMenuCommands", "MxToolMenuCommands"
        .Add "MxViewCommands", "MxViewCommands"
        .Add "NameStringParser", "NameStringParser"
        .Add "NativeTypes", "NativeTypes"
        .Add "NumberFormatPropertyPages", "NumberFormatPropertyPages"
        .Add "OptionsPropertyPages", "OptionsPropertyPages"
        .Add "Pictures", "Pictures"
        .Add "PlugInWorkspaceFactoryHelpers", "PlugInWorkspaceFactoryHelpers"
        .Add "Pre70CoveragePropertyPages", "Pre70CoveragePropertyPages"
        .Add "PrinterDrivers", "PrinterDrivers"
        .Add "PrinterPropertyPages", "PrinterPropertyPages"
        .Add "ProjectedCoordSysPropertyPages", "ProjectedCoordSysPropertyPages"
        .Add "PropertyPages", "PropertyPages"
        .Add "RasterAnalysis", "RasterAnalysis"
        .Add "RasterBandPropertyPages", "RasterBandPropertyPages"
        .Add "RasterConditionalOp", "RasterConditionalOp"
        .Add "RasterConversionOp", "RasterConversionOp"
        .Add "RasterDataExclusionPropSheet", "RasterDataExclusionPropSheet"
        .Add "RasterDatasetPropertyPages", "RasterDatasetPropertyPages"
        .Add "RasterDensityOp", "RasterDensityOp"
        .Add "RasterDistanceOp", "RasterDistanceOp"
        .Add "RasterExportMenuCommands", "RasterExportMenuCommands"
        .Add "RasterExtractionOp", "RasterExtractionOp"
        .Add "RasterGeneralizeOp", "RasterGeneralizeOp"
        .Add "RasterGxExportFilters", "RasterGxExportFilters"
        .Add "RasterHydrologyOp", "RasterHydrologyOp"
        .Add "RasterInterpolationOp", "RasterInterpolationOp"
        .Add "RasterLayerPropertyPages", "RasterLayerPropertyPages"
        .Add "RasterLocalOp", "RasterLocalOp"
        .Add "RasterMapAlgebraOp", "RasterMapAlgebraOp"
        .Add "RasterMathOps", "RasterMathOps"
        .Add "RasterNeighborhoodOp", "RasterNeighborhoodOp"
        .Add "RasterOverlayOp", "RasterOverlayOp"
        .Add "RasterRasterMakerOp", "RasterRasterMakerOp"
        .Add "RasterReclassOp", "RasterReclassOp"
        .Add "RasterRendererMakers", "RasterRendererMakers"
        .Add "RasterStatEnvDialog", "RasterStatEnvDialog"
        .Add "RasterStatisticsOp", "RasterStatisticsOp"
        .Add "RasterSurfaceOp", "RasterSurfaceOp"
        .Add "RasterTransformationOp", "RasterTransformationOp"
        .Add "RasterZonalOp", "RasterZonalOp"
        .Add "RelationshipClassPropertyPages", "RelationshipClassPropertyPages"
        .Add "RendererPropertyPages", "RendererPropertyPages"
        .Add "SAExtensions", "SAExtensions"
        .Add "SAOperation", "SAOperation"
        .Add "ScaleBars", "ScaleBars"
        .Add "SceneExporter3dDrivers", "SceneExporter3dDrivers"
        .Add "SceneExporter3dPropertyPages", "SceneExporter3dPropertyPages"
        .Add "SceneExporterVideoDrivers", "SceneExporterVideoDrivers"
        .Add "SceneExporterVideoPropertyPages", "SceneExporterVideoPropertyPages"
        .Add "ScenePropertyPages", "ScenePropertyPages"
        .Add "SearchResultsContextMenuCommands", "SearchResultsContextMenuCommands"
        .Add "Shadows", "Shadows"
        .Add "ShapefileFeatureClassPropertyPages", "ShapefileFeatureClassPropertyPages"
        .Add "ShapefileTablePropertyPages", "ShapefileTablePropertyPages"
        .Add "SketchMenuCommands", "SketchMenuCommands"
        .Add "SketchToolMenuCommands", "SketchToolMenuCommands"
        .Add "SnapAgents", "SnapAgents"
        .Add "StandaloneTablePropertyPages", "StandaloneTablePropertyPages"
        .Add "StyleGalleryClasses", "StyleGalleryClasses"
        .Add "StyleImporter", "StyleImporter"
        .Add "StylePropertyPages", "StylePropertyPages"
        .Add "SxCommandBars", "SxCommandBars"
        .Add "SxCommands", "SxCommands"
        .Add "SxContentsViews", "SxContentsViews"
        .Add "SxContextAnalyzers", "SxContextAnalyzers"
        .Add "SxDockableWindows", "SxDockableWindows"
        .Add "SxExtensions", "SxExtensions"
        .Add "SxFinders", "SxFinders"
        .Add "SxLayerPropertyPages", "SxLayerPropertyPages"
        .Add "SxOptionsPropertyPages", "SxOptionsPropertyPages"
        .Add "SxPalettes", "SxPalettes"
        .Add "SxPropertyPages", "SxPropertyPages"
        .Add "SxRootObjects", "SxRootObjects"
        .Add "SymbolPropertyPages", "SymbolPropertyPages"
        .Add "TablePropertyPages", "TablePropertyPages"
        .Add "TableWindowPropertyPages", "TableWindowPropertyPages"
        .Add "TextBackground", "TextBackground"
        .Add "TextBackgroundPropertyPages", "TextBackgroundPropertyPages"
        .Add "TextFilePropertyPages", "TextFilePropertyPages"
        .Add "TextSymbol", "TextSymbol"
        .Add "TinDatasetPropertyPages", "TinDatasetPropertyPages"
        .Add "TinRendererPropertyPages", "TinRendererPropertyPages"
        .Add "TraceTasks", "TraceTasks"
        .Add "TransformationMethods", "TransformationMethods"
        .Add "UtilityNetworkAnalysisGNPropertyPages", "UtilityNetworkAnalysisGNPropertyPages"
        .Add "UtilityNetworkAnalysisPropertyPages", "UtilityNetworkAnalysisPropertyPages"
        .Add "UtilityNetworkTasks", "UtilityNetworkTasks"
        .Add "VersionPropertyPages", "VersionPropertyPages"
        .Add "WorkspaceDomainsPropertyPages", "WorkspaceDomainsPropertyPages"
        .Add "WorkspaceFactory", "WorkspaceFactory"
        .Add "WorkspacePropertyPages", "WorkspacePropertyPages"
    End With
End Sub

'***************************************************************************
'Name:                  CatalogComponentGuids
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       4/11/2007
'Purpose:       Initialize collection of component GUIDs
'Called From:   Class_Initialize
'Description:   None
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
'John Walton    4/11/2007    Initial creation
'***************************************************************************

Private Sub CatalogComponentGuids()
    With m_colComponentGuids
        .Add "{0772EAEC-F5B4-4F2C-BFE8-019861BD3A43}", "AdjustmentMethodPropertyPages"
        .Add "{3E15409F-572B-47DE-91C3-7F310CE04F5B}", "AdjustmentPropertyPages"
        .Add "{E8778391-8124-11D5-B29B-00508BCDDE28}", "AnimationPropertyPages"
        .Add "{1532016F-7C3E-46EA-8753-F3A0E05EA015}", "AnnotateMapPropertyPages"
        .Add "{6ED8A829-5036-11D3-9FB3-00C04F6BC6A5}", "AnnotationEngines"
        .Add "{F3E803A1-6EC5-11D2-A2C6-080009B6F22B}", "AnnotationExpresionPages"
        .Add "{02D588E1-3D42-11D2-A2A2-080009B6F22B}", "AnnotationExpressionParsers"
        .Add "{2B5310E1-6EC9-11D2-A2C6-080009B6F22B}", "AnnotationPlacementPages"
        .Add "{6ED8B756-1BAA-12D3-9FB3-00C04F6BB7A6}", "AnnotationPropertiesConverters"
        .Add "{B5E9A20C-0D77-11D4-9FE0-00C04F6BC61D}", "ArcToolboxTools"
        .Add "{F21E3E7B-E3BA-11D2-B868-00600802E603}", "AreaPatches"
        .Add "{51C6C6AB-DACE-49DB-AF63-32F241C35E49}", "AttributeTransferTypes"
        .Add "{56BBC8A4-E583-11D3-9FD8-00C04F6BC78E}", "AVFSrcImporters"
        .Add "{56BBC8A1-E583-11D3-9FD8-00C04F6BC78E}", "AVThemeImporters"
        .Add "{EB5A6D3A-DE07-11D2-B868-00600802E603}", "Backgrounds"
        .Add "{EB5A6D39-DE07-11D2-B868-00600802E603}", "Borders"
        .Add "{83EEE9C7-0CB1-11D3-9B35-00C04FA33299}", "CadDrawingPropertyPages"
        .Add "{83EEE9C6-0CB1-11D3-9B35-00C04FA33299}", "CadFeatureClassPropertyPages"
        .Add "{83EEE9C5-0CB1-11D3-9B35-00C04FA33299}", "CadFeaturePropertyPages"
        .Add "{F5B68AD5-EB4E-11D1-AAB9-00C04FA334B3}", "ClassificationMethods"
        .Add "{52BB5364-947E-11D2-ACFF-0000F87808EE}", "ClipboardFormats"
        .Add "{3745A6D7-50EF-11D3-9FF4-00C04F6BC8DD}", "ColorCorrectionPages"
        .Add "{31CB4BD9-2FDC-11D3-9F58-00C04F6BC709}", "ColorPropertyPages"
        .Add "{2411EC71-8BCC-11D2-A2EE-080009B6F22B}", "ColorRamp"
        .Add "{2411EC72-8BCC-11D2-A2EE-080009B6F22B}", "ColorRampPropetyPages"
        .Add "{089874FC-CC18-11D2-9F39-00C04F6BC78E}", "ContentsViews"
        .Add "{0C452011-84CB-11D2-AE68-080009EC732A}", "ContextAnalyzers"
        .Add "{5F7BBE84-E6EC-11D2-9F30-00C04F79927C}", "CoverageFeatureClassPropertyPages"
        .Add "{5F7BBE83-E6EC-11D2-9F30-00C04F79927C}", "CoveragePropertyPages"
        .Add "{75C2A63E-A98A-11D2-AB27-00C04FA334B3}", "DataExclusionPropertyPages"
        .Add "{977D75CF-E00C-11D3-A664-0008C7DF88DB}", "DataGraphPropertyPages"
        .Add "{5C15FD91-2DB2-11D3-9F52-00C04F6BC709}", "DataSamplingPropertyPages"
        .Add "{B56A7C49-83D4-11D2-A2E9-080009B6F22B}", "DataWindowFactory"
        .Add "{5980E69C-A95E-11D5-B2A0-00508BCDDE28}", "DDDAnimationTypes"
        .Add "{A1523ADD-FB91-11D4-80FC-00C04FA0ADF8}", "DDDExtensions"
        .Add "{0864F69B-E975-11D3-9F79-00C04F6BDDD9}", "DimensionConstructors"
        .Add "{58122E51-D89C-11D2-9F25-00C04F6BC979}", "EditorExtensions"
        .Add "{58122E52-D89C-11D2-9F25-00C04F6BC979}", "EditorPropertyPages"
        .Add "{58122E50-D89C-11D2-9F25-00C04F6BC979}", "EditTasks"
        .Add "{43B21395-F1E2-11D2-B870-00600802E603}", "ElementPropertyPages"
        .Add "{DA805B10-54D7-11D3-8D39-00C04F5B87B2}", "ESDAEngineMethods"
        .Add "{5FDE0991-7B12-11D2-ACF9-0000F87808EE}", "ExporterDrivers"
        .Add "{5FDE099A-7B12-11D2-ACF9-0000F87808EE}", "ExporterPropertyPages"
        .Add "{750D818E-E0B5-11D2-99BF-0000F80372B4}", "FeatureClassPropertyPages"
        .Add "{6EDC31E1-E3AD-11D2-99C1-0000F80372B4}", "FeatureDatasetPropertyPages"
        .Add "{7480BE41-EEA1-4E16-858C-B2AB6C92F720}", "FeatureLayerSymbology"
        .Add "{5424FB02-5713-11D2-97E5-0080C7E04196}", "FillSymbol"
        .Add "{FB5216EF-6AC6-11D3-9F4E-00C04F6BC8DD}", "FontMappingPropertyPages"
        .Add "{43B21396-F1E2-11D2-B870-00600802E603}", "FrameElementPropertyPages"
        .Add "{3297E9C1-93A1-11D2-9F33-00C04F8ED1D7}", "GeoAnalysis"
        .Add "{DA10C1C7-09A3-11D4-9FAA-00C04F6BDF0C}", "GeodatabaseWorkspaceExtensions"
        .Add "{B56A7C44-83D4-11D2-A2E9-080009B6F22B}", "GeoElementTools"
        .Add "{7E9913E1-E22E-11D2-99C0-0000F80372B4}", "GeographicCoordSysPropertyPages"
        .Add "{750D818F-E0B5-11D2-99BF-0000F80372B4}", "GeometricNetworkPropertyPages"
        .Add "{0813623A-A72E-11D2-8924-0000F877762D}", "GeoObjectClassDescriptions"
        .Add "{D4E2A322-5D59-11D2-89FD-006097AFF44E}", "GeoObjectClassExtensions"
        .Add "{D4E2A321-5D59-11D2-89FD-006097AFF44E}", "GeoObjects"
        .Add "{9FB18846-5930-4D7A-AFBE-0FF2E4B6DA8C}", "GeoRefSaveAsDlg"
        .Add "{ACCCD296-9D8E-4F6B-97F4-39D65330D891}", "GeostatExtensions"
        .Add "{583A720C-FC86-46A1-A6DD-E5C315076BEB}", "GNEditorPropertyPages"
        .Add "{D618307B-AA29-11D2-87F2-0000F8751720}", "GraphicSnapAgents"
        .Add "{8AF09A6C-A0DD-11D2-AE7F-080009EC732A}", "GridLabels"
        .Add "{23B9BCD6-6F6F-11D2-AAC3-00C04FA375FB}", "GSEngineMethods"
        .Add "{60B20514-C449-11D3-A63F-0008C7BF3347}", "GSRenderers"
        .Add "{9EB1D578-740B-423E-8CF4-C77A1775A5D6}", "GxArcViewExportFeatureClassFilters"
        .Add "{593C6231-9B13-4229-A1DF-A7A36F1AE10D}", "GxArcViewExportTableFilters"
        .Add "{83EEE9CA-0CB1-11D3-9B35-00C04FA33299}", "GxCadDataContextMenuCommands"
        .Add "{83EEE9C1-0CB1-11D3-9B35-00C04FA33299}", "GxCadDrawingContextMenuCommands"
        .Add "{83EEE9C4-0CB1-11D3-9B35-00C04FA33299}", "GxCadDrawingExportMenuCommands"
        .Add "{83EEE9C0-0CB1-11D3-9B35-00C04FA33299}", "GxCadFeatureClassContextMenuCommands"
        .Add "{83EEE9C3-0CB1-11D3-9B35-00C04FA33299}", "GxCadFeatureClassExportMenuCommands"
        .Add "{83EEE9BF-0CB1-11D3-9B35-00C04FA33299}", "GxCadFeatureContextMenuCommands"
        .Add "{83EEE9C2-0CB1-11D3-9B35-00C04FA33299}", "GxCadFeatureExportMenuCommands"
        .Add "{25C0E6F0-CD06-11D2-9F40-00C04F6BC626}", "GxCatalogContextMenuCommands"
        .Add "{BD1DF56D-E644-11D2-99C1-0000F80372B4}", "GxCatalogPropertyPages"
        .Add "{56C205F9-E53A-11D1-9496-080009EEBECB}", "GxCommandBars"
        .Add "{5F08CBCA-E91F-11D1-AEE8-080009EC734B}", "GxCommands"
        .Add "{25B9A22F-E6C7-11D2-9F30-00C04F79927C}", "GxCoverageContextMenuCommands"
        .Add "{A6BF6E97-EB70-11D2-9F31-00C04F79927C}", "GxCoverageExportMenuCommands"
        .Add "{25B9A230-E6C7-11D2-9F30-00C04F79927C}", "GxCoverageFeatureClassContextMenuCommands"
        .Add "{A6BF6E98-EB70-11D2-9F31-00C04F79927C}", "GxCoverageFeatureClassExportMenuCommands"
        .Add "{1449254B-319B-11D3-9F33-00C04F79927C}", "GxCoverageNewMenuCommands"
        .Add "{25C0E6F5-CD06-11D2-9F40-00C04F6BC626}", "GxDatabaseContextMenuCommands"
        .Add "{6CB2F1D8-4E9D-11D3-9F52-00C04F6BDF06}", "GxDatabaseExtensions"
        .Add "{7513EB97-D7EE-11D2-9F50-00C04F6BC626}", "GxDatabaseImportMenuCommands"
        .Add "{25C0E706-CD06-11D2-9F40-00C04F6BC626}", "GxDatabaseNewMenuCommands"
        .Add "{B14315A5-098A-11D4-A676-0008C7DF88DB}", "GxDataGraphContextMenuCommands"
        .Add "{25C0E6F7-CD06-11D2-9F40-00C04F6BC626}", "GxDatasetContextMenuCommands"
        .Add "{BB8F3046-1D26-11D3-9F55-00C04F6BC69E}", "GxDialogObjectFilters"
        .Add "{25C0E6F1-CD06-11D2-9F40-00C04F6BC626}", "GxDiskConnectionContextMenuCommands"
        .Add "{117623B4-F9D1-11D3-A67F-0008C7DF97B9}", "GxDockableWindows"
        .Add "{6E208C99-DBD3-11D2-9F2F-00C04F6BC69E}", "GxEnabledWorkspaceFactories"
        .Add "{F4EDFE54-E856-4D72-AEC6-7BD59A3E7B52}", "GxEnterpriseFeatureClassContextMenuCommands"
        .Add "{9648D913-4475-4165-B746-D2058776F9E8}", "GxEnterpriseFeatureClassContextMenuCommandsAV"
        .Add "{DF03366D-F30C-490B-A353-B05BAA18A3AF}", "GxEnterpriseFeatureDatasetContextMenuCommands"
        .Add "{16685968-2C9E-45A5-9D47-F2DC62942DDE}", "GxEnterpriseRasterBandContextMenuCommands"
        .Add "{F2889296-A1EC-4BDF-AEA9-D08D8E0837D9}", "GxEnterpriseRasterDatasetContextMenuCommands"
        .Add "{2004CEC9-E2BD-4D2D-8F3F-9E791E7AB106}", "GxEnterpriseRelationshipClassContextMenuCommands"
        .Add "{DABA683A-7888-40BE-8E70-49BC0D4F8C30}", "GxEnterpriseStandaloneFeatureClassContextMenuCommands"
        .Add "{39C1320B-01A5-4D61-B286-2880337B53C4}", "GxEnterpriseStandaloneFeatureClassContextMenuCommandsAV"
        .Add "{9B9ACF28-60FE-4ABA-9DC4-5B060A903DF2}", "GxEnterpriseTableContextMenuCommands"
        .Add "{0273158A-0EA1-4627-B60A-9D3DF568F8F6}", "GxEnterpriseTableContextMenuCommandsAV"
        .Add "{8E415006-F854-4ECB-98C3-EFBD68BAF635}", "GxExportFeatureClassFilters"
        .Add "{260A3656-96D0-417B-AC05-3469A444015A}", "GxExportTableFilters"
        .Add "{4531C69D-DC07-11D2-9F2F-00C04F6BC69E}", "GxExtensions"
        .Add "{58374DA9-3A52-11D3-9F70-00C04F6BC69E}", "GxExtensionToolbar"
        .Add "{25C0E6F9-CD06-11D2-9F40-00C04F6BC626}", "GxFeatureClassContextMenuCommands"
        .Add "{7513EB9B-D7EE-11D2-9F50-00C04F6BC626}", "GxFeatureClassExportMenuCommands"
        .Add "{25C0E6F8-CD06-11D2-9F40-00C04F6BC626}", "GxFeatureDatasetContextMenuCommands"
        .Add "{7513EB9A-D7EE-11D2-9F50-00C04F6BC626}", "GxFeatureDatasetExportMenuCommands"
        .Add "{7513EB99-D7EE-11D2-9F50-00C04F6BC626}", "GxFeatureDatasetImportMenuCommands"
        .Add "{25C0E707-CD06-11D2-9F40-00C04F6BC626}", "GxFeatureDatasetNewMenuCommands"
        .Add "{25C0E700-CD06-11D2-9F40-00C04F6BC626}", "GxFileContextMenuCommands"
        .Add "{FBFDAC87-FE57-4ABA-B5AA-571C78B6399D}", "GxFilterFeatureClasses"
        .Add "{3EB92610-DB62-465E-A326-8E0E71C24C20}", "GxFilterFeatureDatasets"
        .Add "{5FC6B144-22A8-487F-A208-67C34877F120}", "GxFilterTables"
        .Add "{25C0E6F4-CD06-11D2-9F40-00C04F6BC626}", "GxFolderContextMenuCommands"
        .Add "{25C0E705-CD06-11D2-9F40-00C04F6BC626}", "GxFolderNewMenuCommands"
        .Add "{1A9A50A4-F123-4CEC-93DC-8AF1CE0CDDA0}", "GxFolderNewMenuCommandsAV"
        .Add "{25C0E6FB-CD06-11D2-9F40-00C04F6BC626}", "GxGeometricNetworkContextMenuCommands"
        .Add "{25C0E6FE-CD06-11D2-9F40-00C04F6BC626}", "GxGridContextMenuCommands"
        .Add "{25B9A231-E6C7-11D2-9F30-00C04F79927C}", "GxInfoTableContextMenuCommands"
        .Add "{A6BF6E99-EB70-11D2-9F31-00C04F79927C}", "GxInfoTableExportMenuCommands"
        .Add "{506187AA-301B-11D4-9FC8-00C04F6BDF06}", "GxInfoTableToFeatureClassMenuCommands"
        .Add "{25C0E702-CD06-11D2-9F40-00C04F6BC626}", "GxLayerContextMenuCommands"
        .Add "{C2C372C2-1EB4-4101-9A38-992536A5C0D0}", "GxLayerFilters"
        .Add "{940C2D60-4AAD-11D3-9F51-00C04F6BDF06}", "GxLocatorContextMenuCommands"
        .Add "{E62D6CB4-4AB0-11D3-9F51-00C04F6BDF06}", "GxLocatorFolderContextMenuCommands"
        .Add "{55EFFA76-622C-11D3-9F60-00C04F8ED1C4}", "GxLocatorMenus"
        .Add "{25C0E703-CD06-11D2-9F40-00C04F6BC626}", "GxMapContextMenuCommands"
        .Add "{2044F8C9-D8D7-11D2-99BE-0000F80372B4}", "GxMapPropertyPages"
        .Add "{E5FB4CF2-660C-11D3-A68B-0008C7D3AE50}", "GxMetadataExporter"
        .Add "{E5FB4CF1-660C-11D3-A68B-0008C7D3AE50}", "GxMetadataImporter"
        .Add "{35DB0ACD-1DAF-11D3-9FAE-00C04F6BC626}", "GxMultipleSelectionContextMenuCommands"
        .Add "{65D77504-895C-11D2-AF5D-080009EC734B}", "GxObjectFactory"
        .Add "{BB8F3045-1D26-11D3-9F55-00C04F6BC69E}", "GxObjectFilters"
        .Add "{8F4CFA6B-B38E-4E4B-9AB1-E1AFB3CBB569}", "GxOleDBDatabaseContextMenuCommands"
        .Add "{44C34C54-E41B-40EF-AC63-8D25A9FC01FA}", "GxOleDBTableContextMenuCommands"
        .Add "{5F08CBCB-E91F-11D1-AEE8-080009EC734B}", "GxPalettes"
        .Add "{4B6C687E-66D2-11D3-9F36-00C04F79927C}", "GxPCCoverageContextMenuCommands"
        .Add "{4B6C687F-66D2-11D3-9F36-00C04F79927C}", "GxPCCoverageFeatureClassContextMenuCommands"
        .Add "{F7F3ABF6-FB3C-4E7F-9F38-D45DCDF5F30C}", "GxPMFMapContextMenuCommands"
        .Add "{C00B1F2C-52D1-404F-8DD5-E3694FD60E00}", "GxPMFMapPropertyPages"
        .Add "{C65A2BB4-32ED-11D3-9F33-00C04F79927C}", "GxPre70CoverageContextMenuCommands"
        .Add "{3EEA2CB0-A730-11D2-AF6D-080009EC734B}", "GxPreviews"
        .Add "{25C0E701-CD06-11D2-9F40-00C04F6BC626}", "GxPrjFileContextMenuCommands"
        .Add "{25C0E6FD-CD06-11D2-9F40-00C04F6BC626}", "GxRasterBandContextMenuCommands"
        .Add "{25C0E6FC-CD06-11D2-9F40-00C04F6BC626}", "GxRasterDatasetContextMenuCommands"
        .Add "{A44C92BC-79D6-4F89-A0A8-40CDCC279EDC}", "GxReadOnlyFeatureClassContextMenuCommands"
        .Add "{5484E3DE-86FB-489C-8EB1-73EC4A0E4485}", "GxReadOnlyFeatureDatasetContextMenuCommands"
        .Add "{D1E6735C-4464-4875-8673-59D94A603C3C}", "GxReadOnlyStandaloneFeatureClassContextMenuCommands"
        .Add "{3BF096DF-B346-46AF-99E2-AFB52ED90F4E}", "GxReadOnlyTableContextMenuCommands"
        .Add "{6C9CB46D-DCA2-11D2-9F54-00C04F6BC626}", "GxRelationshipClassContextMenuCommands"
        .Add "{25C0E6F6-CD06-11D2-9F40-00C04F6BC626}", "GxRemoteDatabaseContextMenuCommands"
        .Add "{25C0E6F2-CD06-11D2-9F40-00C04F6BC626}", "GxRemoteDatabaseFolderContextMenuCommands"
        .Add "{6E208C9A-DBD3-11D2-9F2F-00C04F6BC69E}", "GxRootObjects"
        .Add "{3E804E72-8FE3-11D4-AB88-0008C73FD50C}", "GxSceneContextMenuCommands"
        .Add "{A72131C2-90FA-11D4-AB88-0008C73FD50C}", "GxScenePropertyPages"
        .Add "{95BCD5A1-8C27-4ACE-9675-BFF5391CE339}", "GxSDCFeatureClassExportMenuCommands"
        .Add "{291B2148-0495-43AC-9108-F9891917A0D1}", "GxSDCTableExportMenuCommands"
        .Add "{1236649F-CD2B-11D3-A6F0-0008C7D3AE50}", "GxSearchEngines"
        .Add "{E10D6EC0-06E6-11D3-9F87-00C04F6BC626}", "GxShapefileFeatureClassContextMenuCommands"
        .Add "{E10D6EC2-06E6-11D3-9F87-00C04F6BC626}", "GxShapefileFeatureClassExportMenuCommands"
        .Add "{E10D6EC1-06E6-11D3-9F87-00C04F6BC626}", "GxShapefileTableContextMenuCommands"
        .Add "{E10D6EC3-06E6-11D3-9F87-00C04F6BC626}", "GxShapefileTableExportMenuCommands"
        .Add "{124C8C26-301B-11D4-9FC8-00C04F6BDF06}", "GxShapefileTableToFeatureClassMenuCommands"
        .Add "{AB6C1071-13FC-4274-8287-A8AA35A1011A}", "GxShortcutContextMenuCommands"
        .Add "{25C0E6F3-CD06-11D2-9F40-00C04F6BC626}", "GxSpatialReferencesFolderContextMenuCommands"
        .Add "{9828BC65-3AD7-11D3-A62F-0008C7D3AE8D}", "GxSpatialReferencesFolderNewMenuCommands"
        .Add "{22211295-7778-11D3-A00D-00C04F6BC626}", "GxStandaloneFeatureClassContextMenuCommands"
        .Add "{3B5C1D75-C440-4EBF-B629-ACCAD41BF396}", "GxStreetMapFeatureClassContextMenuCommands"
        .Add "{5DBC42A8-3F19-4138-BDFE-C8EB60456A0F}", "GxStreetMapFeatureClassExportMenuCommands"
        .Add "{F1B979F7-AE98-478A-8A80-C731B8C8F3E9}", "GxStreetMapFeatureClassPropertyPages"
        .Add "{25C0E6FA-CD06-11D2-9F40-00C04F6BC626}", "GxTableContextMenuCommands"
        .Add "{16FD2B1D-1556-11D3-9FA7-00C04F6BC626}", "GxTableExportMenuCommands"
        .Add "{20580FBF-2DDE-11D4-9FC8-00C04F6BDF06}", "GxTableToFeatureClassMenuCommands"
        .Add "{3EEA2CAF-A730-11D2-AF6D-080009EC734B}", "GxTabViews"
        .Add "{25C0E6FF-CD06-11D2-9F40-00C04F6BC626}", "GxTinContextMenuCommands"
        .Add "{3EEA2CB1-A730-11D2-AF6D-080009EC734B}", "GxViews"
        .Add "{832A7CA1-F47F-11D3-9B56-00C04FA33299}", "GxVpfCoverageContextMenuCommands"
        .Add "{832A7CA3-F47F-11D3-9B56-00C04FA33299}", "GxVpfCoverageExportMenuCommands"
        .Add "{832A7CA5-F47F-11D3-9B56-00C04FA33299}", "GxVpfCoveragePropertyPages"
        .Add "{832A7CA2-F47F-11D3-9B56-00C04FA33299}", "GxVpfFeatureClassContextMenuCommands"
        .Add "{832A7CA4-F47F-11D3-9B56-00C04FA33299}", "GxVpfFeatureClassExportMenuCommands"
        .Add "{832A7CA6-F47F-11D3-9B56-00C04FA33299}", "GxVpfFeatureClassPropertyPages"
        .Add "{1CB448E4-F485-11D3-9B56-00C04FA33299}", "GxVpfTableContextMenuCommands"
        .Add "{1CB448E5-F485-11D3-9B56-00C04FA33299}", "GxVpfTableExportMenuCommands"
        .Add "{1CB448E6-F485-11D3-9B56-00C04FA33299}", "GxVpfTablePropertyPages"
        .Add "{997E7AC0-D353-11D2-8CD0-00C04F5B951E}", "IdentifyObj"
        .Add "{FEE7FBA1-C40E-11D5-B2A2-00508BCDDE28}", "ImportPathPropertyPages"
        .Add "{079E9B6C-DA59-44EB-B129-E013139D3812}", "IMSMetadataServicePropertyPages"
        .Add "{5F7BBE85-E6EC-11D2-9F30-00C04F79927C}", "InfoTablePropertyPages"
        .Add "{4C90DE7D-CB77-11D2-9F34-00C04F6BC6A5}", "LabelStylePages"
        .Add "{34C20001-4D3C-11D0-92D8-00805F7C28B0}", "LayerFactory"
        .Add "{D80B7114-739C-4ECB-AD06-4DF67C0E6596}", "LayerLabelsPropertyPages"
        .Add "{1476C782-6F57-11D2-A2C6-080009B6F22B}", "LayerPropertyPages"
        .Add "{DC70783B-5CE1-11D0-92DA-00805F7C28B0}", "LayerPropertySheet"
        .Add "{4E91E870-4C86-11D1-880B-080009EC732A}", "LegendItemPropSheets"
        .Add "{4E91E86F-4C86-11D1-880B-080009EC732A}", "LegendItems"
        .Add "{9C275BE0-2A66-11D3-80E3-0080C79F0371}", "LegendSymbolEditors"
        .Add "{5673B810-2B1F-11D3-80E5-0080C79F0371}", "LineDecorationElement"
        .Add "{F21E3E7C-E3BA-11D2-B868-00600802E603}", "LinePatches"
        .Add "{5424FB00-5713-11D2-97E5-0080C7E04196}", "LineSymbol"
        .Add "{8AF09A6E-A0DD-11D2-AE7F-080009EC732A}", "MapGridBorders"
        .Add "{A0C55AF7-F9D5-11D1-AA8E-000000000000}", "MapGridFactory"
        .Add "{452A7917-1B8E-11D3-9F92-00C04F6BC78E}", "MapGridPropertyPages"
        .Add "{AE76C8D2-A58A-11D2-A2FE-080009B6F22B}", "MapPropertyPages"
        .Add "{5424FB01-5713-11D2-97E5-0080C7E04196}", "MarkerSymbol"
        .Add "{A7386B03-775E-11D3-9F54-00C04F8ED1D7}", "MBDerivedGeoDatasetProp"
        .Add "{A7386B01-775E-11D3-9F54-00C04F8ED1D7}", "MBGeoDatasetProp"
        .Add "{A7386B05-775E-11D3-9F54-00C04F8ED1D7}", "MBGeoOperationProp"
        .Add "{A7386B02-775E-11D3-9F54-00C04F8ED1D7}", "MBLayerProp"
        .Add "{A7386B06-775E-11D3-9F54-00C04F8ED1D7}", "MBMethodProp"
        .Add "{A7386B07-775E-11D3-9F54-00C04F8ED1D7}", "MBProcessProp"
        .Add "{A7386B04-775E-11D3-9F54-00C04F8ED1D7}", "MBTableProp"
        .Add "{7AD0DA0A-0D4A-11D3-A626-0008C7D3AE50}", "MetadataEditor"
        .Add "{5907D721-04C9-11D3-83FF-00C04F8ED211}", "MetadataEditorPropertyPages"
        .Add "{A802F1A4-E795-4169-997C-6B8FE2F53F98}", "MetadataSynchronizers"
        .Add "{B56A7C4A-83D4-11D2-A2E9-080009B6F22B}", "MxCommandBars"
        .Add "{B56A7C42-83D4-11D2-A2E9-080009B6F22B}", "MxCommands"
        .Add "{F440A933-565C-447C-8C56-28BB0404E5B9}", "MxDDECommandHandler"
        .Add "{117623B5-F9D1-11D3-A67F-0008C7DF97B9}", "MxDockableWindows"
        .Add "{9168AC07-5DBE-11D4-A689-0008C7DF88DB}", "MxDocumentDropTarget"
        .Add "{B56A7C45-83D4-11D2-A2E9-080009B6F22B}", "MxExtension"
        .Add "{F8FC396D-5BA0-44A8-98C4-96C70C5312DE}", "MxFileMenuCommands"
        .Add "{6F400F1B-1B8F-4469-9581-A0A3826B9FD5}", "MxFileMenuDefs"
        .Add "{8803C8EF-75FE-11D3-A6A6-0008C7D3AE50}", "MxFinders"
        .Add "{67713804-1368-4C15-A064-639396316EC8}", "MxInsetWindowPages"
        .Add "{B56A7C48-83D4-11D2-A2E9-080009B6F22B}", "MxPalettes"
        .Add "{6B1DF87B-DEC8-49B0-884F-345FE2EB1274}", "MxToolMenuCommands"
        .Add "{B56A7C43-83D4-11D2-A2E9-080009B6F22B}", "MxViewCommands"
        .Add "{7BF26B1F-C776-11D2-9F24-00C04F6BC69E}", "NameStringParser"
        .Add "{D80D64ED-0B0F-4C45-B68C-F71F1C310353}", "NativeTypes"
        .Add "{88002C0B-939A-11D2-AE73-080009EC732A}", "NumberFormatPropertyPages"
        .Add "{FB8C2199-9D0A-11D3-A645-0008C7DF8DE1}", "OptionsPropertyPages"
        .Add "{289C9956-C063-11D2-9F22-00C04F6BC8DD}", "Pictures"
        .Add "{6871A79C-CD65-46F5-A7BB-88F97D0E2CDC}", "PlugInWorkspaceFactoryHelpers"
        .Add "{C65A2BB7-32ED-11D3-9F33-00C04F79927C}", "Pre70CoveragePropertyPages"
        .Add "{DF933391-7A7A-11D2-ACF9-0000F87808EE}", "PrinterDrivers"
        .Add "{9A797CC1-5E01-11D2-9181-0000F87808EE}", "PrinterPropertyPages"
        .Add "{7E9913E2-E22E-11D2-99C0-0000F80372B4}", "ProjectedCoordSysPropertyPages"
        .Add "{C899AE27-E515-11D1-877C-0000F8751720}", "PropertyPages"
        .Add "{59CA6922-47FB-11D2-8D29-0000F8780535}", "RasterAnalysis"
        .Add "{DC942D39-E3A8-11D2-9F52-00C04F8ED21A}", "RasterBandPropertyPages"
        .Add "{09FF656C-F398-11D3-9F9B-00C04F8EF00E}", "RasterConditionalOp"
        .Add "{CD568B5A-CA31-11D2-9F3C-00C04F8ED1D7}", "RasterConversionOp"
        .Add "{9A90E0A4-6265-11D3-9FD1-00C04F68E699}", "RasterDataExclusionPropSheet"
        .Add "{74F98CC2-E2E7-11D2-9F51-00C04F8ED21A}", "RasterDatasetPropertyPages"
        .Add "{EB3F5D32-1306-11D3-9F45-00C04F8ED1D7}", "RasterDensityOp"
        .Add "{4606F178-B47D-11D2-9F3B-00C04F8ED1D7}", "RasterDistanceOp"
        .Add "{837FAC81-38C9-11D4-8D94-00C04F5B87B2}", "RasterExportMenuCommands"
        .Add "{09FF656D-F398-11D3-9F9B-00C04F8EF00E}", "RasterExtractionOp"
        .Add "{EB3F5D31-1306-11D3-9F45-00C04F8ED1D7}", "RasterGeneralizeOp"
        .Add "{8837B6E0-7F9E-11D4-B284-00508BCDC7C8}", "RasterGxExportFilters"
        .Add "{EB3F5D2C-1306-11D3-9F45-00C04F8ED1D7}", "RasterHydrologyOp"
        .Add "{3297E9D4-93A1-11D2-9F33-00C04F8ED1D7}", "RasterInterpolationOp"
        .Add "{2AD25BF1-3C02-11D3-A3FA-0004AC1B1D86}", "RasterLayerPropertyPages"
        .Add "{83370E8B-9F6B-11D2-9F35-00C04F8ED1D7}", "RasterLocalOp"
        .Add "{3297E9C3-93A1-11D2-9F33-00C04F8ED1D7}", "RasterMapAlgebraOp"
        .Add "{4606F177-B47D-11D2-9F3B-00C04F8ED1D7}", "RasterMathOps"
        .Add "{4606F17A-B47D-11D2-9F3B-00C04F8ED1D7}", "RasterNeighborhoodOp"
        .Add "{EB3F5D2E-1306-11D3-9F45-00C04F8ED1D7}", "RasterOverlayOp"
        .Add "{CC0B70D8-EA20-11D3-9F63-00C04F8ED1D7}", "RasterRasterMakerOp"
        .Add "{EB3F5D2D-1306-11D3-9F45-00C04F8ED1D7}", "RasterReclassOp"
        .Add "{F38F6DB0-C73E-11D4-B290-00508BCDC7C8}", "RasterRendererMakers"
        .Add "{80F65394-FD12-11D2-9F63-00C04F68E699}", "RasterStatEnvDialog"
        .Add "{EB3F5D2F-1306-11D3-9F45-00C04F8ED1D7}", "RasterStatisticsOp"
        .Add "{4606F179-B47D-11D2-9F3B-00C04F8ED1D7}", "RasterSurfaceOp"
        .Add "{EB3F5D30-1306-11D3-9F45-00C04F8ED1D7}", "RasterTransformationOp"
        .Add "{4606F17B-B47D-11D2-9F3B-00C04F8ED1D7}", "RasterZonalOp"
        .Add "{05318B65-DE31-11D2-9F54-00C04F6BC626}", "RelationshipClassPropertyPages"
        .Add "{962BD9A9-1EC8-11D3-9F4D-00C04F6BC709}", "RendererPropertyPages"
        .Add "{BA648A2A-B79E-433C-999A-AA3C7B075E17}", "SAExtensions"
        .Add "{FAF0CFBB-CD00-11D3-A051-00C04F68E699}", "SAOperation"
        .Add "{34DC96D1-186D-11D3-B8A3-00600802E603}", "ScaleBars"
        .Add "{D6AC48F8-37F4-11D4-A168-444553547777}", "SceneExporter3dDrivers"
        .Add "{D6AC48FC-37F4-11D4-A168-444553547777}", "SceneExporter3dPropertyPages"
        .Add "{4251BA09-9106-11D5-B29F-00508BCDDE28}", "SceneExporterVideoDrivers"
        .Add "{4251BA0A-9106-11D5-B29F-00508BCDDE28}", "SceneExporterVideoPropertyPages"
        .Add "{CF228351-10A9-11D4-9F7F-00C04F6BC619}", "ScenePropertyPages"
        .Add "{13A17004-CA63-46D3-B8CF-031F4C186B3C}", "SearchResultsContextMenuCommands"
        .Add "{4DE36362-E872-4159-B2DB-3ECEA4F57F0E}", "Shadows"
        .Add "{E10D6EC4-06E6-11D3-9F87-00C04F6BC626}", "ShapefileFeatureClassPropertyPages"
        .Add "{E10D6EC5-06E6-11D3-9F87-00C04F6BC626}", "ShapefileTablePropertyPages"
        .Add "{3DD0D4DD-01E6-41AC-BB50-108D11C763A7}", "SketchMenuCommands"
        .Add "{3DD0D4DC-01E6-41AC-BB50-108D11C763A7}", "SketchToolMenuCommands"
        .Add "{5844F9F6-DC03-11D2-9F26-00C04F6BC979}", "SnapAgents"
        .Add "{DA7C157D-E58B-11D3-ADEA-00C04FA33A15}", "StandaloneTablePropertyPages"
        .Add "{AC0E982B-91CB-11D1-8813-080009EC732A}", "StyleGalleryClasses"
        .Add "{17049F82-8E15-11D2-983E-0080C7E04196}", "StyleImporter"
        .Add "{B6DD2405-29D6-11D3-B8AD-00600802E603}", "StylePropertyPages"
        .Add "{F27D8291-A383-11D3-8206-0080C7597E71}", "SxCommandBars"
        .Add "{F27D8292-A383-11D3-8206-0080C7597E71}", "SxCommands"
        .Add "{BB07DB51-E57B-11D3-9F74-00C04F6BC619}", "SxContentsViews"
        .Add "{3BB42A1D-05A0-11D4-9F79-00C04F6BC619}", "SxContextAnalyzers"
        .Add "{B9789374-FADC-11D3-A67F-0008C7DF97B9}", "SxDockableWindows"
        .Add "{369448D7-A6BD-11D3-ABB8-0008C73FCA1C}", "SxExtensions"
        .Add "{B3E00B2A-25D3-11D4-A383-00C04F6BC619}", "SxFinders"
        .Add "{3F82C99B-1C5F-11D4-A381-00C04F6BC619}", "SxLayerPropertyPages"
        .Add "{F58B83FF-31AE-11D4-A385-00C04F6BC619}", "SxOptionsPropertyPages"
        .Add "{A51B87D1-A383-11D3-8206-0080C7597E71}", "SxPalettes"
        .Add "{779E31F8-25E1-11D4-9B80-00C04FA33299}", "SxPropertyPages"
        .Add "{5A10D191-A385-11D3-8206-0080C7597E71}", "SxRootObjects"
        .Add "{818B37C0-F34E-11D2-BC8F-0080C7E04196}", "SymbolPropertyPages"
        .Add "{750D818D-E0B5-11D2-99BF-0000F80372B4}", "TablePropertyPages"
        .Add "{DA51F911-EA46-4E57-9B9E-2CF443D34CAB}", "TableWindowPropertyPages"
        .Add "{F8AEACE0-25D1-11D3-80DA-0080C79F0371}", "TextBackground"
        .Add "{F1CAF8B0-25C3-11D3-80DA-0080C79F0371}", "TextBackgroundPropertyPages"
        .Add "{090BE115-10AB-11D4-AE05-00C04FA33A15}", "TextFilePropertyPages"
        .Add "{5424FB03-5713-11D2-97E5-0080C7E04196}", "TextSymbol"
        .Add "{2A3354C9-24E1-11D3-9F4C-00C04F6BC619}", "TinDatasetPropertyPages"
        .Add "{19283BDF-656F-11D3-9F57-00C04F6BC619}", "TinRendererPropertyPages"
        .Add "{5BF495FA-95D1-11D2-AAEB-00C04FA37D59}", "TraceTasks"
        .Add "{D9246646-BEDB-4D2B-88B0-5B94D1BBD5A7}", "TransformationMethods"
        .Add "{69BCACD9-019C-11D3-9F7D-00C04F6BDF0E}", "UtilityNetworkAnalysisGNPropertyPages"
        .Add "{97499635-F279-11D2-9F6A-00C04F6BDF0E}", "UtilityNetworkAnalysisPropertyPages"
        .Add "{39DC1B4F-BE34-11D2-BABE-00C04FA33C20}", "UtilityNetworkTasks"
        .Add "{0224F7BC-F736-11D2-9F2E-00C04F6BC979}", "VersionPropertyPages"
        .Add "{C255D347-CB5C-11D2-9F3C-00C04F6BC626}", "WorkspaceDomainsPropertyPages"
        .Add "{07BC056E-DB8A-11D1-AA81-00C04FA33A15}", "WorkspaceFactory"
        .Add "{C255D343-CB5C-11D2-9F3C-00C04F6BC626}", "WorkspacePropertyPages"
    End With
End Sub


