VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCatalogFileDlg"
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
' File name:            clsCatalogFileDlg
'
' Initial Author:       JWalton
'
' Date Created:         2/5/2007
'
' Description: Programmatically expose the ArcCatalog file dialog as one integral
'              unit for simple access and use.
'
' Entry points:
'       <<None>>
'       <<None>>
'
' Dependencies:
'       File References
'           esriCatalog
'           esriCatalogUI
'
' Issues:
'       <<None>>
'
' Method:
'       Standard implementation of IToolbarDef
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
Private m_pDialog As esriCatalogUI.IGxDialog
Private m_colSelection As Collection
Private m_strFileFilter As String

'***************************************************************************
'Name:                  Class_Initialize
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
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
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub Class_Initialize()
    ' Initialize class objects
    Set m_pDialog = New esriCatalogUI.GxDialog
    
    ' Initialize miscellaneous settings
    m_pDialog.RememberLocation = True
    m_pDialog.AllowMultiSelect = False
End Sub

'***************************************************************************
'Name:                  Class_Terminate
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Clean up any class variable and objects
'Called From:   Nowhere
'Description:   Class termination routine.
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
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub Class_Terminate()
    ' Clean up class objects
    Set m_pDialog = Nothing
End Sub

'***************************************************************************
'Name:                  AllowMultiSelect
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Class Property
'Called From:   Nowhere
'Description:   Status of multiple selection in Open file dialog boxes
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
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Property Let AllowMultiSelect(ByVal blnAllowMultiSelect As Boolean)
    m_pDialog.AllowMultiSelect = blnAllowMultiSelect
End Property

'***************************************************************************
'Name:                  ButtonCaption
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Class Property
'Called From:   Nowhere
'Description:   Save As/Open button caption (Optional)
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
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Property Let ButtonCaption(ByVal strCaption As String)
    m_pDialog.ButtonCaption = strCaption
End Property

'***************************************************************************
'Name:                  Name
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Class Property
'Called From:   Multiple Locations
'Description:   Name of the file specified from a call to the ShowSave
'               method
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
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Property Let Name(ByVal strName As String)
    m_pDialog.Name = strName
End Property

Public Property Get Name() As String
    Name = m_pDialog.Name
End Property

'***************************************************************************
'Name:                  StartingLocation
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Class Property
'Called From:   Nowhere
'Description:   Initial file path for either open or save dialog boxes
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
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Property Let StartingLocation(ByVal varLocation As Variant)
    m_pDialog.StartingLocation = varLocation
End Property

'***************************************************************************
'Name:                  FinalLocation
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Class Property
'Called From:   Nowhere
'Description:   The file path present when the user specifies a file to open
'               or a file name to save in Open/Save As dialog boxes
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
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Property Get FinalLocation() As String
    FinalLocation = m_pDialog.FinalLocation.FullName
End Property

'***************************************************************************
'Name:                  Title
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Class Property
'Called From:   Nowhere
'Description:   Title of the Open/Save As dialog box
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
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Property Let Title(ByVal strTitle As String)
    m_pDialog.Title = strTitle
End Property

'***************************************************************************
'Name:                  SelectedObject
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Retrieve an item selected from a file dialog box
'Called From:   Multiple locations.
'Description:   Return the nth selected element from the most recent file
'               save/open dialog request
'Methods:       None
'Inputs:        iSelection (Read Only) -- Index of item to return.
'Parameters:    None
'Outputs:       None
'Returns:       Variant object representing the item at the nth position in
'               the collection of objects that were returned from the last
'               file save/open request.
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Function SelectedObject( _
  Optional ByVal iSelection As Integer = -1) As Variant
    ' Returns an item from the current selection set, or the selection set itself
    If iSelection > m_colSelection.Count Then
        SelectedObject = ""
      ElseIf iSelection = -1 Then
        Set SelectedObject = m_colSelection
      Else
        SelectedObject = m_colSelection(iSelection)
    End If
End Function

'***************************************************************************
'Name:                  SetFilter
'Initial Author:        John Walton
'Subsequent Author:     Type your name here.
'Created:       2/5/2007
'Purpose:       Simplify adding a filter to the file dialog box
'Called From:   Multiple locations
'Description:   Adds a ESRI ArcCatalog defined filter to a file dialog box
'               filter list
'Methods:       None
'Inputs:        pFilter (R/W) -- An ESRI ArcObject defined object filter
'               blnIsDefault (R) -- Indicates if the filter should be the
'                   default filter
'               blnResetAll (R) -- Indicates whether or not all of the
'                   current filters should be cleared
'Parameters:    None
'Outputs:       None
'Returns:       Boolean variable indicate success or failure to add the
'               specified filter to the file dialog box
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Function SetFilter( _
  ByRef pFilter As esriCatalog.IGxObjectFilter, _
  Optional blnIsDefault As Boolean, _
  Optional blnResetAll As Boolean = True) As Boolean
On Error GoTo Err_Handler
    ' Variable declarations
    Dim pFilters As esriCatalog.IGxObjectFilterCollection

    ' QI to get the filter collection
    Set pFilters = m_pDialog
    
    ' Reset the filters if necessary
    If blnResetAll Then pFilters.RemoveAllFilters
    
    ' Adds the passed in filter
    pFilters.AddFilter pFilter, blnIsDefault
    
    ' Returns the function's value and exits
    SetFilter = True
    Exit Function
    
Err_Handler:
    ' Returns the function's value and exits
    SetFilter = False
End Function

'***************************************************************************
'Name:                  ShowOpen
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Show the ESRI ArcCatalog file open dialog box
'Called From:   Multiple locations
'Description:   Shows the ESRI ArcCatalog file open dialog box as defined by
'               this class
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       A collection of names of objects that have been selected by
'               the user from the dialog box
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Function ShowOpen() As Collection
    ' Variable declarations
    Dim pSelection As esriCatalog.IEnumGxObject
    Dim pSelectedObj As esriCatalog.IGxObject

    ' Initialize objects
    Set pSelection = New esriCatalog.GxObjectArray
    Set m_colSelection = New Collection

    ' Shows the file open dialog box
    If Not m_pDialog.DoModalOpen(Application.hwnd, pSelection) Then
        Set ShowOpen = Nothing
        Exit Function
    End If
    
    ' Returns the selected items as a collection
    pSelection.Reset
    Set pSelectedObj = pSelection.Next
    Do While Not pSelectedObj Is Nothing
        m_colSelection.Add pSelectedObj.FullName
        Set pSelectedObj = pSelection.Next
    Loop
    
    ' Returns the collection to the function
    Set ShowOpen = m_colSelection
End Function

'***************************************************************************
'Name:                  ShowSave
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Show the ESRI ArcCatalog file save dialog box
'Called From:   Multiple locations
'Description:   Shows the ESRI ArcCatalog file save dialog box as defined by
'               this class
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       A collection holding the full path that is a concatenation
'               of the final path the specified name
'Errors:        This routine raises no known errors
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Function ShowSave() As Collection
    ' Variable declarations
    Dim pSelection As esriCatalog.IEnumGxObject
    Dim pSelectedObj As esriCatalog.IGxObject

    ' Initialize objects
    Set pSelection = New esriCatalog.GxObjectArray
    Set m_colSelection = New Collection

    ' Shows the file open dialog box
    If Not m_pDialog.DoModalSave(Application.hwnd) Then
        Set ShowSave = Nothing
        Exit Function
    End If
    
    ' Returns the selected items as a collection
    Set pSelectedObj = m_pDialog.FinalLocation
    m_colSelection.Add pSelectedObj.FullName & "\" & m_pDialog.Name
        
    ' Returns the collection to the function
    Set ShowSave = m_colSelection
End Function
