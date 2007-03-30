VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTaxlotAssignment 
   Caption         =   "Taxlot Assignment"
   ClientHeight    =   2415
   ClientLeft      =   1170
   ClientTop       =   2340
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTaxlotNum 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2190
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1980
      Width           =   800
   End
   Begin VB.ComboBox cmbTaxlotNum 
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   1545
   End
   Begin VB.Label lblAutoIncrements 
      Caption         =   "Auto Increment Options"
      Height          =   225
      Left            =   210
      TabIndex        =   9
      Top             =   540
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000A&
      Height          =   1245
      Left            =   120
      Top             =   660
      Width           =   4005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Starting From:"
      Height          =   225
      Left            =   870
      TabIndex        =   8
      Top             =   1500
      Width           =   1245
   End
   Begin MSForms.ToggleButton Increments_0 
      Height          =   375
      Left            =   330
      TabIndex        =   6
      Top             =   900
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "None"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton Increments_100 
      Height          =   375
      Left            =   3030
      TabIndex        =   5
      Top             =   900
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "100"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton Increments_10 
      Height          =   375
      Left            =   2130
      TabIndex        =   4
      Top             =   900
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "10"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton Increments_1 
      Height          =   375
      Left            =   1230
      TabIndex        =   3
      Top             =   900
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "1"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label2 
      Caption         =   "Type of Polygon:"
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   180
      Width           =   1275
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1245
      Left            =   135
      Top             =   645
      Width           =   4005
   End
End
Attribute VB_Name = "frmTaxlotAssignment"
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
'
' Keyword expansion for source code control
' Tag for this file : $Name$
' SCC Revision number: $Revision: 77 $
' Date of last change: $Date: 2007-02-15 10:24:03 -0800 (Thu, 15 Feb 2007) $
'
' File name:            frmTaxlotAssignment
'
' Initial Author:       <<Unknown>>
'
' Date Created:         <<Unknown>>
'
' Description:
'       Form used as the user interface for the tool defined in cmdTaxlotAssignment.
'       This form defines the taxlot number to be assigned to each polygon.
'
' Entry points:
'       Form Object
'       Properties
'           Increment    (R)
'               The taxlot increment to use when moving from one polygon to another
'           PolygonType  (R)
'               The type of polygon being attributes
'           CurrentValue (R/W)
'               The current taxlot value
'       Methods
'           Frame
'               The Frame object that ArcGIS uses to display the form
'           IsValidData
'               Flag that indicates the data on the form is valid
'
' Dependencies:
'       File References
'           esriArcMapUI
'           esriFramework
'
'
' Issues:
'       None known at this time (2/6/2007 JWalton)
'
' Method:
'       This form is implemented as a modeless form that can sit on top of ArcMap
'       while allowing the user continued access to ArcMap.
'       This implementation is made possible through the Frame property of the
'       form that implements ESRI's Modeless Window interface.
'
' Updates:
'       10/11/2006 -- Added this header comment section (JWM)
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)


Option Explicit
'******************************
' Private Definitions
'------------------------------
' Private Variables
'------------------------------
Private m_pMxDoc As esriArcMapUI.IMxDocument
'++ START JWalton 1/31/2007 Additional variable declarations
Private m_pFrame As esriFramework.IModelessFrame
Private m_bIndexMouseDown As Boolean
Private m_iIndexItem As Integer
Private m_lIncrement As Long
'++ END JWalton 1/31/2007
Private m_ParentHWND As Long

'------------------------------
'Private Constants and Enums
'------------------------------
' Variables used by the Error handler function - DO NOT REMOVE
Private Const c_sModuleFileName As String = "frmTaxlotAssignment.frm"

'++START JWalton 1/31/2007 Properties of the form
Public Property Get Increment() As Long
    Increment = m_lIncrement
End Property

Public Property Get PolygonType() As String
    PolygonType = cmbTaxlotNum.Text
End Property

Public Property Get CurrentValue() As Long
    CurrentValue = CLng(txtTaxlotNum.Text)
End Property

Public Property Let CurrentValue( _
  ByVal lngValue As Long)
    txtTaxlotNum.Text = lngValue
End Property
'++END JWalton 1/31/2007

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
        Set m_pFrame = New esriFramework.ModelessFrame
        m_pFrame.Create Me
        m_pFrame.Caption = Me.Caption
    End If
    Set Frame = m_pFrame
End Function

'***************************************************************************
'Name:                  IsValidData
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Indicate a valid dataset
'Called From:   Multiple Locations
'Description:   Indicates a valid dataset by examining the user entries and
'               determining whether or not they constitute a valid query
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       A boolean value the indicates whether or not a valid query
'               exists from the data on the form
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Function IsValidData() As Boolean
On Error Resume Next
    ' Test the taxlot type field
    If Len(cmbTaxlotNum.Text) = 0 Then
        IsValidData = False
        Exit Function
    End If
    
    ' Test the taxlot number field
    If StrComp(Me.cmbTaxlotNum.Text, "NUMBER", vbTextCompare) = 0 Then
        If Not IsNumeric(Me.txtTaxlotNum.Text) Then
            IsValidData = False
            Exit Function
        End If
    End If
  
    ' Tests for an auto increment selection
    If Increments_0.Value = False And _
       Increments_1.Value = False And _
       Increments_10.Value = False And _
       Increments_100.Value = False Then
        IsValidData = False
        Exit Function
    End If
  
    ' Returns the function's value
    IsValidData = True
End Function
'++END JWalton 1/31/2007

'***************************************************************************
'Name:                  cmbTaxlotNum_Click
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:
'Called From:
'Description:
'Methods:
'
'Inputs:
'Parameters:
'Outputs:
'Returns:
'Errors:        This routine raises no known errors.
'Assumptions:
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007        Initial creation
'***************************************************************************

Private Sub cmbTaxlotNum_Click()
On Error Resume Next
    'If one of the generic taxlot numbers is chosen, disable the taxlot textbox
'++ JWalton 1/31/2007 Revamped using enabling/disabling in contrast to locking/unlocking of controls
    '++  JWM 10/11/2006 using strcomp function
    If StrComp(Me.cmbTaxlotNum.Text, "NUMBER", vbTextCompare) <> 0 Then
        With txtTaxlotNum
            .BackColor = &HC0C0C0
            .Enabled = False
        End With
        Increments_0.Enabled = False
        Increments_1.Enabled = False
        Increments_10.Enabled = False
        Increments_100.Enabled = False
    Else
        With txtTaxlotNum
            .BackColor = ColorConstants.vbWhite
            .Enabled = True
        End With
        Increments_0.Enabled = True
        Increments_1.Enabled = True
        Increments_10.Enabled = True
        Increments_100.Enabled = True
    End If
'++ END JWalton 1/31/2007
End Sub

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
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
    sFilePath = app.Path & "\" & "Assignment_help.rtf"
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
'John Walton    2/5/2007        Initial creation
'***************************************************************************

Private Sub Form_Initialize()
    ' Variable declarations
    Dim pWindowPos As esriFramework.IWindowPosition
    
    ' Initialize objects
    Set pWindowPos = Me.Frame
    
    ' Get the initial position
    pWindowPos.Top = GetSetting("ArcGIS.ArcMap.ORMAP.Tools", "FrmTaxlotAssignment", "Top", 0)
    pWindowPos.Left = GetSetting("ArcGIS.ArcMap.ORMAP.Tools", "FrmTaxlotAssignment", "Left", 0)
    
    ' Makes the form the topmost form
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

'***************************************************************************
'Name:                  Form_Load
'Initial Author:        John Walton
'Subsequent Author:     <<Type your name here>>
'Created:               1/31/2007
'Purpose:       Initialize the form and all controls on the form, and
'               register the form
'Called From:   Form Object
'Description:   Event handler for Form Load event
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       m_pMxDoc
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007        Initial creation
'***************************************************************************

Private Sub Form_Load()
On Error GoTo ErrorHandler
    'Populate drop down combobox and set default settings
    '++ START JWalton 2/6/2007 Removed m_pApp in favor of g_pApp
    Set m_pMxDoc = g_pApp.Document
    
    'Populate with preset values
    cmbTaxlotNum.AddItem "NUMBER"
    cmbTaxlotNum.AddItem "0ROAD"
    cmbTaxlotNum.AddItem "WATER"
    cmbTaxlotNum.AddItem "0RLRD"
    cmbTaxlotNum.AddItem "00GAP"
    cmbTaxlotNum.AddItem "00LAP"
    
    '++ START JWalton 1/31/2007
    ' Control defaults
    cmbTaxlotNum.Text = "NUMBER"
    Increments_0.Value = True
    
    ' Sets the form status to open
    g_pForms.SetFormStatus Me.Name, True
    '++ END JWalton 1/31/2007

  Exit Sub
ErrorHandler:
  HandleError True, _
              "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), _
              Err.Number, _
              Err.Source, _
              Err.Description, _
              4, _
              m_ParentHWND
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

Private Sub Form_Unload( _
  Cancel As Integer)
    ' Variable declarations
    Dim pWindowPos As esriFramework.IWindowPosition
    
    ' Initialize objects
    Set pWindowPos = Me.Frame
    
    ' Saves the position of the form
    SaveSetting "ArcGIS.ArcMap.ORMAP.Tools", "FrmTaxlotAssignment", "Top", pWindowPos.Top
    SaveSetting "ArcGIS.ArcMap.ORMAP.Tools", "FrmTaxlotAssignment", "Left", pWindowPos.Left
End Sub

'***************************************************************************
'Name:                  Form_QueryUnload
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Unregister the form and save the form's position
'Called From:   Form Object
'Description:   Event handler for Form QueryUnload event
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Sets the form status to not open
    g_pForms.SetFormStatus Me.Name, False
End Sub



'***************************************************************************
'Name:                  Increments_(0,1,10,100)_Click
'Initial Author:        John Walton
'Subsequent Author:     <<Type your name here>>
'Created:               2/5/2007
'Purpose:       Manipulate the controls in Increments controls so that they'
'               appear as an option group
'Called From:   Increments_(0,1,10,100)
'Description:   Disables all controls except for the currently chosen
'               control in the control group.
'               If no control is selected, then the zero control will be
'               selected.
'Methods:       None
'Inputs:        Index, Button, Shift, X, Y,
'Parameters:    None
'Outputs:       m_lIncrement, m_iIndexItem, m_bIndexMouseDown
'Returns:       Nothing
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Private Sub Increments_0_Change()
    If Increments_0.Value Then
        ' Disables all other increments
        Increments_1.Value = False
        Increments_10.Value = False
        Increments_100.Value = False
        
        ' Sets the value of this control
        Increments_0.Value = True
        
        ' Saves the value to a property
        m_lIncrement = 0
      Else
        ' Saves the value to a property
        m_lIncrement = -1
    End If
End Sub

Private Sub Increments_1_Change()
    If Increments_1.Value Then
        ' Disables all other increments
        Increments_0.Value = False
        Increments_10.Value = False
        Increments_100.Value = False
        
        ' Sets the value of this control
        Increments_1.Value = True
    
        ' Saves the value to a property
        m_lIncrement = 1
      Else
        ' Saves the value to a property
        m_lIncrement = -1
    End If
End Sub

Private Sub Increments_10_Change()
    If Increments_10.Value Then
        ' Sets the value of this control
        Increments_10.Value = True
    
        ' Disables all other increments
        Increments_0.Value = False
        Increments_1.Value = False
        Increments_100.Value = False
        
        ' Saves the value to a property
        m_lIncrement = 10
      Else
        ' Saves the value to a property
        m_lIncrement = -1
    End If
End Sub

Private Sub Increments_100_Change()
    If Increments_100.Value Then
        ' Sets the value of this control
        Increments_100.Value = True
    
        ' Disables all other increments
        Increments_0.Value = False
        Increments_1.Value = False
        Increments_10.Value = False
        
        ' Saves the value to a property
        m_lIncrement = 100
      Else
        ' Saves the value to a property
        m_lIncrement = -1
    End If
End Sub
