VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTaxLotEditingToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' File name:            clsTaxLotEditingToolbar
'
' Initial Author:       Type your name here
'
' Date Created:     10/11/2006
'
' Description: CLASS FOR THE TAXLOTEDITING TOOLBAR
'
'
' Entry points:
'       List the public variables and their purposes.
'       List the properties and routines that the module exposes to the rest of the program.
'
' Dependencies:
'       How does this file depend or relate to other files?
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
'++  JWM 10/11/2006 added this file header

Option Explicit
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

'------------------------------
'Private Constants and Enums
'------------------------------

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------


Implements IToolBarDef
 
Private Property Get IToolBarDef_ItemCount() As Long

81:     IToolBarDef_ItemCount = 5
    
End Property
 
Private Sub IToolBarDef_GetItemInfo(ByVal pos As Long, ByVal itemDef As esriSystemUI.IItemDef)
    
  Select Case pos
    Case 0
      'itemDef.Group = True
90:       itemDef.ID = "TaxlotEditing.cmdTaxlotAssignment"
    Case 1
92:       itemDef.ID = "TaxlotEditing.cmdLocate"
    Case 2
94:       itemDef.ID = "TaxlotEditing.cmdTaxlotCombine"
    Case 3
96:       itemDef.ID = "TaxlotEditing.cmdArrows"
    Case 4
98:       itemDef.ID = "TaxlotEditing.cmdMapIndex"
99:   End Select
    
End Sub
 
Private Property Get IToolBarDef_Name() As String
    
105:     IToolBarDef_Name = "ORMAP Taxlot Editor"
    
End Property
 
Private Property Get IToolBarDef_Caption() As String

111:     IToolBarDef_Caption = "ORMAP Taxlot Editor"
    
End Property
