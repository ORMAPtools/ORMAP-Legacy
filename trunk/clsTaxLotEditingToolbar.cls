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
'CLASS FOR THE TAXLOTEDITING TOOLBAR
Implements IToolBarDef
 
Private Property Get IToolBarDef_ItemCount() As Long

6:     IToolBarDef_ItemCount = 5
    
End Property
 
Private Sub IToolBarDef_GetItemInfo(ByVal pos As Long, ByVal itemDef As esriSystemUI.IItemDef)
    
  Select Case pos
    Case 0
      'itemDef.Group = True
15:       itemDef.ID = "TaxlotEditing.cmdTaxlotAssignment"
    Case 1
17:       itemDef.ID = "TaxlotEditing.cmdLocate"
    Case 2
19:       itemDef.ID = "TaxlotEditing.cmdTaxlotCombine"
    Case 3
21:       itemDef.ID = "TaxlotEditing.cmdArrows"
    Case 4
23:       itemDef.ID = "TaxlotEditing.cmdMapIndex"
24:   End Select
    
End Sub
 
Private Property Get IToolBarDef_Name() As String
    
30:     IToolBarDef_Name = "ORMAP Taxlot Editor"
    
End Property
 
Private Property Get IToolBarDef_Caption() As String

36:     IToolBarDef_Caption = "ORMAP Taxlot Editor"
    
End Property
