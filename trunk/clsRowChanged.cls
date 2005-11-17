VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRowChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'During the compare and save process, set a flag so the app
' knows whether to store the row when done comparing
Private m_bRowChanged As Boolean

Private Sub Class_Initialize()

8:   m_bRowChanged = False
  
End Sub

Public Property Let RowChanged(bRowChanged As Boolean)
13:   m_bRowChanged = bRowChanged
End Property

Public Property Get RowChanged() As Boolean
17:   RowChanged = m_bRowChanged
End Property
