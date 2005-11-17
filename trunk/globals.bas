Attribute VB_Name = "globals"
Option Explicit

Public g_pApp As esriFramework.IApplication
Public g_pFldnames As New clsFieldNames
'Public g_pAutoUpdate As New clsAutoUpdate

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Variables used by the Error handler function - DO NOT REMOVE

Const c_sModuleFileName As String = "globals.bas"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms

