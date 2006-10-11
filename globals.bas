Attribute VB_Name = "globals"
'
' File name:            globals
'
' Initial Author:       Type your name here
'
' Date Created:
'
' Description:
'       Short description of the file's overall purpose.
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
'JWM 10/11/2006 added this file header

Option Explicit
'******************************
' Global/Public Definitions
'------------------------------
' Public API Declarations
'------------------------------
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'------------------------------
' Public Enums and Constants
'------------------------------

'------------------------------
' Public variables
'------------------------------
Public g_pApp As esriFramework.IApplication
Public g_pFldnames As New clsFieldNames
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
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms
'------------------------------
'Private Constants and Enums
'------------------------------
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "globals.bas"

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------






