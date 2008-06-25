Attribute VB_Name = "basWin32API"

' Keyword expansion for source code control
' Tag for this file : $Name$
' SCC Revision number: $Revision$
' Date of last change: $Date$

' File name:            basWin32API.bas
'
' Initial Author:       John Walton
'
' Date Created:         2/6/2007
'
' Description:          32-bit Windows API Function Definitions
'
'
' Entry points:
'       SetParent
'       SetWindowPos
'       Sleep
'
' Dependencies:
'       Windows API Functions are called from the following DLL files
'           User32.dll
'
'
' Issues:
'       No known issues.
'
' Method:
'       Generic Windows function calls.
'
' Updates:
'       2/7/2007 -- Implemented (JWalton)
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)

Option Explicit
'******************************
' Global/Public Definitions
'------------------------------
' Public API Declarations
'------------------------------
Public Declare Function GetPrivateProfileString _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" _
               (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpDefault As String, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpFileName As String) As Long
Public Declare Function GetUserName _
               Lib "advapi32.dll" _
               Alias "GetUserNameA" _
               (ByVal lpBuffer As String, _
                nSize As Long) As Long
Public Declare Function SetParent _
               Lib "user32" _
               (ByVal hWndChild As Long, _
                ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos _
               Lib "user32" _
               (ByVal hwnd As Long, _
                ByVal hWndInsertAfter As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal cx As Long, _
                ByVal cy As Long, _
                ByVal wFlags As Long) As Long
Public Declare Function SendMessageString _
               Lib "user32" _
               Alias "SendMessageA" _
               (ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As String) As Long
Public Declare Function ShellExecute& _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" _
               (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long)
Public Declare Sub Sleep _
               Lib "kernel32" _
               (ByVal dwMilliseconds As Long)
Public Declare Function GetTempPath Lib "kernel32" _
                            Alias "GetTempPathA" _
                            (ByVal nBufferLength As Long, _
                            ByVal lpBuffer As String) As Long

'------------------------------
' Public Enums and Constants
'------------------------------
' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200

' Listbox message strings
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2

' Combobox message strings
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158

' Error constants
Public Const SW_SHOWNORMAL = 1
Public Const SE_ERR_FNF = 2&
Public Const SE_ERR_PNF = 3&
Public Const SE_ERR_ACCESSDENIED = 5&
Public Const SE_ERR_OOM = 8&
Public Const SE_ERR_DLLNOTFOUND = 32&
Public Const SE_ERR_SHARE = 26&
Public Const SE_ERR_ASSOCINCOMPLETE = 27&
Public Const SE_ERR_DDETIMEOUT = 28&
Public Const SE_ERR_DDEFAIL = 29&
Public Const SE_ERR_DDEBUSY = 30&
Public Const SE_ERR_NOASSOC = 31&
Public Const ERROR_BAD_FORMAT = 11&

' String constants for GetUserName
Public Const UNLEN = 256                               ' Maximum username length
Public Const UNLEN_MAX = UNLEN + 1                     ' Maximum username length buffer
