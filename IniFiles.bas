Attribute VB_Name = "IniFiles"
Attribute VB_Description = "Set of function to read and extract values from INI files"
'MODULE THAT PROCESSES INI FILE
Option Explicit

Private mIniFile As String

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Function TestIniFile(FilePath As String) As Boolean
10:   mIniFile = FilePath
11:   TestIniFile = Len(Dir$(FilePath)) > 0
End Function

Public Function ReadIniFile(ByVal strSection As String, ByVal strKey As String) As String
  
  Dim strBuffer As String
  Dim intPos As Integer
18:   strBuffer = Space$(255)
19:   If GetPrivateProfileString(strSection, strKey, "", strBuffer, 255, mIniFile) > 0 Then
20:     ReadIniFile = RTrim$(StripTerminator(strBuffer))
21:   Else
22:     ReadIniFile = ""
23:   End If
24:   strBuffer = ""
End Function

Private Function StripTerminator(ByVal strString As String) As String
  'function to strip out chr$(0) from the ReadIniFile function
  Dim intZeroPos As Integer
30:   intZeroPos = InStr(strString, Chr$(0))
31:   If intZeroPos > 0 Then
32:     StripTerminator = Left$(strString, intZeroPos - 1)
33:   Else
34:     StripTerminator = strString
35:   End If
End Function





