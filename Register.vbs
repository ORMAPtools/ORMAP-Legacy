' File name:            Register.vbs
'
' Initial Author:       JWalton
'
' Date Created:         4/11/2007
'
' Description: Run CatInstall.exe utility and register the ORMAP Dll
'
' Entry points:
'       <<None>>
'
' Dependencies:
'       <<None>>
'
' Issues:
'       This script file must be located in the same directory as the Dll.
'
' Method:
'       <<None>>
'
' Updates:
'       4/11/2007 -- Initial implementation. (JWalton)


' Variable declarations
Dim objWshShell
Dim strFullPath
Dim strFileName
Dim strPath 
Dim strCatInstall

' Initialize the script shell object
Set objWshShell=WScript.CreateObject("WScript.Shell")

' Initialize path to CatInstall.exe
strCatInstall="""" & "C:\Program Files\ArcGIS\Bin\CatInstall.exe" & """"

' Determine the path to the Dll
strFullPath=WScript.ScriptFullName
strFileName=WScript.ScriptName
strPath=Left(strFullPath,Len(strFullPath)-len(strFileName))

' Register the Dll
On Error Resume Next
objWshShell.Run(strCatInstall & " " & """" & strPath & "TaxlotEditing.dll" & """")


