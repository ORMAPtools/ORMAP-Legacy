Attribute VB_Name = "basErrorHandling"
' File name:            basErrorHandling
'
' Initial Author:       Environmental Systems Research Institute (ESRI)
'
' Date Created:         <<Unknown>>
'
' Description:
'       Common error handling routines
'
' Entry points:
'
'
'
' Dependencies:
'       File References
'           ESRI Error Handler v1.0
'
' Issues:
'       All versions of this error handler, with the exception of 4, produce
'       an error.  The error is that a error dialog will show with a modal
'       dialog box behind.  This gives the effect of ArcMap locking up, and
'       is a difficult issue for the user to resolve.  Consequently, all
'       handler for less than 4 have been removed, and all versions are
'       referred to version 4.
'
' Method:
'
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
Dim pErrorLog As New ErrorHandlerUI.ErrorDialog

'++ START JWalton 1/26/2007
    ' Removed the following procedures:
    '   DisplayVersion2Dialog
    '   DisplayVersion3Dialog
    ' These routines cause a dialog to appear on the application menu.
    ' The dialog hides behind the error log, and makes the application
    ' appear to be locked
'++ END JWalton 1/26/2007

'***************************************************************************
'Name:                  DisplayVersion4Dialog
'Initial Author:        Environmental Systems Research Institute (ESRI)
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Append error data to the error dialog box
'Called From:   HandleError
'Description:   Append error data to the error dialog box
'Methods:       None
'Inputs:        sProcedureName, sErrDescription, parentHwnd
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************

Private Sub DisplayVersion4Dialog( _
  ByVal sProcedureName As String, _
  ByVal sErrDescription As String, _
  ByVal parentHWND As Long)
    pErrorLog.AppendErrorText "Record Call Stack Sequence - Bottom line is error line." & vbCrLf & _
                              vbCrLf & _
                              vbTab & sProcedureName & vbCrLf & _
                              sErrDescription
    pErrorLog.Visible = True
    
    Dim objFileSys As Scripting.FileSystemObject
    Dim objFile As Scripting.TextStream
    
    Set objFileSys = New Scripting.FileSystemObject
    Set objFile = objFileSys.OpenTextFile(gfn_s_GetWindowsTempPath & "/Errors.log", ForAppending, True)
    objFile.WriteBlankLines 2
    objFile.WriteLine "Record Call Stack Sequence - Bottom line is error line." & vbCrLf & _
                              vbCrLf & _
                              vbTab & sProcedureName & vbCrLf & _
                              sErrDescription
    objFile.Close
End Sub

'***************************************************************************
'Name:                  HandleError
'Initial Author:        Environmental Systems Research Institute (ESRI)
'Subsequent Author:     <<Type your name here>>
'Created:               <<Unknown>>
'Purpose:       Common Error Handling
'Called From:   Multiple Locations
'Description:   Handles errors in a common manner across the whole DLL
'Methods:       None
'Inputs:
'   bTopProcedure -- True if called from a top level procedure
'   sProcedureName -- Name of function called from
'   lErrNumber -- Error Number
'   sErrSource -- Error Source
'   sErrDescription -- Error Description
'   version -- Version of Function
'   parentHWND -- Parent Hwnd for error dialogs
'   reserved1 -- Reserved for later use
'   reserved2 -- Reserved for later use
'   reserved3 -- Reserved for later use
'Parameters:    None
'Outputs:       None
'Returns:       None
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************

Public Sub HandleError( _
  ByVal bTopProcedure As Boolean, _
  ByVal sProcedureName As String, _
  ByVal lErrNumber As Long, _
  ByVal sErrSource As String, _
  ByVal sErrDescription As String, _
  Optional ByVal version As Long = 1, _
  Optional ByVal parentHWND As Long = 0, _
  Optional ByVal reserved1 As Variant = 0, _
  Optional ByVal reserved2 As Variant = 0, _
  Optional ByVal reserved3 As Variant = 0)
    ' Clear the error object
    Err.Clear

    ' Static variable used to control the call stack formatting
    Static entered As Boolean

    If (bTopProcedure) Then
        ' Top most procedure in call stack so report error to user
        ' Via a dialog
        If (Not entered) Then
          sErrDescription = vbCrLf & _
                           "Error Number " & vbCrLf & _
                           vbTab & CStr(lErrNumber) & vbCrLf & _
                           "Description" & vbCrLf & _
                           vbTab & sErrDescription & vbCrLf & _
                           vbCrLf
        End If
        entered = False
'++ START JWalton 1/26/2007 Referred handling of all cases to version 4
        DisplayVersion4Dialog sProcedureName, sErrDescription, parentHWND
'++ END JWalton 1/26/2007
      Else
        ' An error has occured but we are not at the top of the call stack
        ' so append the callstack and raise another error
        If (Not entered) Then sErrDescription = vbCrLf & _
                                                "Error Number " & vbCrLf & _
                                                vbTab & CStr(lErrNumber) & vbCrLf & _
                                                "Description" & vbCrLf & _
                                                vbTab & sErrDescription & vbCrLf & _
                                                vbCrLf
        entered = True
        Err.Raise lErrNumber, _
                  sErrSource, _
                  vbTab & sProcedureName & vbCrLf & sErrDescription
    End If
End Sub

'***************************************************************************
'Name:                  GetErrorLineNumberString
'Initial Author:        Environmental Systems Research Institute (ESRI)
'Subsequent Author:     <Type your name here>
'Created:               2/5/2007
'Purpose:       Determines whether or not an error number is valid
'Called From:   Multiple Locations
'Description:   Tests the error number and adds it to the error number
'               string if it is not zero.
'Methods:       None
'Inputs:        lLineNumber
'Parameters:    None
'Outputs:       None
'Returns:       A string representing the line number
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'
'***************************************************************************

Public Function GetErrorLineNumberString( _
  ByVal lLineNumber As Long) As String
    ' Test the line number if it is non zero create a string
    If (lLineNumber <> 0) Then GetErrorLineNumberString = "Line : " & lLineNumber
End Function

