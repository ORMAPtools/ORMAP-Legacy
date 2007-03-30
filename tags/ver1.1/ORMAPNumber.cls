VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ORMAPNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Keyword expansion for source code control
' Tag for this file : $Name$
' SCC Revision number: $Revision: 18 $
' Date of last change: $Date: 2006-11-14 13:15:34 -0800 (Tue, 14 Nov 2006) $
'
'
'
' File name:            ORMAPNumber
'
' Initial Author:       JWalton
'
' Date Created:         2/5/2007
'
' Description:
'       Encapsulates all elements of a ORMAP number into one multipurpose class that
'       allows an ORMAP number to be either created or parsed, manipulated, or
'       validated against the current ORMAP number model.
'
' Entry points:
'   Public Properties:
'       Anomaly
'           Two digit code for any oddball situations
'       County
'           Two digit code for the Oregon County -- Default: 00
'       PartialRangeCode
'           Three-digit code for the partial range code -- Default: .00
'       PartialTownshipCode
'           Three digit code for the partial township code -- Default: .00
'       Range
'           Two digit code for the range -- Default: 01
'       RangeDirectional
'           One digit code for the directional for the range -- Default: W
'       Quarter
'           One digit code for the quarter from A to J -- Default: 0
'       QuarterQuarter
'           One digit code for the quarter/quarter from A to J -- Default: 0
'       SuffixNumber
'           Three digit code for the suffix number from 000 to 999 -- Default: 000
'       SuffixType
'           One digit code, S, D, T, or 0, for suffix type -- Default: 0
'       Section
'           Two digit code for the section number from 00 to 37 -- Default: 00
'       Township
'           Two digit code for the township -- Default: 01
'       TownshipDirectional
'           One digit directional for the township -- Default: N
'
'   Public Methods:
'       ParseNumber
'           Parse an ORMAP Number into its component pieces
'       ORMAPNumber
'           Returns a properly formatted ORMAP Number
'       ORMAPMapNumber
'           Returns a properly formatted ORMAP Number minus the County
'       ORMAPTaxlotNumber
'           Returns a properly formatted ORMAP Number
'
' Dependencies:
'       File Dependencies
'           basGlobals
' Issues:
'       None that are known at this time (2/5/2007).
'
' Method:
'       This class should be fairly straight forward. All parsing, creation, and
'       validation of ORMAP Numbers is simply amalgamated here instead of being
'       spread across multiple classes, modules, and forms.
'
' Updates:
'       2/5/2007 -- Implemented (JWalton)
'       2/6/2007 -- All inline documentation reviewed/revised (JWalton)


Option Explicit
'******************************
' Event Definitions
'------------------------------
Event OnChange(ByVal sNewNumber As String)

'******************************
' Private Definitions
'------------------------------
' Private Variables
'------------------------------
Private m_sCounty As String
Private m_sTownship As String
Private m_sTownPartialCode As String
Private m_sTownDirectional As String
Private m_sRange As String
Private m_sRangePartialCode As String
Private m_sRangeDirectional As String
Private m_sSection As String
Private m_sQtr As String
Private m_sQtrQtr As String
Private m_sSuffixType As String
Private m_sSuffixNumber As String
Private m_sAnomaly As String

Public Property Let County( _
  ByVal sValue As String)
    ' Variable declarations
    Dim iLength As Integer

    ' Extracts and formats the ORMAP County value
    iLength = Len(sValue)
    Select Case iLength
      Case Is < 2
        m_sCounty = String(2 - iLength, "0") & sValue
      Case 2
        m_sCounty = sValue
      Case Is > 2
        m_sCounty = Left(sValue, 2)
    End Select
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get County() As String
    County = m_sCounty
End Property

Public Property Let Township( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Township value
    If Len(sValue) <> 2 Then
        m_sTownship = "00"
      Else
        m_sTownship = sValue
    End If
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get Township() As String
    Township = m_sTownship
End Property

Public Property Let PartialTownshipCode( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Partial Township Code value
    Select Case sValue
      Case "0.25", "0.50", "0.75"
        m_sTownPartialCode = Mid$(sValue, 2)
      Case Else
        m_sTownPartialCode = ".00"
    End Select
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get PartialTownshipCode() As String
    PartialTownshipCode = m_sTownPartialCode
End Property

Public Property Let TownshipDirectional( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Township Directional value
    Select Case sValue
      Case "N", "S"
        m_sTownDirectional = sValue
      Case Else
        m_sTownDirectional = "N"
    End Select
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get TownshipDirectional() As String
    TownshipDirectional = m_sTownDirectional
End Property

Public Property Let Range( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Range value
    If Len(sValue) <> 2 Then
        m_sRange = "01"
      Else
        m_sRange = sValue
    End If
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get Range() As String
    Range = m_sRange
End Property

Public Property Let PartialRangeCode( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Partial Range Code value
    Select Case sValue
      Case "0.25", "0.50", "0.75"
        m_sRangePartialCode = Mid$(sValue, 2)
      Case Else
        m_sRangePartialCode = ".00"
    End Select
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get PartialRangeCode() As String
    PartialRangeCode = m_sRangePartialCode
End Property

Public Property Let RangeDirectional( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Range Directional value
    Select Case sValue
      Case "E", "W"
        m_sRangeDirectional = sValue
      Case Else
        m_sRangeDirectional = "W"
    End Select
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get RangeDirectional() As String
    RangeDirectional = m_sRangeDirectional
End Property

Public Property Let Section( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Section value
    If IsNumeric(sValue) Then
        Select Case CInt(sValue)
          Case 0
            m_sSection = "00"
          Case Is < 10
            m_sSection = "0" & CInt(sValue)
          Case Is <= 37
            m_sSection = sValue
          Case Else
            m_sSection = "00"
        End Select
      Else
        m_sSection = "00"
    End If
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get Section() As String
    Section = m_sSection
End Property

Public Property Let Quarter( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Quarter value
    Select Case UCase(sValue)
      Case "0", "A" To "J"
        m_sQtr = sValue
      Case Else
        m_sQtr = "0"
    End Select
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get Quarter() As String
    Quarter = m_sQtr
End Property

Public Property Let QuarterQuarter( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP QuarterQuarter value
    Select Case UCase(sValue)
      Case "0", "A" To "J"
        m_sQtrQtr = sValue
      Case Else
        m_sQtrQtr = "0"
    End Select
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get QuarterQuarter() As String
    QuarterQuarter = m_sQtrQtr
End Property

Public Property Let SuffixType( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Suffix Type value
    Select Case UCase(sValue)
      Case "0", "D", "S", "T"
        m_sSuffixType = sValue
      Case Else
        m_sSuffixType = "0"
    End Select
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get SuffixType() As String
    SuffixType = m_sSuffixType
End Property

Public Property Let SuffixNumber( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Suffix Number value
    If IsNumeric(sValue) Then
        Select Case CInt(sValue)
          Case Is < 0
            m_sSuffixNumber = "00"
          Case Is < 1000
            m_sSuffixNumber = String(3 - Len(sValue), "0") & sValue
          Case Else
            m_sSuffixNumber = "00"
        End Select
      Else
        m_sSuffixNumber = "00"
    End If
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get SuffixNumber() As String
    SuffixNumber = m_sSuffixNumber
End Property

Public Property Let Anomaly( _
  ByVal sValue As String)
    ' Extracts and formats the ORMAP Anomaly value
    If Len(sValue) > 2 Then
        m_sAnomaly = "00"
      Else
        m_sAnomaly = String(2 - Len(sValue), "0") & sValue
    End If
    
    ' Raise the ORMAP Number Change event
    RaiseEvent OnChange(Me.ORMAPNumber)
End Property

Public Property Get Anomaly() As String
    Anomaly = m_sAnomaly
End Property

'***************************************************************************
'Name:  IsValidNumber
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Validate ORMAP Numbers
'Called From:   Multiple Locations
'Description:   Validates the ORMAP Number represented the class instance
'Methods:       Determines validity based on all elements having a length of
'               greater than 0
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       Boolean value representing the Valid status of the number
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007        Initial creation
'***************************************************************************

Public Function IsValidNumber() As Boolean
    ' Insures that each element of the ORMAP Number exists
    IsValidNumber = True
    IsValidNumber = IsValidNumber And (Len(m_sCounty) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sTownship) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sTownPartialCode) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sTownDirectional) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sRange) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sRangePartialCode) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sRangeDirectional) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sSection) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sQtr) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sQtrQtr) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sSuffixType) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sSuffixNumber) > 0)
    IsValidNumber = IsValidNumber And (Len(m_sAnomaly) > 0)
End Function

'***************************************************************************
'Name:                  OrmapNumber
'Initial Author:        John Walton
'Subsequent Author:     Type your name here.
'Created:       2/5/2007
'Purpose:       Create ORMAP Number
'Called From:   Multiple Locations
'Description:   Creates an ORMAP Number from the sub-components stored as
'               properties in the class instance
'Methods:       None
'Inputs:        None
'Parameters:    None
'Outputs:       None
'Returns:       A properly formatted ORMAP Number as a string
'Errors:        This routine raises no known errors.
'Assumptions:   None
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007    Initial creation
'***************************************************************************

Public Function ORMAPNumber() As String
    ' Variable declarations
    Dim sNumber As String
    
    ' Creates a formatted ORMAP Map Number
    If IsValidNumber Then
        sNumber = m_sCounty
        sNumber = sNumber & m_sTownship
        sNumber = sNumber & m_sTownPartialCode
        sNumber = sNumber & m_sTownDirectional
        sNumber = sNumber & m_sRange
        sNumber = sNumber & m_sRangePartialCode
        sNumber = sNumber & m_sRangeDirectional
        sNumber = sNumber & m_sSection
        sNumber = sNumber & m_sQtr
        sNumber = sNumber & m_sQtrQtr
        sNumber = sNumber & m_sAnomaly
        sNumber = sNumber & m_sSuffixType
        sNumber = sNumber & m_sSuffixNumber
      Else
        sNumber = ""
    End If
    
    ' Returns the value of the function
    ORMAPNumber = sNumber
End Function

'***************************************************************************
'Name:                  OrmapTaxlotNumber
'Initial Author:        John Walton
'Subsequent Author:     <Type your name here>
'Created:       2/5/2007
'Purpose:       Create base ORMAP Taxlot Number
'Called From:   Multiple Locations
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007        Initial creation
'***************************************************************************

Public Function OrmapTaxlotNumber() As String
    ' Variable declarations
    Dim sNumber As String
    
    ' Creates a formatted ORMAP Map Number
    If IsValidNumber Then
        sNumber = m_sCounty
        sNumber = sNumber & m_sTownship
        sNumber = sNumber & m_sTownPartialCode
        sNumber = sNumber & m_sTownDirectional
        sNumber = sNumber & m_sRange
        sNumber = sNumber & m_sRangePartialCode
        sNumber = sNumber & m_sRangeDirectional
        sNumber = sNumber & m_sSection
        sNumber = sNumber & m_sQtr
        sNumber = sNumber & m_sQtrQtr
        sNumber = sNumber & m_sAnomaly
        sNumber = sNumber & m_sSuffixType
        sNumber = sNumber & m_sSuffixNumber
      Else
        sNumber = ""
    End If
    
    ' Returns the value of the function
    OrmapTaxlotNumber = sNumber
End Function

'***************************************************************************
'Name:  OrmapMapNumber
'Initial Author:        John Walton
'Subsequent Author:     Type your name here.
'Created:       2/5/2007
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007        Initial creation
'***************************************************************************

Public Function OrmapMapNumber() As String
    ' Variable declarations
    Dim sNumber As String
    
    ' Creates a formatted ORMAP Map Number
    If IsValidNumber Then
        sNumber = m_sTownship
        sNumber = sNumber & m_sTownPartialCode
        sNumber = sNumber & m_sTownDirectional
        sNumber = sNumber & m_sRange
        sNumber = sNumber & m_sRangePartialCode
        sNumber = sNumber & m_sRangeDirectional
        sNumber = sNumber & m_sSection
        sNumber = sNumber & m_sQtr
        sNumber = sNumber & m_sQtrQtr
        sNumber = sNumber & m_sAnomaly
        sNumber = sNumber & m_sSuffixType
        sNumber = sNumber & m_sSuffixNumber
      Else
        sNumber = ""
    End If
    
    ' Returns the value of the function
    OrmapMapNumber = sNumber
End Function

'***************************************************************************
'Name:  ParseNumber
'Initial Author:        John Walton
'Subsequent Author:     Type your name here.
'Created:       2/5/2007
'Purpose:
'Called From:
'Description:   Type the description of the function here.
'Methods:       Describe any complex details.
'Inputs:        What variables are brought into this routine?
'Parameters:
'Outputs:       What variables are changed in this routine?
'Returns:
'Errors:        What known errors does this routine cause that are NOT captured in error handling routine? If none, say: This routine raises no known errors.
'Assumptions:   What parameters or variable values are assumed to be true?
'Updates:
'       Type any updates here.
'Developer:     Date:       Comments:
'----------     ------      ---------
'John Walton    2/5/2007        Initial creation
'***************************************************************************

Public Function ParseNumber( _
  ByVal sNumber As String) As Boolean
On Error GoTo Err_Handler
    ' Validates the length of the field value
    If Len(sNumber) <> ORMAP_MAPNUM_FIELD_LENGTH Then
        ParseNumber = False
        Exit Function
    End If

    ' Parse County'
    Me.County = Mid$(sNumber, 1, 2)
    
    ' Parse Township'
    Me.Township = Mid$(sNumber, 3, 2)
    
    ' Parse Partial Township Code'
    Me.PartialTownshipCode = Mid$(sNumber, 5, 3)
    
    ' Parse Township Directional'
    Me.TownshipDirectional = Mid$(sNumber, 8, 1)
    
    ' Parse Range'
    Me.Range = Mid$(sNumber, 9, 2)
    
    ' Parse Partial Range Code'
    Me.PartialRangeCode = Mid$(sNumber, 11, 3)
    
    ' Parse Range Directional'
    Me.RangeDirectional = Mid$(sNumber, 14, 1)
    
    ' Parse Section'
    Me.Section = Mid$(sNumber, 15, 2)
    
    ' Parse Quarter'
    Me.Quarter = Mid$(sNumber, 17, 1)
    
    ' Parse QuarterQuarter'
    Me.QuarterQuarter = Mid$(sNumber, 18, 1)
    
    ' Parse Anomaly
    Me.Anomaly = Mid$(sNumber, 19, 2)
    
    ' Parse Map Suffix Type
    Me.SuffixType = Mid$(sNumber, 21, 1)
    
    ' Parse Map Suffix Number
    Me.SuffixNumber = Mid$(sNumber, 22, 3)
    
Err_Handler_Resume:
    ParseNumber = True
    Exit Function
    
Err_Handler:
    ParseNumber = False
    Resume
End Function
