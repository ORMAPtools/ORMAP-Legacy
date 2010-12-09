Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Text
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	
	Private Sub cmdApplyMask_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdApplyMask.Click


        txtNewOutput.Text = String.Empty
        txtOut.Text = String.Empty
        txtNewOutput.Text = CreateMapTaxlotValue(txtIn.Text, txtMask.Text)
        txtOut.Text = gfn_s_CreateMapTaxlotValue(txtIn.Text, txtMask.Text)
	End Sub
	
	'************************************************************
	'Name:  gfn_l_CountTokens
	'Purpose: Given a string of token characters and a single character token to search for, the number
	'           of tokens in the string will be returned. This function is useful
	'           for dimensioning an array to store the delimited items.
	'Called From:
	'Inputs:    sSource: A list of tokens
	'           sToken:  The character token to search for.
	'Outputs:   None
	'Return value:  The number of tokens in sSource. If sSource is empty, 0 is returned.
	'Method:    This function uses Unicode representation of characters
	'Errors:    This routine raises no known errors.
	'Assumptions:What values or params are assumed to be true?
	'Post-conditions:
	'Pre-conditions:
	'Developer:     Date:           Comments:
	'----------     ----------      ----------
	'James Moore    September,23 05 Initial creation
	'************************************************************
	
	Function gfn_l_CountTokens(ByVal ps_Source As String, ByRef ps_Token As String) As Integer
		' Number of tokens = 0 if the source string is empty
		' or there is no token to count
		Dim ll_Count As Integer
		Dim i As Integer
		Dim MyByteArray() As Byte
		Dim ll_UnicodeValue As Integer
		If Len(ps_Source) = 0 Or Len(ps_Token) = 0 Then
			gfn_l_CountTokens = 0
		Else
			
			'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
			MyByteArray = System.Text.UnicodeEncoding.Unicode.GetBytes(ps_Source) 'this assignment creates a unicode character array
			ll_UnicodeValue = AscW(ps_Token) 'The AscW() function returns the Unicode character code
			For i = 0 To UBound(MyByteArray) Step 2 'this is a Unicode byte array so we must step by 2
				' if this is the char, increase the counter
				If MyByteArray(i) = ll_UnicodeValue Then ll_Count = ll_Count + 1
			Next i
			gfn_l_CountTokens = ll_Count
		End If
	End Function
	'************************************************************
	'Name:  ffn_s_CreateParcelID
	'Purpose:   Create a parcel ID from a mask.
	'Called From:
	'Inputs:
	'Outputs:What variables are changed in this routine?
	'Return value: If a value is passed in that is not numeric then just pass it straight through
	'       else return a parcel id with or without leading zeros
	'Method:    I use the Format function with user-defined string formats
	'       which consist of either all at (@) characters or all ampersands (&)
	'Errors:    This routine raises no known errors
	'Assumptions: That the mask will be all ampersands or @ characters
	'
	'Post-conditions:
	'Pre-conditions:
	'Developer:     Date:           Comments:
	'----------     ----------      ----------
	'James Moore    October,12 05   Initial Creation
	'************************************************************
	Private Function ffn_s_CreateParcelID(ByRef ps_ValueToMask As String, ByVal ps_MaskToApply As String) As String
		On Error GoTo ffn_s_CreateParcelID_Error
		
		If Len(ps_MaskToApply) = 0 Or Len(ps_ValueToMask) = 0 Then
			GoTo ProcessExit
		End If
		
		Dim ls_Temp As String
		ls_Temp = Space(Len(ps_MaskToApply))
		'add exclamation point to mask so that the string will be formatted left to right
		ps_MaskToApply = "!" & ps_MaskToApply
		If IsNumeric(ps_ValueToMask) Then
			ls_Temp = VB6.Format(ps_ValueToMask, ps_MaskToApply)
		Else
			ls_Temp = ps_ValueToMask
		End If
		ffn_s_CreateParcelID = ls_Temp
		
ProcessExit:

        Exit Function

ffn_s_CreateParcelID_Error:
        Return String.Empty
		MsgBox("Error has occurred in ffn_s_CreateParcelID" & Err.Description, MsgBoxStyle.OKOnly + MsgBoxStyle.Critical, "Error")
	End Function
	'************************************************************
	'Name:  ffn_s_StripLeadingZeros
	'Purpose:   Remove leading zeros from a string
	'Called From:
	'Inputs:    psStringToStrip: A string that may have leading zeros
	'Return value: A string of same length with blank spaces instead of leading zeros.
	'Method:
	'Errors:This routine raises no known errors
	'Assumptions:   A string with leading zeros may not be passed in.
	'               In that case the whole string will be returned
	'Post-conditions:
	'Pre-conditions:
	'Developer:     Date:           Comments:
	'----------     ----------      ----------
	'James Moore    September,23 05   Initial creation
	'************************************************************
	
	Private Function ffn_s_StripLeadingZeros(ByRef ps_StringToParse As String) As String
		Dim ll_InputCharCount As Integer
		Dim ll_Counter As Integer
		Dim ls_Char, ls_Temp As String
		
		ll_InputCharCount = Len(ps_StringToParse)
		ls_Temp = Space(ll_InputCharCount) 'create string of same length
		
		For ll_Counter = 1 To ll_InputCharCount
			ls_Char = Mid(ps_StringToParse, ll_Counter, 1)
			If InStr(1, "0", ls_Char, CompareMethod.Text) < 1 Then 'go past all leading zeros
				Mid(ls_Temp, ll_Counter) = Mid(ps_StringToParse, ll_Counter) 'get all remaing chars
				Exit For 'and exit
			End If
		Next ll_Counter
		ffn_s_StripLeadingZeros = ls_Temp ' do not trim off leading spaces
	End Function
	
	
	
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		lblInfo.Text = "TT for Township, RR for Range, D for direction, SS for Section, QQ for section breakdown, (this piece is disabled II for map suffix type), @ for the parcel ID with leading zeros, && for the parcel ID without leading zeros."
	End Sub
	
	'***************************************************************************
	'Name:                  gfn_s_CreateMapTaxlotValue
	'Initial Author:        James Moore
	'Subsequent Author:     <<Type your name here>>
	'Created:               9-23-2005
	'Purpose:       Use the ORMapTaxlot value to create a MapTaxlot value based
	'               on the mask from the ini file.
	'Called From:   basUtilities.CalcTaxlotValues
	'               basUtilities.gsb_StartDoc
	'               cmdTaxlotAssignment.ITool_OnMouseDown
	'               frmMapIndex.UpdateTaxlots
	'Methods:       The string parsing procedures depends on a valid ORTaxlot
	'               string 29 characters long as defined in version 1.3 of the
	'               ORMAP data structure.
	'               Extensive use of the Mid function causes heavy reliance on
	'               the position of values in the string
	'               Need to find a better way to handle half townships and
	'               ranges
	'               May want to change this function to use Regular Expressions
	'Inputs:        as_ORMapTaxlotString: The ORTaxlot value
	'               as_MaskFormatString: the formatting string
	'Parameters:    None
	'Outputs:       None
	'Returns:       A formatted string that can be used as parcel ID or/and as
	'               a MapTaxlot value
	'Errors:        This routine raises no known errors.
	'Assumptions:   A valid ORTaxlot string and Mask value is passed in.
	'Updates:
	'       Type any updates here.
	'Developer:     Date:       Comments:
	'----------     ------      ---------
	'James Moore    9-23-2005   Initial creation of this routine
	'James Moore    11-8-06     Adding special case for County that uses a Q to
	'                           store half ranges
	'John Walton    2/7/2007    Renamed variables to conform to variable naming
	'                           conventions
	'***************************************************************************
	
	Public Function gfn_s_CreateMapTaxlotValue(ByVal as_ORMapTaxlotString As String, ByRef as_MaskFormatString As String) As String

		'++ START JWalton 2/7/2007 Centralized Variable Declarations
		Dim bProcessedParcelID As Boolean ' Flag
		Dim bProcessedRangeFractional As Boolean
		Dim bHasAlphaQtr As Boolean ' Flag for processing the mask
		Dim bHasAlphaQtrQtr As Boolean ' Flag for processing the mask
		Dim bHasTownPart As Boolean ' Flag for processing the mask
		Dim bHasRangePart As Boolean ' Flag for processing the mask
		Dim bProcessedTownFractional As Boolean
		Dim i As Short
		Dim iCharCode As Short
		Dim iPosCharMaskForward As Short 'Marks the current postion in the mask array
		Dim iMaskLength As Short
		Dim iCountyCode As Short
		Dim lMaskTokenCount As Integer ' How many characters in the mask
		Dim sArr_MaskValues() As String
		Dim sCurrORMapNumValue As String ' To hold a char from ORMAP string
		Dim sFormattedString As String ' The result of our work
		Dim sMaskToApply As String
        Dim sPrevCharInMaskArray As String = " "c ' To use as check for character position
		Dim sTemp As String
		'++ END JWalton 2/7/2007
        Try
            If Len(as_ORMapTaxlotString) = 0 Or Len(as_MaskFormatString) = 0 Then
                Exit Try
            End If

            iCountyCode = CShort(VB.Left(as_ORMapTaxlotString, 2))

            ' flag for half townships,ranges
            bHasTownPart = (Val(Mid(as_ORMapTaxlotString, 5, 3)) > 0)
            bHasRangePart = (Val(Mid(as_ORMapTaxlotString, 11, 3)) > 0)

            'set flags for section qtrs
            Select Case iCountyCode
                Case 1 To 19, 21 To 36
                    If Not IsNumeric(Mid(as_ORMapTaxlotString, 17, 1)) Then
                        bHasAlphaQtr = True
                    End If
                    If Not IsNumeric(Mid(as_ORMapTaxlotString, 18, 1)) Then
                        bHasAlphaQtrQtr = True
                    End If
                Case 20 'lane county uses a totally numeric identifier for qtrs of sections with zeros as placeholders
                    bHasAlphaQtr = False
                    bHasAlphaQtrQtr = False
            End Select

            'We must adjust the mask for clackamas county if there are no  half ranges in the current string
            If InStr(Mid(as_MaskFormatString, 2, 6), "^") > 0 Then
                If bHasRangePart = False Then
                    sMaskToApply = Replace(as_MaskFormatString, "^", vbNullString) 'remove this character
                Else 'if there is a range part the letter Q will be  placed in the position where D sits
                    sMaskToApply = Replace(as_MaskFormatString, "D", vbNullString)
                End If
            Else
                sMaskToApply = as_MaskFormatString
            End If

            iMaskLength = Len(sMaskToApply)

            'Dimension the mask array and fill each position with a character from the mask
            ' I am using an array that begins at dimension one for ease of use
            'UPGRADE_WARNING: Lower bound of array sArr_MaskValues was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            ReDim sArr_MaskValues(iMaskLength)

            For i = 1 To iMaskLength
                sArr_MaskValues(i) = UCase(Mid(sMaskToApply, i, 1))
            Next i

            ' Create a string of spaces to place our results in. This helps a speed up string manipulation a little.
            sFormattedString = Space(iMaskLength)

            For i = 1 To UBound(sArr_MaskValues)
                ' Increment our position in the mask
                iPosCharMaskForward = InStr(i, sMaskToApply, sArr_MaskValues(i), CompareMethod.Text)
                iCharCode = Asc(sArr_MaskValues(i)) 'the ascii value of the character

                ' Returns how many of these characters appear in the mask, AND when used in
                ' Mid function gets/sets that many chars
                lMaskTokenCount = gfn_l_CountTokens(UCase(sMaskToApply), sArr_MaskValues(i))

                Select Case iCharCode
                    Case 68 '"D"
                        If StrComp(sPrevCharInMaskArray, "^", CompareMethod.Text) = 0 Then 'for clackamas county which uses a Q for halfs
                            If StrComp(Mid(sMaskToApply, iPosCharMaskForward - 2, 1), "T", CompareMethod.Text) = 0 Then ' TOWNSHIP DIRECTION
                                Mid(sFormattedString, iPosCharMaskForward, 1) = Mid(as_ORMapTaxlotString, 8, 1)
                            ElseIf StrComp(Mid(sMaskToApply, iPosCharMaskForward - 2, 1), "R", CompareMethod.Text) = 0 Then  'RANGE DIRECTION
                                Mid(sFormattedString, iPosCharMaskForward, 1) = Mid(as_ORMapTaxlotString, 14, 1)
                            End If
                        Else
                            If StrComp(sPrevCharInMaskArray, "T", CompareMethod.Text) = 0 Then ' TOWNSHIP DIRECTION
                                Mid(sFormattedString, iPosCharMaskForward, 1) = Mid(as_ORMapTaxlotString, 8, 1)
                            ElseIf StrComp(sPrevCharInMaskArray, "R", CompareMethod.Text) = 0 Then  'RANGE DIRECTION
                                Mid(sFormattedString, iPosCharMaskForward, 1) = Mid(as_ORMapTaxlotString, 14, 1)
                            End If
                        End If
                        'Formats for the parcel id
                    Case 64 '"@"
                        If Not bProcessedParcelID Then
                            Mid(sFormattedString, iPosCharMaskForward) = ffn_s_CreateParcelID(Mid(as_ORMapTaxlotString, 25, 5), Mid(sMaskToApply, iPosCharMaskForward, lMaskTokenCount))
                            bProcessedParcelID = True
                        End If
                    Case 38 '"&" 'Using these characters in mask will strip leading zeros from parcel id
                        If Not bProcessedParcelID Then
                            sTemp = ffn_s_CreateParcelID(Mid(as_ORMapTaxlotString, 25, 5), Mid(sMaskToApply, iPosCharMaskForward, lMaskTokenCount))
                            Mid(sFormattedString, iPosCharMaskForward) = ffn_s_StripLeadingZeros(sTemp)
                            bProcessedParcelID = True
                        End If
                        'QUARTER and QUARTER QUARTER
                    Case 81 '"Q"
                        If StrComp(sPrevCharInMaskArray, "Q", CompareMethod.Text) = 0 Then ' Quarter Quarter
                            If bHasAlphaQtrQtr Then
                                Mid(sFormattedString, iPosCharMaskForward, 1) = Mid(as_ORMapTaxlotString, 18, 1)
                            Else
                                sCurrORMapNumValue = UCase(Mid(as_ORMapTaxlotString, 18, 1))
                                If sCurrORMapNumValue Like "[A-J]" Then
                                    Mid(sFormattedString, iPosCharMaskForward, 1) = sCurrORMapNumValue
                                Else
                                    If iCountyCode <> 3 And iCountyCode <> 22 Then 'Leave the space
                                        Mid(sFormattedString, iPosCharMaskForward, 1) = Chr(48) 'ZERO
                                    End If
                                End If
                            End If
                        Else ' Quarter
                            If bHasAlphaQtr Then
                                Mid(sFormattedString, iPosCharMaskForward, 1) = Mid(as_ORMapTaxlotString, 17, 1)
                            Else
                                sCurrORMapNumValue = UCase(Mid(as_ORMapTaxlotString, 17, 1))
                                If sCurrORMapNumValue Like "[A-J]" Then
                                    Mid(sFormattedString, iPosCharMaskForward, 1) = sCurrORMapNumValue
                                Else
                                    If iCountyCode <> 3 And iCountyCode <> 22 Then 'leave the space
                                        Mid(sFormattedString, iPosCharMaskForward, 1) = Chr(48) 'ZERO
                                    End If
                                End If
                            End If
                        End If
                        'Range
                    Case 82 '"R"
                        If StrComp(sPrevCharInMaskArray, "R", CompareMethod.Text) <> 0 Then
                            If lMaskTokenCount > 1 Then
                                Mid(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid(as_ORMapTaxlotString, 9, lMaskTokenCount)
                            Else
                                Mid(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid(as_ORMapTaxlotString, 10, lMaskTokenCount)
                            End If
                        End If
                        'SECTION
                    Case 83 '"S"
                        If StrComp(sPrevCharInMaskArray, "S", CompareMethod.Text) = 0 Then 'SECOND pos
                            Mid(sFormattedString, iPosCharMaskForward, 1) = Mid(as_ORMapTaxlotString, 16, lMaskTokenCount)
                        Else 'FIRST POS
                            Mid(sFormattedString, iPosCharMaskForward, 1) = Mid(as_ORMapTaxlotString, 15, lMaskTokenCount)
                        End If
                        'Township
                    Case 84 '"T"
                        If StrComp(sPrevCharInMaskArray, "T", CompareMethod.Text) <> 0 Then
                            If lMaskTokenCount > 1 Then
                                Mid(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid(as_ORMapTaxlotString, 3, lMaskTokenCount)
                            Else
                                Mid(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid(as_ORMapTaxlotString, 4, lMaskTokenCount)
                            End If
                        End If
                        ' Fractional parts of a township
                    Case 80 '"P"
                        If StrComp(sPrevCharInMaskArray, "T", CompareMethod.Text) = 0 Then
                            If Not bProcessedRangeFractional Then
                                Mid(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid(as_ORMapTaxlotString, 11, lMaskTokenCount)
                                bProcessedRangeFractional = True
                            End If
                        ElseIf StrComp(sPrevCharInMaskArray, "R", CompareMethod.Text) = 0 Then
                            If Not bProcessedTownFractional Then
                                Mid(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Mid(as_ORMapTaxlotString, 5, lMaskTokenCount)
                                bProcessedTownFractional = True
                            End If
                        End If
                    Case 94 '^ special case for Clackamas county
                        If StrComp(sPrevCharInMaskArray, "R", CompareMethod.Text) = 0 Then
                            If bHasRangePart Then
                                Mid(sFormattedString, iPosCharMaskForward, 1) = Chr(81) 'Q
                            End If
                        ElseIf StrComp(sPrevCharInMaskArray, "T", CompareMethod.Text) = 0 Then  'fractional part of township
                            If bHasTownPart Then
                                Mid(sFormattedString, iPosCharMaskForward, lMaskTokenCount) = Chr(81)
                            End If
                        End If
                End Select

                sPrevCharInMaskArray = sArr_MaskValues(i)

            Next i

            ' Returns the value of the function
            Return sFormattedString
            'gfn_s_CreateMapTaxlotValue = Trim(sFormattedString)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return String.Empty
        End Try

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="mapTaxlotIDValue"></param>
    ''' <param name="formatString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateMapTaxlotValue(ByVal mapTaxlotIDValue As String, ByVal formatString As String) As String

        If mapTaxlotIDValue Is Nothing OrElse mapTaxlotIDValue.Length = 0 Then
            Throw New ArgumentNullException("mapTaxlotIDValue")
        End If
        If formatString Is Nothing OrElse formatString.Length = 0 Then
            Throw New ArgumentNullException("formatString")
        End If
        If mapTaxlotIDValue.Length < 29 Then
            Throw New Exception("Invalid arguement length for mapTaxlotValue", Nothing)
        End If
        Try
            Dim countyCode As Short
            countyCode = CShort(mapTaxlotIDValue.Substring(0, 2))

            Dim hasTownPart As Boolean
            Dim hasRangePart As Boolean
            Dim hasAlphaQtr As Boolean = False
            Dim hasAlphaQtrQtr As Boolean = False

            'flag for half township
            hasTownPart = (Convert.ToDouble(mapTaxlotIDValue.Substring(4, 3)) > 0)
            hasRangePart = (Convert.ToDouble(mapTaxlotIDValue.Substring(10, 3)) > 0)

            'flags for section quarters
            Select Case countyCode
                Case 1 To 19, 21 To 36
                    If Not IsNumeric(mapTaxlotIDValue.Substring(16, 1)) Then
                        hasAlphaQtr = True
                    End If
                    If Not IsNumeric(mapTaxlotIDValue.Substring(17, 1)) Then
                        hasAlphaQtrQtr = True
                    End If
            End Select

            'We must adjust the mask for clackamas county if there are no half ranges in the current string
            If formatString.IndexOf("^"c) > 0 Then
                If hasRangePart = False Then
                    formatString = formatString.Remove(formatString.IndexOf("^"c), 1)
                Else
                    'if there is a range part the letter Q will be  placed in the position where D sits
                    formatString = formatString.Remove(formatString.IndexOf("D"c), 1)
                End If
            End If
            'copy of the formatstring
            Dim maskValues As New StringBuilder(formatString.ToUpper)
            ' Create a string of spaces to place our results in. This helps a speed up string manipulation a little.
            Dim formattedResult As New StringBuilder(New String(" ", formatString.Length), formatString.Length)

            Dim positionInMask As Integer
            Dim characterCode As Integer
            Dim tokenCount As Integer
            Dim previousCharInMask As Char
            Dim hasProcessedParcelId As Boolean = False
            Dim hasProcessedTownFractional As Boolean = False
            Dim hasProcessedRangeFractional As Boolean = False

            For charIdx As Integer = 0 To maskValues.Length - 1
                positionInMask = formatString.IndexOf(maskValues.Chars(charIdx).ToString, charIdx, StringComparison.CurrentCultureIgnoreCase)
                characterCode = Convert.ToInt32(maskValues.Chars(charIdx))
                ' Returns how many of these characters appear in the mask
                Dim c As Char
                For Each c In formatString
                    If c.Equals(maskValues.Chars(charIdx)) Then
                        tokenCount += 1
                    End If
                Next c

                Select Case characterCode
                    Case 68 'D
                        If String.CompareOrdinal(previousCharInMask, "^") = 0 Then
                            If String.CompareOrdinal(maskValues.Chars(positionInMask - 2), "T") = 0 Then 'township
                                formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(7, 1))
                            ElseIf String.CompareOrdinal(maskValues.Chars(positionInMask - 2), "R") = 0 Then 'range
                                formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(13, 1))
                            End If
                        Else
                            If String.CompareOrdinal(previousCharInMask, "T") = 0 Then 'township
                                formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(7, 1))
                            ElseIf String.CompareOrdinal(previousCharInMask, "R") = 0 Then 'range
                                formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(13, 1))
                            End If
                        End If
                    Case 64 '@
                        'Formats for the parcel id
                        If Not hasProcessedParcelId Then
                            'since we are at the end of the string use Insert
                            formattedResult.Insert(positionInMask, mapTaxlotIDValue.Substring(24, 5)) 'TODO: JWM verify
                            hasProcessedParcelId = True
                        End If
                    Case 38 '& Using these characters in mask will strip leading zeros from parcel id
                        If Not hasProcessedParcelId Then '
                            'since we are at the end of the string use Insert
                            Dim s As String = New String(mapTaxlotIDValue.Substring(24, 5))
                            formattedResult.Insert(positionInMask, StripLeadingZeros(s))
                            hasProcessedParcelId = True
                        End If
                    Case 81 'Q
                        If String.CompareOrdinal(previousCharInMask, "Q") = 0 Then 'qtr qtr
                            If hasAlphaQtrQtr Then
                                formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(17, 1))
                            Else 'it is not alphabetical could be a number or a space
                                Dim currentORMAPNumValue As String
                                currentORMAPNumValue = mapTaxlotIDValue.Substring(17, 1).ToUpper
                                If currentORMAPNumValue Like "[A-J]" Then 'TODO: why am i doing this
                                    formattedResult.Chars(positionInMask) = CChar(currentORMAPNumValue)
                                Else
                                    If countyCode <> 3 And countyCode <> 22 Then 'Clackamas County wants the space/blank value left in the string NO ZEROES PLEASE
                                        formattedResult.Chars(positionInMask) = "0"c
                                    End If
                                End If
                            End If
                        Else 'qtr
                            If hasAlphaQtr Then
                                formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(16, 1))
                            Else 'it is not alphabetical could be a number or a space
                                Dim currentORMAPNum As String
                                currentORMAPNum = mapTaxlotIDValue.Substring(16, 1)
                                If currentORMAPNum Like "[A-J]" Then 'TODO: why am i doing this
                                    formattedResult.Chars(positionInMask) = CChar(currentORMAPNum)
                                Else
                                    If countyCode <> 3 And countyCode <> 22 Then
                                        formattedResult.Chars(positionInMask) = "0"c
                                    End If
                                End If
                            End If
                        End If

                    Case 82 'Range
                        If String.CompareOrdinal(previousCharInMask, "R") <> 0 Then
                            If tokenCount > 1 Then
                                formattedResult.Insert(positionInMask, mapTaxlotIDValue.Substring(8, tokenCount))
                            Else
                                formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(9, 1))
                            End If
                        End If
                    Case 83 'S section
                        If String.CompareOrdinal(previousCharInMask, "S") = 0 Then 'second position
                            formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(15, 1)) 'TODO: JWM verify
                        Else 'first position
                            formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(14, 1))
                        End If

                    Case 84 'T township
                        If String.CompareOrdinal(previousCharInMask, "T") <> 0 Then
                            If tokenCount > 1 Then
                                formattedResult.Insert(positionInMask, mapTaxlotIDValue.Substring(2, tokenCount))
                            Else
                                formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(3, 1))
                            End If
                        End If

                    Case 80 'P fractional parts
                        If String.CompareOrdinal(previousCharInMask, "T") = 0 Then
                            If Not hasProcessedRangeFractional Then
                                formattedResult.Insert(positionInMask, mapTaxlotIDValue.Substring(10, tokenCount))
                                hasProcessedRangeFractional = True
                            ElseIf String.CompareOrdinal(previousCharInMask, "R") = 0 Then
                                If Not hasProcessedTownFractional Then
                                    formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(4, 1))
                                    hasProcessedTownFractional = True
                                End If
                            End If
                        End If

                    Case 94 '^ special case for clackamas county
                        If String.CompareOrdinal(previousCharInMask, "R") = 0 Then
                            If hasRangePart Then
                                formattedResult.Chars(positionInMask) = "Q"c
                            End If
                        ElseIf String.CompareOrdinal(previousCharInMask, "T") = 0 Then 'fractional part of township
                            If hasTownPart Then
                                formattedResult.Chars(positionInMask) = "Q"c
                            End If
                        End If
                End Select
                previousCharInMask = maskValues.Chars(charIdx)
                tokenCount = 0
            Next charIdx

            Return formattedResult.ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return String.Empty
        End Try

    End Function

    ''' <summary>
    ''' Create a parcel ID from a mask.
    ''' </summary>
    ''' <param name="valueToMask"></param>
    ''' <param name="maskToApply"></param>
    ''' <returns> If a value is passed in that is not numeric then just pass it straight through else return a parcel id with or without leading zeros</returns>
    ''' <remarks>I use the Format function with user-defined string formats which consist of either all (@) characters or all ampersands</remarks>
    Private Shared Function CreateParcelID(ByVal valueToMask As String, ByVal maskToApply As String) As String

        If valueToMask.Length = 0 Then
            Throw New ArgumentNullException("valueToMask")
        End If
        If maskToApply.Length = 0 Then
            Throw New ArgumentNullException("maskToApply")
        End If

        If maskToApply.Contains("&") Then
            maskToApply = maskToApply.Replace("&", "#")
        ElseIf maskToApply.Contains("@") Then
            maskToApply = maskToApply.Replace("@", "0")
        End If

        Dim formatItem As New String("{0," & valueToMask.Length & ":" & maskToApply & "}")

        If IsNumeric(valueToMask) Then


            Return new String(string.Format(formatItem,valueToMask))
        Else
            Return valueToMask
        End If

    End Function

    Private Shared Function StripLeadingZeros(ByRef stringToParse As String) As String
        Try
            Dim sb As New StringBuilder(stringToParse)

            For charIdx As Integer = 0 To sb.Length
                If Char.GetNumericValue(sb.Chars(charIdx)) = 0 Then
                    sb = sb.Replace(sb.Chars(charIdx), " "c, charIdx, 1)
                Else
                    Exit For
                End If
            Next charIdx

            Return sb.ToString  ' do not trim off leading spaces
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
End Class