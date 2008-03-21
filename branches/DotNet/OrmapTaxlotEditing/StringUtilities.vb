#Region "Copyright 2008 ORMAP Tech Group"
' File:  StringUtilities.vb
'
' Original Author:  OPET.NET Migration Team (Shad Campbell, James Moore, 
'                   Nick Seigal)
'
' Date Created:  20080221
'
' Copyright Holder:  ORMAP Tech Group  
' Contact Info:  ORMAP Tech Group (a.k.a. opet developers) may be reached at 
' opet-developers@lists.sourceforge.net
'
' This file is part of the ORMAP Taxlot Editing Toolbar.
'
' ORMAP Taxlot Editing Toolbar is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License as published by 
' the Free Software Foundation; either version 3 of the License, or (at your 
' option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT 
' ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or 
' FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License located
' in the COPYING.txt file for more details.
'
' You should have received a copy of the GNU General Public License along
' with the ORMAP Taxlot Editing Toolbar; if not, write to the Free Software 
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
#End Region

#Region "Subversion Keyword expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision$
'Date of Last Change: $Date$
#End Region

#Region "Imported Namespaces"
Imports System.Windows.Forms
Imports System.Text
#End Region

#Region "Class Declaration"
Public NotInheritable Class StringUtilities

#Region "Custom Class Members"

#Region "Public Members"

    ''' <summary>
    ''' Adds leading zeros if necessary
    ''' </summary>
    ''' <param name="currentString">The string to pad with zeros</param>
    ''' <param name="width">The final length of the string</param>
    ''' <returns>A string of length width characters</returns>
    ''' <remarks>Creates a string of width characters padded on the left with zeros</remarks>
    Public Shared Function AddLeadingZeros(ByVal currentString As String, ByVal width As Integer) As String
        Try
            If currentString.Length < width Then
                Return currentString.PadLeft(width - currentString.Length, "0"c)
            Else
                Return currentString
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return String.Empty
        End Try
    End Function

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
                            Else
                                Dim currentORMAPNumValue As String
                                currentORMAPNumValue = mapTaxlotIDValue.Substring(17, 1).ToUpper
                                If currentORMAPNumValue Like "[A-D]" Then
                                    Select Case currentORMAPNumValue
                                        Case "A"
                                            formattedResult.Chars(positionInMask) = "A"c
                                        Case "B"
                                            formattedResult.Chars(positionInMask) = "B"c
                                        Case "C"
                                            formattedResult.Chars(positionInMask) = "C"c
                                        Case "D"
                                            formattedResult.Chars(positionInMask) = "D"c
                                    End Select
                                Else
                                    If countyCode <> 3 Then 'Clackamas County wants the space/blank value left in the string NO ZEROES PLEASE
                                        formattedResult.Chars(positionInMask) = "0"c
                                    End If
                                End If
                            End If
                        Else 'qtr
                            If hasAlphaQtr Then
                                formattedResult.Chars(positionInMask) = CChar(mapTaxlotIDValue.Substring(16, 1))
                            Else
                                Dim currentORMAPNum As String
                                currentORMAPNum = mapTaxlotIDValue.Substring(16, 1)
                                If currentORMAPNum Like "[A-D]" Then
                                    Select Case currentORMAPNum
                                        Case "A"
                                            formattedResult.Chars(positionInMask) = "A"c
                                        Case "B"
                                            formattedResult.Chars(positionInMask) = "B"c
                                        Case "C"
                                            formattedResult.Chars(positionInMask) = "C"c
                                        Case "D"
                                            formattedResult.Chars(positionInMask) = "D"c
                                    End Select
                                Else
                                    If countyCode <> 3 Then
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
            Dim returnValue As New String(formattedResult.ToString)
            Return returnValue
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return String.Empty
        End Try

    End Function

    ''' <summary>
    ''' Isolate elements of a string
    ''' </summary>
    ''' <param name="theWholeString">The string to isolate the substring from</param>
    ''' <param name="lowPart"></param>
    ''' <param name="highPart"></param>
    ''' <returns>A string that is a substring of theWholeString.</returns>
    ''' <remarks></remarks>
    Public Shared Function ExtractString(ByVal theWholeString As String, ByVal lowPart As Integer, ByVal highPart As Integer) As String
        Try 'HACK: JWM Probably can be replaced with String.Substring()
            If lowPart <= highPart Then
                'Return Mid(theWholeString, lowPart, highPart - lowPart + 1)
                Return theWholeString.Substring(lowPart, highPart - lowPart + 1)
            Else
                Return String.Empty
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Remove two characters (the county code) from the right end of the OrmapMapNumber.
    ''' </summary>
    ''' <param name="theOrmapMapNumber">The Ormap Map Number string.</param>
    ''' <returns>A string that is a substring of the input.</returns>
    ''' <remarks>For the purpose of populating OrmapTaxlot.</remarks>
    Public Shared Function OrmapMapNumberNoCountyCodeSuffix(ByVal theOrmapMapNumber As String) As String
        Try
            ' Remove two characters (the county code) from the right end of 
            ' the OrmapMapNumber.
            Return Left(theOrmapMapNumber, 20)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return String.Empty
        End Try
    End Function

#End Region

#Region "Private Members"
    ''' <summary>
    ''' Create a parcel ID from a mask.
    ''' </summary>
    ''' <param name="valueToMask"></param>
    ''' <param name="maskToApply"></param>
    ''' <returns> If a value is passed in that is not numeric then just pass it straight through else return a parcel id with or without leading zeros</returns>
    ''' <remarks>I use the Format function with user-defined string formats which consist of either all (@) characters or all ampersands</remarks>
    Private Shared Function createParcelID(ByVal valueToMask As String, ByVal maskToApply As String) As String
        Dim sb As StringBuilder
        If valueToMask.Length = 0 OrElse maskToApply.Length = 0 Then
            Return String.Empty
        End If

        If IsNumeric(valueToMask) Then
            sb = New StringBuilder(Format(valueToMask, maskToApply), maskToApply.Length)

        Else
            sb = New StringBuilder(valueToMask, maskToApply.Length)
        End If
        Return sb.ToString

    End Function

    Private Shared Function stripLeadingZeros(ByRef stringToParse As String) As String
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
			MessageBox.Show(ex.Message)
            Return String.Empty
        End Try
    End Function

#End Region

#End Region

End Class
#End Region
