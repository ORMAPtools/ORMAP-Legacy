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

    Public Shared Function CreateMapTaxlotValue(ByVal mapTaxlotIDValue As String, ByVal formatString As String) As String
        Return String.Empty 'TODO:jwm flesh this out
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
    Private Shared Function CreateParcelID(ByVal valueToMask As String, ByVal maskToApply As String) As String
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

    Private Shared Function StripLeadingZeros(ByRef stringToParse As String) As String
        Dim inputCharCount As Integer
        Dim aChar As Char
        Dim sTemp As StringBuilder

        inputCharCount = stringToParse.Length
        'create string of same length
        sTemp = New StringBuilder(" ", inputCharCount)

        'TODO JWM test this function

        For counter As Integer = 1 To inputCharCount
            'aChar = Mid(stringToParse, counter, 1)
            aChar = stringToParse.Chars(counter)
            If stringToParse.Contains(aChar) Then
                sTemp.Insert(counter, stringToParse.Substring(counter, inputCharCount - counter))
                Exit For
                'If InStr(1, "0", aChar, CompareMethod.Text) < 1 Then 'go past all leading zeros
                '    Mid(sTemp, counter) = Mid(stringToParse, counter) 'get all remaing chars
                '    Exit For 'and exit
                'End If
            End If
        Next counter
        Return sTemp.ToString  ' do not trim off leading spaces
    End Function

#End Region

#End Region

End Class
#End Region
