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
'SCC revision number: $Revision:$
'Date of Last Change: $Date:$
#End Region

#Region "Imported namespace statements"
Imports System.Windows.Forms
#End Region
#Region "Class Declaration"
Public NotInheritable Class StringUtilities
#Region "Custom Class Members"
#Region "Public Members"
    ''' <summary>
    ''' Isolate elements of a string
    ''' </summary>
    ''' <param name="theWholeString">The string to isolate the substring from</param>
    ''' <param name="lowPart"></param>
    ''' <param name="highPart"></param>
    ''' <returns>A string that is a substring of theWholeString.</returns>
    ''' <remarks></remarks>
    Public Shared Function ExtractString(ByVal theWholeString As String, ByVal lowPart As Integer, ByVal highPart As Integer) As String
        Try
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

#End Region

#Region "Private Members"

#End Region
#End Region
End Class
#End Region
