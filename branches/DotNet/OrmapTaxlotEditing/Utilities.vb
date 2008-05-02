#Region "Copyright 2008 ORMAP Tech Group"

' File:  Utilities.vb
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

#Region "Subversion Keyword Expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision$
'Date of Last Change: $Date$
#End Region

#Region "Imported Namespaces"
Imports System.Windows.Forms
Imports System.IO
#End Region

#Region "Class Declaration"
Public NotInheritable Class Utilities

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    ''' <summary>
    ''' Private empty constructor to prevent instantiation.
    ''' </summary>
    ''' <remarks>This class follows the singleton pattern and thus has a 
    ''' private constructor and all shared members. Instances of types 
    ''' that define only shared members do not need to be created, so no
    ''' constructor should be needed. However, many compilers will 
    ''' automatically add a public default constructor if no constructor 
    ''' is specified. To prevent this an empty private constructor is 
    ''' added.</remarks>
    Private Sub New()
    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Public Members"
    Public Const FieldNotFoundIndex As Integer = -1

    Friend Enum EsriMouseButtons
        Left = 1
        Right = 2
        Middle = 4
    End Enum

    ''' <summary>
    ''' Stores the current computer user name.
    ''' </summary>
    ''' <value></value>
    ''' <returns>A username string.</returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property UserName() As String
        Get
            ' Note: ALL Since this a dll, My.User.InitializeWithWindowsUser()
            ' is called in EditorExtension.Startup to set this value
            If TypeOf My.User.CurrentPrincipal Is _
                    Security.Principal.WindowsPrincipal Then
                '[The application is using Windows authentication...]
                '[The name format is "DOMAIN\USERNAME"...]
                ' Parse out USERNAME from DOMAIN\USERNAME pair
                Dim parts() As String = Split(My.User.Name, "\")
                Dim name As String = parts(1)
                Return name
            Else
                ' The application is using custom authentication.
                Return My.User.Name
            End If
        End Get
    End Property

    ''' <summary>
    ''' Determine file existence
    ''' </summary>
    ''' <param name="path">A string that represents the file to check</param>
    ''' <returns>True or False</returns>
    ''' <remarks></remarks>
    Public Shared Function FileExists(ByVal path As String) As Boolean
        Try
            If path Is Nothing OrElse path.Length = 0 Then
                Throw New ArgumentNullException("path")
            End If
            Dim fInfo As New FileInfo(path)
            If fInfo.Exists Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Opens a document with its associated application.
    ''' </summary>
    ''' <param name="path">Fully qualified path to document (including file name).</param>
    ''' <remarks></remarks>
    Public Shared Sub StartDoc(ByVal path As String)
        Try
            System.Diagnostics.Process.Start(path)
        Catch fex As FileNotFoundException
            MessageBox.Show("File not Found", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return
        End Try
    End Sub

#End Region

#Region "Private Members (none)"
#End Region

#End Region

End Class
#End Region
