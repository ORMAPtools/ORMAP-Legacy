#Region "Copyright 2008 ORMAP Tech Group"

' File:  OrmapSettings.vb
'
' Original Author:  OPET.NET Migration Team (Shad Campbell, James Moore, 
'                   Nick Seigal)
'
' Date Created:  January 8, 2008
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

' TODO: Implement this class and move code here from the form so as to seperate
' the UI and business logic layers.

Imports System.Runtime.InteropServices

<ComVisible(False)> _
Public NotInheritable Class OrmapSettings

#Region "Class-Level Constants And Enumerations"
    ' None
#End Region

#Region "Built-In Class Members (Properties, Methods, Events, Event Handlers, Delegates, Etc.)"

#Region "Constructors"

    ''' <summary>
    ''' OrmapSettings constructor.
    ''' </summary>
    ''' <remarks>This class follows the singleton pattern and thus has a 
    ''' private constructor and all shared members.</remarks>
    Private Sub New()
    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"
    ' None
#End Region

#Region "Properties"

    Private WithEvents _ormapSettingsForm As OrmapSettingsForm

    Friend ReadOnly Property OrmapSettingsForm() As OrmapSettingsForm
        Get
            Return _ormapSettingsForm
        End Get
    End Property

    Private Sub SetOrmapSettingsForm(ByVal value As OrmapSettingsForm)
        ' TODO: Add validation code?
        _ormapSettingsForm = value
    End Sub

#End Region

#Region "Event Handlers"
    ' TODO: Add event handlers for the form controls.
#End Region

#Region "Methods"
    ' None
#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"
    ' None
#End Region

#Region "Methods"
    ' None
#End Region

#End Region

#Region "Implemented Interface Members"
    ' None
#End Region

#Region "Other Members"
    ' None
#End Region

End Class
