#Region "Copyright 2008 ORMAP Tech Group"

' File:  CatalogFileDlg.vb
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
'Tag for this file: $Name:$
'SCC revision number: $Revision:$
'Date of Last Change: $Date:$
#End Region

#Region "Imported Namespaces"
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports System.Windows.Forms

#End Region

#Region "Class declaration"
''' <summary>
''' Programmatically expose the ArcCatalog file dialog as one integral unit for simple access and use.
''' </summary>
''' <remarks></remarks>
Public Class CatalogFileDialog
    'TODO: JWM Would like to remove the dependency on VB.Collection object
#Region "Class level fields"
    Private _theGxDialog As IGxDialog
    Private _colSelection As Collection
#End Region

    Public Sub New()
        MyBase.New()
        _theGxDialog = New GxDialog
        _theGxDialog.RememberLocation = True
        _theGxDialog.AllowMultiSelect = False

    End Sub

    Protected Overrides Sub Finalize()
        _theGxDialog = Nothing
        MyBase.Finalize()
    End Sub

#Region "Built-in class members"

    Public Property Name() As String
        Get
            Name = _theGxDialog.Name
        End Get
        Set(ByVal value As String)
            _theGxDialog.Name = value
        End Set
    End Property

    ''' <summary>
    ''' The file path present when the user specifies a file to open
    ''' or a file name to save in Open/Save As dialog boxes
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FinalLocation() As String
        Get
            FinalLocation = _theGxDialog.FinalLocation.FullName
        End Get
    End Property

#End Region

#Region "Custom class members"

    Public Const NoSelectedElementIndex As Integer = -1

    ''' <summary>
    ''' Retrieve an item selected from a file dialog box
    ''' </summary>
    ''' <param name="selection"></param>
    ''' <returns></returns>
    ''' <remarks>Return the nth selected element from the most recent file save/open dialog request</remarks>
    Public Function SelectedObject(Optional ByVal selection As Integer = NoSelectedElementIndex) As Object
        If selection > _colSelection.Count Then
            Return String.Empty
        ElseIf selection = NoSelectedElementIndex Then
            Return _colSelection
        Else
            Return _colSelection.Item(selection)
        End If
    End Function

    Public Sub SetAllowMultiSelect(ByVal allow As Boolean)
        _theGxDialog.AllowMultiSelect = allow
    End Sub

    Public Sub SetButtonCaption(ByVal caption As String)
        _theGxDialog.ButtonCaption = caption
    End Sub


    ''' <summary>
    ''' Initial file path for either open or save dialog boxes
    ''' </summary>
    ''' <param name="pointer"></param>
    ''' <remarks></remarks>
    Public Sub SetStartingLocation(ByVal pointer As System.IntPtr)
        _theGxDialog.StartingLocation = pointer
    End Sub

    Public Sub SetTitle(ByVal title As String)
        _theGxDialog.Title = title
    End Sub

    ''' <summary>
    ''' Simplify adding a filter to the file dialog box
    ''' </summary>
    ''' <param name="filter">An ArcObject defined object filter</param>
    ''' <param name="isDefault">Indicates if the filter should be the default filter</param>
    ''' <param name="resetAll">Indicates whether or not all of the current filters should be cleared</param>
    ''' <returns></returns>
    ''' <remarks>Adds a ESRI ArcCatalog defined filter to a file dialog box filter list</remarks>
    Public Function SetFilter(ByRef filter As IGxObjectFilter, Optional ByRef isDefault As Boolean = False, Optional ByRef resetAll As Boolean = True) As Boolean
        Try
            Dim filters As IGxObjectFilterCollection
            filters = DirectCast(_theGxDialog, IGxObjectFilterCollection)
            If resetAll Then
                filters.RemoveAllFilters()
            End If
            filters.AddFilter(filter, isDefault)
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Show the ArcCatalog file open dialog box
    ''' </summary>
    ''' <returns>  A collection of names of objects that have been selected by the user from the dialog box</returns>
    ''' <remarks></remarks>
    Public Function ShowOpen() As Collection
        Try
            Dim selection As IEnumGxObject
            Dim thisSelectedObject As IGxObject
            selection = New GxObjectArray

            _colSelection = New Collection

            If Not _theGxDialog.DoModalOpen(EditorExtension.Application.hWnd, selection) Then
                'need to return a empty collection
                Return New Collection
            End If

            selection.Reset()
            thisSelectedObject = selection.Next
            Do While Not thisSelectedObject Is Nothing
                _colSelection.Add(thisSelectedObject.FullName)
                thisSelectedObject = selection.Next
            Loop
            Return _colSelection
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return New Collection
        End Try
    End Function

    ''' <summary>
    ''' Show the ArcCatalog file save dialog box
    ''' </summary>
    ''' <returns>A collection holding the full path that is a concatenation of the final path and the specified name</returns>
    ''' <remarks></remarks>
    Public Function ShowSave() As Collection
        Try
            If Not _theGxDialog.DoModalSave(EditorExtension.Application.hWnd) Then
                ' Return an empty collection
                Return New Collection
            End If
            Dim selectedObject As IGxObject
            selectedObject = _theGxDialog.FinalLocation
            _colSelection.Add(String.Concat(selectedObject.FullName, "\", _theGxDialog.Name))
            Return _colSelection
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return New Collection
        End Try
    End Function

#End Region
End Class
#End Region