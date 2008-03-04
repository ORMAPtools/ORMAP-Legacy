#Region "Copyright 2008 ORMAP Tech Group"

' File:  TaxlotAssignment.vb
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
'SCC revision number: $Revision$
'Date of Last Change: $Date$
#End Region

Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework

<ComVisible(True)> _
<ComClass(TaxlotAssignment.ClassId, TaxlotAssignment.InterfaceId, TaxlotAssignment.EventsId), _
ProgId("ORMAPTaxlotEditing.TaxlotAssignment")> _
Public NotInheritable Class TaxlotAssignment
    Inherits BaseTool

#Region "Class-Level Constants And Enumerations"

    ' Taxlot number type constants
    Private Const taxlotNumberTypeTaxlot As String = "TAXLOT" 'normal taxlot number
    Private Const taxlotNumberTypeRoads As String = "ROADS"
    Private Const taxlotNumberTypeWater As String = "WATER"
    Private Const taxlotNumberTypeRails As String = "RAILS"
    Private Const taxlotNumberTypeNontaxlot As String = "NONTL"

    Private Const defaultCommand As String = "esriArcMapUI.SelectTool"

#End Region

#Region "Built-In Class Members (Properties, Methods, Events, Event Handlers, Delegates, Etc.)"

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' Define protected instance field values for the public properties.
        MyBase.m_category = "OrmapToolbar"  'localizable text 
        MyBase.m_caption = "TaxlotAssignment"   'localizable text 
        MyBase.m_message = "Populate values in the Taxlots feature class based on a starting value and an increment value."   'localizable text 
        MyBase.m_toolTip = "Assign Taxlots" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_TaxlotAssignment"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            ' Set the bitmap based on the name of the class.
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As ArgumentException
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

        Try
            ' Set the (enabled) cursor based on the name of the class.
            Dim cursorResourceName As String = Me.GetType().Name + ".cur"
            MyBase.m_cursor = New Cursor(Me.GetType(), cursorResourceName)
        Catch ex As ArgumentException
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Cursor")
        End Try

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication

    Private _disabledCursor As System.Drawing.Image

#End Region

#Region "Properties"

    Private _incrementNumber As Integer

    Public ReadOnly Property IncrementNumber() As Integer
        Get
            _incrementNumber = CInt(PartnerTaxlotAssignmentForm.uxType.SelectedItem.ToString)
            Return _incrementNumber
        End Get
    End Property

    Private _taxlotType As String

    Public ReadOnly Property TaxlotType() As String
        Get
            _taxlotType = PartnerTaxlotAssignmentForm.uxType.SelectedItem.ToString
            Return _taxlotType
        End Get
    End Property

    Private _numberStartingFrom As Integer

    Public ReadOnly Property NumberStartingFrom() As Integer
        Get
            _numberStartingFrom = CInt(PartnerTaxlotAssignmentForm.uxStartingFrom.Text)
            Return _numberStartingFrom
        End Get
    End Property

    Private WithEvents _partnerTaxlotAssignmentForm As TaxlotAssignmentForm  ' TODO: NIS Is WithEvents needed here?

    Friend ReadOnly Property PartnerTaxlotAssignmentForm() As TaxlotAssignmentForm
        Get
            If _partnerTaxlotAssignmentForm Is Nothing OrElse _partnerTaxlotAssignmentForm.IsDisposed Then
                SetPartnerTaxlotAssignmentForm(New TaxlotAssignmentForm())
            End If
            Return _partnerTaxlotAssignmentForm
        End Get
    End Property

    Private Sub SetPartnerTaxlotAssignmentForm(ByVal value As TaxlotAssignmentForm)
        ' TODO: NIS Add validation code?
        If value IsNot Nothing Then
            _partnerTaxlotAssignmentForm = value
            ' Wire up partner form events.
            AddHandler _partnerTaxlotAssignmentForm.Load, AddressOf PartnerTaxlotAssignmentForm_Load
            AddHandler _partnerTaxlotAssignmentForm.uxHelp.Click, AddressOf uxHelp_Click
            AddHandler _partnerTaxlotAssignmentForm.uxType.SelectedValueChanged, AddressOf uxType_SelectedValueChanged
        End If
    End Sub

#End Region

#Region "Event Handlers"

    Private Sub PartnerTaxlotAssignmentForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles PartnerTaxlotAssignmentForm.Load

        With PartnerTaxlotAssignmentForm

            'Populate multi-value controls
            .uxType.Items.Add(taxlotNumberTypeTaxlot)
            .uxType.Items.Add(taxlotNumberTypeRoads)
            .uxType.Items.Add(taxlotNumberTypeWater)
            .uxType.Items.Add(taxlotNumberTypeRails)
            .uxType.Items.Add(taxlotNumberTypeNontaxlot)

            ' Set control defaults
            .uxType.Text = taxlotNumberTypeTaxlot
            .uxIncrementByNone.Checked = True
            .uxStartingFrom.Text = "0"

            ' Enable the numbering settings controls by enabling the group
            .uxTaxlotNumberingOptions.Enabled = True
            'With .uxStartingFrom
            '    '.BackColor = System.Drawing.Color.White
            '    .Enabled = True
            'End With
            '.uxIncrementByNone.Enabled = True
            '.uxIncrementBy1.Enabled = True
            '.uxIncrementBy10.Enabled = True
            '.uxIncrementBy100.Enabled = True
            '.uxIncrementBy1000.Enabled = True
        End With


    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles TaxlotAssignmentForm.uxHelp.Click
        ' TODO: NIS Could be replaced with new help mechanism.
        ' Open a custom help file.
        ' Note: Requires a specific file in the help subdirectory of the application directory.
        Dim filePath As String
        filePath = My.Application.Info.DirectoryPath & "\help\TaxlotAssignmentHelp.rtf"
        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(filePath) Then
            ' Open help file from the application directory.
            Dim helpForm As New HelpForm
            helpForm.uxContent.LoadFile(filePath, RichTextBoxStreamType.RichText)
            helpForm.Text = "Taxlot Assignment Help"
            helpForm.Show()
        Else
            MessageBox.Show("No help file available in the directory " & My.Application.Info.DirectoryPath & "\help" & ".")
        End If
    End Sub

    Private Sub uxType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles TaxlotAssignmentForm.uxType.SelectedValueChanged
        With PartnerTaxlotAssignmentForm
            If .uxType.SelectedIndex <> -1 Then
                If .uxType.SelectedItem.ToString = taxlotNumberTypeTaxlot Then
                    ' Enable the numbering settings controls by enabling the group
                    .uxTaxlotNumberingOptions.Enabled = True
                Else
                    ' Disable the numbering settings controls by disabling the group
                    .uxTaxlotNumberingOptions.Enabled = False
                End If
            End If
        End With

    End Sub

#End Region

#Region "Methods (none)"
#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"

    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Return MyBase.Enabled AndAlso _
                EditorExtension.Editor IsNot Nothing AndAlso _
                EditorExtension.Editor.EditState = esriEditState.esriStateEditing AndAlso _
                EditorExtension.IsValidWorkspace AndAlso _
                EditorExtension.HasValidLicense AndAlso _
                EditorExtension.CanEditTaxlots
        End Get
    End Property

    Public Overrides ReadOnly Property Cursor() As Integer
        Get
            'Return MyBase.Cursor
            ' Sets the proper cursor
            If PartnerTaxlotAssignmentForm.Visible Then
                If Me.Enabled Then
                    Return MyBase.Cursor
                Else
                    ' TODO: NIS
                End If
            End If
        End Get
    End Property

#End Region

#Region "Methods"

    ''' <summary>
    ''' Occurs when this command is created.
    ''' </summary>
    ''' <param name="hook">A generic <c>Object</c> hook to an instance of the application.</param>
    ''' <remarks>The application hook may not point to an <c>IMxApplication</c> object.</remarks>
    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            _application = DirectCast(hook, IApplication)

            'Disable tool if parent application is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If

            If MyBase.m_enabled Then
                ' Set partner form.
                SetPartnerTaxlotAssignmentForm(New TaxlotAssignmentForm())
            End If
        End If

        ' TODO: Add other initialization code

    End Sub

    Public Overrides Sub OnClick()
        'TODO: NIS Port more of TaxlotAssignment.OnClick implementation
        System.Windows.Forms.MessageBox.Show("Port more of TaxlotAssignment.OnClick implementation")

        ' HACK: NIS Don't know why this form is disposed after first use, but 
        ' this check insures it is available again.
        If PartnerTaxlotAssignmentForm.IsDisposed Then
            SetPartnerTaxlotAssignmentForm(New TaxlotAssignmentForm())
        End If

        ' Show and activate the partner form.
        If PartnerTaxlotAssignmentForm.Visible Then
            PartnerTaxlotAssignmentForm.Activate()
        Else
            PartnerTaxlotAssignmentForm.Show()
        End If

    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: NIS Port TaxlotAssignment.OnMouseDown implementation
        System.Windows.Forms.MessageBox.Show("Port TaxlotAssignment.OnMouseDown implementation")
    End Sub

#End Region

#End Region

#Region "Implemented Interface Members"
#End Region

#Region "Other Members"

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "d091f7ea-0626-4d05-9a7c-533e0961f1cc"
    Public Const InterfaceId As String = "938bb5ab-a827-4731-be72-9db650fb8ed3"
    Public Const EventsId As String = "04e3b3ef-1fb3-46ed-b8d7-11dc19763f32"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub

    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

#End Region

End Class



