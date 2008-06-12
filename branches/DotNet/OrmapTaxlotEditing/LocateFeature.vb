#Region "Copyright 2008 ORMAP Tech Group"

' File:  LocateFeature.vb
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
' modify it under the terms of the Lesser GNU General Public License as 
' published by the Free Software Foundation; either version 3 of the License, 
' or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT 
' ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or 
' FITNESS FOR A PARTICULAR PURPOSE.  See the Lesser GNU General Public License 
' located in the COPYING.LESSER.txt file for more details.
'
' You should have received a copy of the Lesser GNU General Public License 
' along with the ORMAP Taxlot Editing Toolbar; if not, write to the Free 
' Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 
' 02110-1301 USA.

#End Region

#Region "Subversion Keyword Expansion"
'Tag for this file: $Name$
'SCC revision number: $Revision$
'Date of Last Change: $Date$
#End Region

#Region "Imported Namespaces"
Imports System.Drawing
Imports System.Environment
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports OrmapTaxlotEditing.DataMonitor
Imports OrmapTaxlotEditing.SpatialUtilities
Imports OrmapTaxlotEditing.StringUtilities
Imports OrmapTaxlotEditing.Utilities
#End Region

<ComVisible(True)> _
<ComClass(LocateFeature.ClassId, LocateFeature.InterfaceId, LocateFeature.EventsId), _
ProgId("ORMAPTaxlotEditing.LocateFeature")> _
Public NotInheritable Class LocateFeature
    Inherits BaseCommand
    Implements IDisposable

#Region "Class-Level Constants and Enumerations (none)"
#End Region

#Region "Built-In Class Members (Constructors, Etc.)"

#Region "Constructors"

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' Define protected instance field values for the public properties
        MyBase.m_category = "OrmapToolbar"  'localizable text 
        MyBase.m_caption = "LocateFeature"   'localizable text 
        MyBase.m_message = "Locate a Taxlot or Mapindex"   'localizable text 
        MyBase.m_toolTip = "Locate Taxlot or Mapindex" 'localizable text 
        MyBase.m_name = MyBase.m_category & "_LocateFeature"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            ' Set the bitmap based on the name of the class.
            _bitmapResourceName = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), _bitmapResourceName)
        Catch ex As ArgumentException
            EditorExtension.ProcessUnhandledException(ex)
        End Try

    End Sub

#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication
    Private _bitmapResourceName As String

#End Region

#Region "Properties"

    Private WithEvents _partnerLocateFeatureForm As LocateFeatureForm

    Friend ReadOnly Property PartnerLocateFeatureForm() As LocateFeatureForm
        Get
            If _partnerLocateFeatureForm Is Nothing OrElse _partnerLocateFeatureForm.IsDisposed Then
                setPartnerLocateFeatureForm(New LocateFeatureForm())
            End If
            Return _partnerLocateFeatureForm
        End Get
    End Property

    Private Sub setPartnerLocateFeatureForm(ByVal value As LocateFeatureForm)
        If value IsNot Nothing Then
            _partnerLocateFeatureForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerLocateFeatureForm.Load, AddressOf PartnerLocateFeatureForm_Load
            AddHandler _partnerLocateFeatureForm.uxMapNumber.TextChanged, AddressOf uxMapNumber_TextChanged
            AddHandler _partnerLocateFeatureForm.uxFind.Click, AddressOf uxFind_Click
            AddHandler _partnerLocateFeatureForm.uxHelp.Click, AddressOf uxHelp_Click
        Else
            ' Unsubscribe to partner form events.
            RemoveHandler _partnerLocateFeatureForm.Load, AddressOf PartnerLocateFeatureForm_Load
            RemoveHandler _partnerLocateFeatureForm.uxMapNumber.TextChanged, AddressOf uxMapNumber_TextChanged
            RemoveHandler _partnerLocateFeatureForm.uxFind.Click, AddressOf uxFind_Click
            RemoveHandler _partnerLocateFeatureForm.uxHelp.Click, AddressOf uxHelp_Click
        End If
    End Sub


#End Region

#Region "Event Handlers"

    Private Sub PartnerLocateFeatureForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.Load

        With PartnerLocateFeatureForm
            Try
                .UseWaitCursor = True

                '' NOTE: [SC] Calculating a AutoCompleteCustomSource using the List of strings is considerably faster than
                '' adding with the .add method.
                Dim mapNumberStringList As List(Of String) = New List(Of String)
                Dim theMapIndexFClass As IFeatureClass = MapIndexFeatureLayer.FeatureClass
                Dim theQueryFilter As IQueryFilter = New QueryFilter
                theQueryFilter.SubFields = EditorExtension.MapIndexSettings.MapNumberField
                Dim theFeatCursor As IFeatureCursor = theMapIndexFClass.Search(theQueryFilter, True)
                Dim theFeature As IFeature = theFeatCursor.NextFeature
                If theFeature Is Nothing Then Exit Sub
                Dim theFieldIdx As Integer = theFeature.Fields.FindField(EditorExtension.MapIndexSettings.MapNumberField)
                Do Until theFeature Is Nothing
                    Dim theMapNumberVal As String = theFeature.Value(theFieldIdx).ToString
                    If Not mapNumberStringList.Contains(theMapNumberVal) Then
                        mapNumberStringList.Add(theMapNumberVal)
                    End If
                    theFeature = theFeatCursor.NextFeature
                Loop

                ' Populate the combobox from the List object
                .uxMapNumber.AutoCompleteCustomSource.AddRange(mapNumberStringList.ToArray)

                ' Set control defaults
                .uxMapNumber.Text = String.Empty
                .uxTaxlot.Text = String.Empty

            Finally
                .UseWaitCursor = False
            End Try
        End With

    End Sub

    Private Sub uxMapNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.uxMapNumber.TextChanged

        With PartnerLocateFeatureForm
            Try
                .UseWaitCursor = True

                Dim uxMapnumber As TextBox = .uxMapNumber
                Dim uxTaxlot As TextBox = .uxTaxlot

                If uxMapnumber.AutoCompleteCustomSource.Contains(uxMapnumber.Text.Trim) Then

                    Dim theMapNumberVal As String = uxMapnumber.Text.Trim
                    If theMapNumberVal = String.Empty Then Exit Sub

                    Dim theTaxlotFClass As IFeatureClass = TaxlotFeatureLayer.FeatureClass

                    Dim theQueryFilter As IQueryFilter = New QueryFilter
                    theQueryFilter.SubFields = EditorExtension.TaxLotSettings.MapNumberField & ", " & EditorExtension.TaxLotSettings.TaxlotField
                    '' NOTE: [SC] The following whereclause was slower than querying the entire recordset and then filtering 
                    '' results (see feature loop below).
                    'Dim theWhereClause As String
                    'theWhereClause = "[" & EditorExtension.TaxLotSettings.MapNumberField & "] = '" & theMapNumberVal & "'"
                    'theQueryFilter.WhereClause = formatWhereClause(theWhereClause, theTaxlotFClass)

                    '' NOTE: [SC] Calculating a AutoCompleteCustomSource using the List of strings is considerably faster than
                    '' adding with the .add method.
                    Dim taxlotStringList As List(Of String) = New List(Of String)
                    Dim theFeatCursor As IFeatureCursor = theTaxlotFClass.Search(theQueryFilter, True)
                    Dim theFeature As IFeature = theFeatCursor.NextFeature
                    If theFeature Is Nothing Then Exit Sub
                    Dim theTaxlotFieldIdx As Integer = theFeature.Fields.FindField(EditorExtension.TaxLotSettings.TaxlotField)
                    Dim theMapNumberFieldIdx As Integer = theFeature.Fields.FindField(EditorExtension.MapIndexSettings.MapNumberField)

                    Do Until theFeature Is Nothing
                        Dim theTaxlotVal As String = theFeature.Value(theTaxlotFieldIdx).ToString
                        If Not uxTaxlot.AutoCompleteCustomSource.Contains(theTaxlotVal) AndAlso theMapNumberVal = theFeature.Value(theMapNumberFieldIdx).ToString Then
                            taxlotStringList.Add(theTaxlotVal)
                        End If
                        theFeature = theFeatCursor.NextFeature
                    Loop

                    uxTaxlot.AutoCompleteCustomSource.AddRange(taxlotStringList.ToArray)

                Else
                    uxTaxlot.AutoCompleteCustomSource.Clear()
                    uxTaxlot.Text = String.Empty
                End If

            Finally
                .UseWaitCursor = False
            End Try

        End With


    End Sub

    Private Sub uxFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.uxFind.Click

        With PartnerLocateFeatureForm
            Try
                .UseWaitCursor = True

                Dim uxMapnumber As TextBox = PartnerLocateFeatureForm.uxMapNumber
                Dim theMapNumberVal As String = uxMapnumber.Text.Trim

                If theMapNumberVal = String.Empty OrElse Not uxMapnumber.AutoCompleteCustomSource.Contains(theMapNumberVal) Then
                    .UseWaitCursor = False
                    MessageBox.Show("Invalid MapNumber. Please try again.", "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If

                Dim uxTaxlot As TextBox = .uxTaxlot
                Dim theTaxlotVal As String = uxTaxlot.Text.Trim

                If theTaxlotVal <> String.Empty AndAlso Not uxTaxlot.AutoCompleteCustomSource.Contains(theTaxlotVal) Then
                    .UseWaitCursor = False
                    MessageBox.Show("Invalid Taxlot. Please try again.", "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If

                Dim theQueryFilter As IQueryFilter = New QueryFilter
                Dim theXFlayer As IFeatureLayer = Nothing '-- Set as either the MapIndex or Taxlot Feature Layer.

                Dim theWhereClause As String

                If theTaxlotVal = String.Empty Then
                    '[Looking for just a MapIndex...]
                    theXFlayer = MapIndexFeatureLayer
                    theQueryFilter.SubFields = "Shape, " & EditorExtension.MapIndexSettings.MapNumberField
                    theWhereClause = "[" & EditorExtension.MapIndexSettings.MapNumberField & "] = '" & theMapNumberVal & "'"
                Else
                    '[Looking for a MapIndex and Taxlot...]
                    theXFlayer = TaxlotFeatureLayer
                    theQueryFilter.SubFields = "Shape, " & EditorExtension.TaxLotSettings.MapNumberField & ", " & EditorExtension.TaxLotSettings.TaxlotField
                    theWhereClause = "[" & EditorExtension.TaxLotSettings.MapNumberField & "] = '" & theMapNumberVal & "' AND [" & EditorExtension.TaxLotSettings.TaxlotField & "] = '" & theTaxlotVal & "'"
                End If

                theQueryFilter.WhereClause = formatWhereClause(theWhereClause, theXFlayer.FeatureClass)

                Dim theXFClass As IFeatureClass = theXFlayer.FeatureClass
                Dim theFeatCursor As IFeatureCursor = theXFClass.Search(theQueryFilter, True)
                Dim thisFeature As IFeature = theFeatCursor.NextFeature

                If thisFeature Is Nothing Then '-- Must be due to invalid mapindex or taxlot entered into the text boxes.
                    MessageBox.Show("Feature does not exist.", "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                Else
                    Dim theEnvelope As IEnvelope = thisFeature.Shape.Envelope
                    Do Until thisFeature Is Nothing
                        theEnvelope.Union(thisFeature.Shape.Envelope)
                        thisFeature = theFeatCursor.NextFeature
                    Loop
                    ZoomToEnvelope(theEnvelope)
                    SetSelectedFeature(theXFlayer, thisFeature, True) ' TODO: [SC] This is not working here. Add new procedure for multiple features.
                End If
            Finally
                .UseWaitCursor = False
            End Try
        End With


    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerLocateFeatureForm.uxHelp.Click
        ' TODO: [ALL] Evaluate help systems and implement.
        MessageBox.Show("Sorry. Help not implemented at this time.")
    End Sub

#End Region

#Region "Methods"

    Friend Sub DoButtonOperation()

        Try
            PartnerLocateFeatureForm.uxMapNumber.Enabled = False
            PartnerLocateFeatureForm.uxTaxlot.Enabled = False
            PartnerLocateFeatureForm.uxFind.Enabled = False

            ' Check for valid data.
            CheckValidMapIndexDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & NewLine & _
                                "Please load this dataset into your map.", _
                                "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            Else
                PartnerLocateFeatureForm.uxMapNumber.Enabled = True
                PartnerLocateFeatureForm.uxFind.Enabled = True
            End If

            CheckValidTaxlotDataProperties()
            If HasValidTaxlotData Then
                PartnerLocateFeatureForm.uxTaxlot.Enabled = True
            Else
                PartnerLocateFeatureForm.uxTaxlot.Enabled = False
            End If

            'PartnerLocateFeatureForm.ShowDialog() 'MODAL
            PartnerLocateFeatureForm.Show() 'NON-MODAL

        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)

        End Try

    End Sub

#End Region

#End Region

#Region "Inherited Class Members"

#Region "Properties"

    ''' <summary>
    ''' Called by ArcMap once per second to check if the command is enabled.
    ''' </summary>
    ''' <remarks>WARNING: Do not put computation-intensive code here.</remarks>
    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Dim canEnable As Boolean
            canEnable = EditorExtension.CanEnableExtendedEditing
            Return canEnable
        End Get
    End Property

#End Region

#Region "Methods"

    ''' <summary>
    ''' Called by ArcMap when this command is created.
    ''' </summary>
    ''' <param name="hook">A generic <c>Object</c> hook to an instance of the application.</param>
    ''' <remarks>The application hook may not point to an <c>IMxApplication</c> object.</remarks>
    Public Overrides Sub OnCreate(ByVal hook As Object)

        Try
            If Not hook Is Nothing Then

                'Disable tool if parent application is not ArcMap
                If TypeOf hook Is IMxApplication Then
                    _application = DirectCast(hook, IApplication)
                    setPartnerLocateFeatureForm(New LocateFeatureForm())
                    MyBase.m_enabled = True
                Else
                    MyBase.m_enabled = False
                End If
            End If

            ' NOTE: Add other initialization code here...

        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try
    End Sub

    Public Overrides Sub OnClick()
        Try
            DoButtonOperation()
        Catch ex As Exception
            EditorExtension.ProcessUnhandledException(ex)
        End Try
    End Sub

#End Region

#End Region

#Region "Implemented Interface Members"

#Region "IDisposable Interface Implementation"

    Private _isDuringDispose As Boolean ' Used to track whether Dispose() has been called and is in progress.

    ''' <summary>
    ''' Dispose of managed and unmanaged resources.
    ''' </summary>
    ''' <param name="disposing">True or False.</param>
    ''' <remarks>
    ''' <para>Member of System::IDisposable.</para>
    ''' <para>Dispose executes in two distinct scenarios. 
    ''' If disposing equals true, the method has been called directly
    ''' or indirectly by a user's code. Managed and unmanaged resources
    ''' can be disposed.</para>
    ''' <para>If disposing equals false, the method has been called by the 
    ''' runtime from inside the finalizer and you should not reference 
    ''' other objects. Only unmanaged resources can be disposed.</para>
    ''' </remarks>
    Friend Sub Dispose(ByVal disposing As Boolean)
        ' Check to see if Dispose has already been called.
        If Not Me._isDuringDispose Then

            ' Flag that disposing is in progress.
            Me._isDuringDispose = True

            If disposing Then
                ' Free managed resources when explicitly called.

                ' Dispose managed resources here.
                '   e.g. component.Dispose()

            End If

            ' Free "native" (shared unmanaged) resources, whether 
            ' explicitly called or called by the runtime.

            ' Call the appropriate methods to clean up 
            ' unmanaged resources here.
            _bitmapResourceName = Nothing
            MyBase.m_bitmap = Nothing

            ' Flag that disposing has been finished.
            _isDuringDispose = False

        End If

    End Sub

#Region " IDisposable Support "

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

#End Region

#End Region

#Region "Other Members"

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "ecace054-c346-4be5-a980-8d81ede68fcb"
    Public Const InterfaceId As String = "effc83d3-a320-460c-9b1e-7f48a4b282d8"
    Public Const EventsId As String = "591a6993-fcb1-4c0c-879f-5488e0c88312"
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



