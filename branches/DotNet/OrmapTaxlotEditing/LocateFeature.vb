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
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
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

#Region "Class-Level Constants And Enumerations (none)"
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
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As ArgumentException
            Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try

    End Sub



#End Region

#End Region

#Region "Custom Class Members"

#Region "Fields"

    Private _application As IApplication

#End Region

#Region "Properties"

    Private WithEvents _partnerLocateFeatureForm As LocateFeatureForm

    Friend ReadOnly Property PartnerLocateFeatureForm() As LocateFeatureForm
        Get
            Return _partnerLocateFeatureForm
        End Get
    End Property

    Private Sub setPartnerLocateFeatureForm(ByRef value As LocateFeatureForm)
        If value IsNot Nothing Then
            _partnerLocateFeatureForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerLocateFeatureForm.Load, AddressOf PartnerTaxlotAssignmentForm_Load
            AddHandler _partnerLocateFeatureForm.uxFind.Click, AddressOf uxFind_Click
            AddHandler _partnerLocateFeatureForm.uxHelp.Click, AddressOf uxHelp_Click
        End If
    End Sub

#End Region

#Region "Event Handlers"

    Private Sub PartnerTaxlotAssignmentForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles PartnerTaxlotAssignmentForm.Load

        With PartnerLocateFeatureForm

            If .uxMapnumber.Items.Count = 0 Then '-- Only load the text box the first time the tool is run.
                Dim mapIndexFClass As IFeatureClass = MapIndexFeatureLayer.FeatureClass
                Dim theQueryFilter As IQueryFilter = New QueryFilter
                theQueryFilter.SubFields = "DISTINCT(" & EditorExtension.MapIndexSettings.MapNumberField & ")"

                Dim theFeatCursor As IFeatureCursor = mapIndexFClass.Search(theQueryFilter, False)
                Dim theQueryField As Integer = theFeatCursor.FindField(EditorExtension.MapIndexSettings.MapNumberField)
                Dim theFeature As IFeature = theFeatCursor.NextFeature
                Do Until theFeature Is Nothing
                    .uxMapnumber.Items.Add(theFeature.Value(theQueryField))
                    theFeature = theFeatCursor.NextFeature
                Loop
            End If

        End With

    End Sub

    Private Sub uxFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerTaxlotAssignmentForm.uxFind.Click

        Dim mapNumber As String = Nothing
        Dim taxlot As String = Nothing

        Dim uxMapnumber As ComboBox = PartnerLocateFeatureForm.uxMapnumber
        If uxMapnumber.FindStringExact(uxMapnumber.Text) = -1 Then
            MessageBox.Show("Invalid MapNumber.  Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        Else
            mapNumber = uxMapnumber.Text.Trim
        End If

        Dim uxTaxlot As TextBox = PartnerLocateFeatureForm.uxTaxlot
        If uxTaxlot.Enabled And uxTaxlot.Text.Trim <> "" Then
            taxlot = uxTaxlot.Text.Trim
        End If

        Dim theQueryFilter As IQueryFilter = New QueryFilter
        Dim theXFlayer As IFeatureLayer = Nothing '-- Set as either the MapIndex or Taxlot Feature Layer.

        If taxlot Is Nothing Then '-- Must be MapIndex Feature Layer.
            theXFlayer = MapIndexFeatureLayer
            theQueryFilter.SubFields = EditorExtension.MapIndexSettings.MapNumberField
            theQueryFilter.WhereClause = "[" & EditorExtension.MapIndexSettings.MapNumberField & "]='" & mapNumber & "'"
        Else '-- Taxlot Feature Layer.
            theXFlayer = MapIndexFeatureLayer
            theQueryFilter.SubFields = EditorExtension.MapIndexSettings.MapNumberField & "," & EditorExtension.TaxLotSettings.TaxlotField
            theQueryFilter.WhereClause = "[" & EditorExtension.MapIndexSettings.MapNumberField & "]='" & mapNumber & "' and [" & EditorExtension.TaxLotSettings.TaxlotField & "]='" & taxlot & "'"
        End If

        Dim theXFClass As IFeatureClass = theXFlayer.FeatureClass
        Dim theFeatCursor As IFeatureCursor = theXFClass.Search(theQueryFilter, False)

        Dim theFeature As IFeature = theFeatCursor.NextFeature()

        If theFeature Is Nothing Then '-- Must be due to invalid taxlot entered into text box.
            MessageBox.Show("Taxlot does not exist.", "Invalid Taxlot", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        Else
            ZoomToEnvelope(theFeature.Shape.Envelope)
            SetSelectedFeature(theXFlayer, theFeature)
        End If

    End Sub

    Private Sub uxHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles PartnerTaxlotAssignmentForm.uxHelp.Click
        ' TODO [SC] Evaluate help systems and implement.
        MessageBox.Show("uxHelp clicked")
    End Sub

#End Region

#Region "Methods"

    Friend Sub DoButtonOperation()

        Try
            ' Check for valid data
            CheckValidDataProperties()
            If Not HasValidMapIndexData Then
                MessageBox.Show("Missing data: Valid ORMAP MapIndex layer not found in the map." & vbNewLine & _
                                "Please load this dataset into your map.", _
                                "Locate Feature", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Try
            End If

            If Not HasValidTaxlotData Then
                PartnerLocateFeatureForm.uxTaxlot.Enabled = False
            Else
                PartnerLocateFeatureForm.uxTaxlot.Enabled = True
            End If

            PartnerLocateFeatureForm.ShowDialog()

        Catch ex As Exception
            MessageBox.Show(ex.Message)

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

        If Not hook Is Nothing Then
            _application = DirectCast(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
                setPartnerLocateFeatureForm(New LocateFeatureForm)
            Else
                MyBase.m_enabled = False
            End If
        End If

        ' NOTE: Add other initialization code here...

    End Sub

    Public Overrides Sub OnClick()

        DoButtonOperation()

    End Sub

#End Region

#End Region

#Region "Implemented Interface Members (none)"
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



