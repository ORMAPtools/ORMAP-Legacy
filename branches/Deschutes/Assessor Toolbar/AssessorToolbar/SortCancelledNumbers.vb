Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Editor
Imports System.Windows.forms
Imports AssessorToolbar.Utilities

<ComClass(SortCancelledNumbers.ClassId, SortCancelledNumbers.InterfaceId, SortCancelledNumbers.EventsId), _
 ProgId("AssessorToolbar.SortCancelledNumbers")> _
Public NotInheritable Class SortCancelledNumbers
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "9b11625c-0739-4b7f-9828-916b72cfe47c"
    Public Const InterfaceId As String = "54c6f6f1-33df-4406-8817-75a5b2194932"
    Public Const EventsId As String = "5df1cdf2-0663-4aca-aa67-c123bca0699d"
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
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "System Class Properties and Events"
    Private _application As IApplication
    Private _editor As IEditor
    Private _mapnumber As String

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "AssessorToolbar"  'localizable text 
        MyBase.m_caption = "SortCancelledNumbers"   'localizable text 
        MyBase.m_message = "Sorts the items in the Cancelled Numbers table."   'localizable text 
        MyBase.m_toolTip = "Sorts the items in the Cancelled Numbers table." 'localizable text 
        MyBase.m_name = MyBase.m_category & "_SortCancelledNumbers"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            'TODO: change bitmap name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try


    End Sub


    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            _application = CType(hook, IApplication)
            _editor = _application.FindExtensionByName("ESRI Object Editor")

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If

        ' TODO:  Add other initialization code
    End Sub


    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            Dim canEnable As Boolean
            canEnable = _editor.EditState
            Return canEnable
        End Get
    End Property


    Public Overrides Sub OnClick()

        '-- Check for necessary items before proceeding...
        If GetCancelledNumbersTable() Is Nothing Then
            MessageBox.Show("Unable to find the cancelled numbers table.  Please ensure it's loaded into your project", "Error", MessageBoxButtons.OK)
            Exit Sub
        End If
        If GetFeatureLayerByName("SeeMaps", _application.Document) Is Nothing Then
            MessageBox.Show("Unable to find the SeeMaps feature class.  Please ensure it's loaded into your project", "Error", MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim theSelectMapIndexDialog As SelectMapindexDialog = MakeSelectMapIndexDialog(_application.Document)
        If theSelectMapIndexDialog.ShowDialog = DialogResult.Cancel Then Exit Sub
        _mapnumber = theSelectMapIndexDialog.MapNumber

        PartnerSortCancelledNumbersForm.ShowDialog()

    End Sub

#End Region


#Region "Custom Class Properties"

    Private WithEvents _partnerSortCancelledNumbersForm As SortCancelledNumbersForm

    Friend ReadOnly Property PartnerSortCancelledNumbersForm() As SortCancelledNumbersForm
        Get
            If _partnerSortCancelledNumbersForm Is Nothing OrElse _partnerSortCancelledNumbersForm.IsDisposed Then
                setPartnerSortCancelledNumbersForm(New SortCancelledNumbersForm())
            End If
            Return _partnerSortCancelledNumbersForm
        End Get
    End Property

    Private Sub setPartnerSortCancelledNumbersForm(ByVal value As SortCancelledNumbersForm)
        If value IsNot Nothing Then
            _partnerSortCancelledNumbersForm = value
            ' Subscribe to partner form events.
            AddHandler _partnerSortCancelledNumbersForm.Load, AddressOf PartnerSortCancelledNumbersForm_Load
            AddHandler _partnerSortCancelledNumbersForm.uxTop.Click, AddressOf uxTop_Click
            AddHandler _partnerSortCancelledNumbersForm.uxUp.Click, AddressOf uxUp_Click
            AddHandler _partnerSortCancelledNumbersForm.uxDown.Click, AddressOf uxDown_Click
            AddHandler _partnerSortCancelledNumbersForm.uxBottom.Click, AddressOf uxBottom_Click
            AddHandler _partnerSortCancelledNumbersForm.uxCancel.Click, AddressOf uxCancel_Click
            AddHandler _partnerSortCancelledNumbersForm.uxOK.Click, AddressOf uxOK_Click
            AddHandler _partnerSortCancelledNumbersForm.uxAdd.Click, AddressOf uxAdd_Click
            AddHandler _partnerSortCancelledNumbersForm.uxDelete.Click, AddressOf uxDelete_Click
            AddHandler _partnerSortCancelledNumbersForm.uxCancelledNumbers.SelectedIndexChanged, AddressOf uxCancelledNumbers_SelectedIndexChanged

        Else
            ' Unsubscribe to partner form events.
            RemoveHandler _partnerSortCancelledNumbersForm.Load, AddressOf PartnerSortCancelledNumbersForm_Load
            RemoveHandler _partnerSortCancelledNumbersForm.uxTop.Click, AddressOf uxTop_Click
            RemoveHandler _partnerSortCancelledNumbersForm.uxUp.Click, AddressOf uxUp_Click
            RemoveHandler _partnerSortCancelledNumbersForm.uxDown.Click, AddressOf uxDown_Click
            RemoveHandler _partnerSortCancelledNumbersForm.uxBottom.Click, AddressOf uxBottom_Click
            RemoveHandler _partnerSortCancelledNumbersForm.uxCancel.Click, AddressOf uxCancel_Click
            RemoveHandler _partnerSortCancelledNumbersForm.uxOK.Click, AddressOf uxOK_Click
            RemoveHandler _partnerSortCancelledNumbersForm.uxAdd.Click, AddressOf uxAdd_Click
            RemoveHandler _partnerSortCancelledNumbersForm.uxDelete.Click, AddressOf uxDelete_Click
            RemoveHandler _partnerSortCancelledNumbersForm.uxCancelledNumbers.SelectedIndexChanged, AddressOf uxCancelledNumbers_SelectedIndexChanged

        End If
    End Sub

#End Region


#Region "Custom Class Events"


    Private Sub PartnerSortCancelledNumbersForm_Load(ByVal sender As Object, ByVal e As System.EventArgs)

        With PartnerSortCancelledNumbersForm
            Try
                .UseWaitCursor = True
                .uxCancelledNumbers.Items.Clear()

                If .uxCancelledNumbers.SelectedIndex = -1 Then .uxDelete.Enabled = False

                Dim theQueryFilter As IQueryFilter = New QueryFilter
                theQueryFilter.WhereClause = "MapNumber = '" & _mapnumber & "'"

                Dim theTableSort As ITableSort = New TableSort
                With theTableSort
                    .Fields = "SortOrder"
                    .Ascending("SortOrder") = True
                    .QueryFilter = theQueryFilter
                    .Table = GetCancelledNumbersTable()
                End With

                theTableSort.Sort(Nothing)

                Dim theCursor As ICursor = theTableSort.Rows
                Dim theRow As IRow = theCursor.NextRow
                While Not theRow Is Nothing
                    If IsDBNull(theRow.Value(2)) Then
                        .uxCancelledNumbers.Items.Add("Null")
                    Else
                        .uxCancelledNumbers.Items.Add(theRow.Value(2))
                    End If
                    theRow = theCursor.NextRow
                End While

            Finally
                .UseWaitCursor = False
            End Try


        End With

    End Sub

    Private Sub MoveItems(ByVal Direction As String) 'pass -1 for moveup and + 1 for move down

        With PartnerSortCancelledNumbersForm.uxCancelledNumbers

            Dim currentIndex As Integer = .SelectedIndex
            Dim currentText As String = .Items(.SelectedIndex).ToString

            Dim newIndex As Integer = 0
            Select Case Direction
                Case "+"
                    newIndex = currentIndex - 1
                Case "-"
                    newIndex = currentIndex + 1
                Case "++"
                    newIndex = 0
                Case "--"
                    newIndex = .Items.Count - 1
            End Select

            .Items.Remove(currentText)
            .Items.Insert(newIndex, currentText)
            .SelectedIndex = newIndex

        End With
    End Sub

    Private Sub uxCancelledNumbers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxCancelledNumbers.SelectedIndexChanged

        With PartnerSortCancelledNumbersForm
            If .uxCancelledNumbers.SelectedIndex = 0 Then
                .uxUp.Enabled = False
                .uxTop.Enabled = False
            Else
                .uxUp.Enabled = True
                .uxTop.Enabled = True
            End If

            If .uxCancelledNumbers.SelectedIndex = .uxCancelledNumbers.Items.Count - 1 Then
                .uxDown.Enabled = False
                .uxBottom.Enabled = False
            Else
                .uxDown.Enabled = True
                .uxBottom.Enabled = True
            End If

            .uxDelete.Enabled = True

        End With

    End Sub

    Private Sub uxTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxTop.Click
        MoveItems("++")
    End Sub

    Private Sub uxUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxUp.Click
        MoveItems("+")
    End Sub

    Private Sub uxDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxDown.Click
        MoveItems("-")
    End Sub

    Private Sub uxBottom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxBottom.Click
        MoveItems("--")
    End Sub

    Private Sub uxCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxCancel.Click
        PartnerSortCancelledNumbersForm.Hide()
    End Sub

    Private Sub uxOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxOK.Click

        '-- Start edit operation (must start edit operation due to versioning)
        Dim theDataset As IDataset = GetCancelledNumbersTable()
        Dim theWorkspaceEdit As IWorkspaceEdit = theDataset.Workspace
        theWorkspaceEdit.StartEditOperation()

        Dim theQueryFilter As IQueryFilter = New QueryFilter
        theQueryFilter.WhereClause = "MapNumber = '" & _mapnumber & "'"
        Dim theCursor As ICursor = GetCancelledNumbersTable.Update(theQueryFilter, False)
        Dim theRow As IRow = theCursor.NextRow
        While Not theRow Is Nothing
            theRow.Value(3) = PartnerSortCancelledNumbersForm.uxCancelledNumbers.FindString(theRow.Value(2))
            theRow.Store()
            theRow = theCursor.NextRow
        End While

        PartnerSortCancelledNumbersForm.Hide()

        '-- Stop edit operation
        theWorkspaceEdit.StopEditOperation()


    End Sub

    Private Sub uxDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxDelete.Click


        With PartnerSortCancelledNumbersForm

            If MessageBox.Show("Are you sure you want to delete this?", "Verify Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub

            '-- Start edit operation (must start edit operation due to versioning)
            Dim theDataset As IDataset = GetCancelledNumbersTable()
            Dim theWorkspaceEdit As IWorkspaceEdit = theDataset.Workspace
            theWorkspaceEdit.StartEditOperation()

            Dim theQueryFilter As IQueryFilter = New QueryFilter
            If .uxCancelledNumbers.Items(.uxCancelledNumbers.SelectedIndex).ToString() = "Null" Then
                theQueryFilter.WhereClause = "MapNumber = '" & _mapnumber & "' AND Taxlot IS NULL"
            Else
                theQueryFilter.WhereClause = "MapNumber = '" & _mapnumber & "' AND Taxlot = '" & .uxCancelledNumbers.Items(.uxCancelledNumbers.SelectedIndex).ToString & "'"
            End If

            Dim theCursor As ICursor = GetCancelledNumbersTable.Search(theQueryFilter, False)
            Dim theRow As IRow = theCursor.NextRow
            If Not theRow Is Nothing Then theRow.Delete()

            '-- Stop edit operation
            theWorkspaceEdit.StopEditOperation()

            .uxCancelledNumbers.Items.Remove(.uxCancelledNumbers.Items(.uxCancelledNumbers.SelectedIndex).ToString)
            .uxDelete.Enabled = False

        End With

    End Sub

    Private Sub uxAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles uxAdd.Click

        Dim theInputDialog As InputDialog = MakeInputDialog("Enter the Cancelled Number:", "Add Cancelled Number", "Please enter a Cancelled Number")
        If theInputDialog.ShowDialog = DialogResult.Cancel Then Exit Sub

        With PartnerSortCancelledNumbersForm

            If .uxCancelledNumbers.Items.Contains(theInputDialog.Value) Then
                MessageBox.Show("Duplicate Entry.  Please try again", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If

            '-- Start edit operation (must start edit operation due to versioning)
            Dim theDataset As IDataset = GetCancelledNumbersTable()
            Dim theWorkspaceEdit As IWorkspaceEdit = theDataset.Workspace
            theWorkspaceEdit.StartEditOperation()

            '-- Add the item to the Cancelled numbers table
            Dim theCancelledNumbersTable As ITable = GetCancelledNumbersTable()
            Dim theCursor As ICursor = theCancelledNumbersTable.Insert(True)
            Dim theRowBuffer As IRowBuffer = theCancelledNumbersTable.CreateRowBuffer
            theRowBuffer.Value(theRowBuffer.Fields.FindField("MapNumber")) = _mapnumber
            theRowBuffer.Value(theRowBuffer.Fields.FindField("Taxlot")) = theInputDialog.Value
            theRowBuffer.Value(theRowBuffer.Fields.FindField("SortOrder")) = .uxCancelledNumbers.Items.Count
            theCursor.InsertRow(theRowBuffer)
            theCursor.Flush()

            '-- Stop edit operation
            theWorkspaceEdit.StopEditOperation()

            .uxCancelledNumbers.Items.Insert(.uxCancelledNumbers.Items.Count, theInputDialog.Value)

        End With

    End Sub

    Function GetCancelledNumbersTable() As ITable

        Dim theMxDocument As IMxDocument = _application.Document
        Dim theTableCollection As IStandaloneTableCollection = theMxDocument.FocusMap
        '-- Hard coded for now...
        Dim theCancelledNumTable As ITable = Nothing

        For i As Integer = 0 To theTableCollection.StandaloneTableCount - 1
            If theTableCollection.StandaloneTable(i).Name.ToUpper = "giscarto.CREATOR_ASR.CANCELLEDNUMBERS".ToUpper Then
                theCancelledNumTable = theTableCollection.StandaloneTable(i).Table
            End If
        Next

        Return theCancelledNumTable

    End Function


#End Region


End Class



