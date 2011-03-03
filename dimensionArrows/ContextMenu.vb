Imports ESRI.ArcGIS.ADF.CATIDs
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.esriSystem
Imports dimensionArrows.curvedArrows

''' <summary>
''' Create a context menu using IMultiItem
''' </summary>
''' <remarks></remarks>
<ComClass(ContextMenu.ClassId, ContextMenu.InterfaceId, ContextMenu.EventsId), _
 ProgId("dimensionArrows.contextMenu")> _
Public Class contextMenu
    Implements IMultiItem

    Private screenPosition As Point

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
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "F0FBA115-5FF2-4D77-B5A1-8C425505CACA"
    Public Const InterfaceId As String = "B4618482-A041-4DDB-B608-C8E22310297D"
    Public Const EventsId As String = "03D1FC36-DED7-443E-99E1-C470BF44DA5B"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public ReadOnly Property Caption() As String Implements IMultiItem.Caption
        Get
            Return "Arrow Menu"
        End Get
    End Property

    Public ReadOnly Property HelpContextID() As Integer Implements IMultiItem.HelpContextID
        Get
            Return 0
        End Get
    End Property

    Public ReadOnly Property HelpFile() As String Implements IMultiItem.HelpFile
        Get
            Return String.Empty
        End Get
    End Property

    Public ReadOnly Property ItemBitmap(ByVal index As Integer) As Integer Implements IMultiItem.ItemBitmap
        Get
            Return 0
        End Get
    End Property

    Public ReadOnly Property ItemCaption(ByVal index As Integer) As String Implements IMultiItem.ItemCaption
        Get
            Return _menuitems(index)
        End Get
    End Property

    Public ReadOnly Property ItemChecked(ByVal index As Integer) As Boolean Implements IMultiItem.ItemChecked
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property ItemEnabled(ByVal index As Integer) As Boolean Implements IMultiItem.ItemEnabled
        Get
            Return True
        End Get
    End Property

    Public ReadOnly Property Message() As String Implements IMultiItem.Message
        Get
            Return String.Empty
        End Get
    End Property

    Public ReadOnly Property Name() As String Implements IMultiItem.Name
        Get
            Return "OR-DOR_dimensionArrows_contextMenu"
        End Get
    End Property

    Public Sub OnItemClick(ByVal index As Integer) Implements IMultiItem.OnItemClick
        Select Case Me.ItemCaption(index)
            Case SHORTER
                arrowUtilities.keyCommands(Windows.Forms.Keys.Down, 1)
            Case LONGER
                arrowUtilities.keyCommands(Windows.Forms.Keys.Up, 1)
            Case FLIP
                arrowUtilities.keyCommands(Windows.Forms.Keys.F, 0)
            Case UNLOCK
                arrowUtilities.keyCommands(Windows.Forms.Keys.U, 0)
            Case SWITCH
                arrowUtilities.keyCommands(Windows.Forms.Keys.S, 0)
            Case CANCEL
                arrowUtilities.clearAll()
            Case SCALE10
                _arrowScale = 0.1
            Case SCALE20
                _arrowScale = 0.2
            Case SCALE30
                _arrowScale = 0.3
            Case SCALE40
                _arrowScale = 0.4
            Case SCALE50
                _arrowScale = 0.5
            Case SCALE100
                _arrowScale = 1
            Case SCALE200
                _arrowScale = 2
            Case SCALE400
                _arrowScale = 4
            Case SCALE800
                _arrowScale = 8
            Case SCALE1000
                _arrowScale = 10
            Case SCALE2000
                _arrowScale = 20
            Case NARROWER
                _zigzagWidth = _zigzagWidth - 1
                If _zigzagWidth < 1 Then _zigzagWidth = 1
            Case WIDER
                _zigzagWidth = _zigzagWidth + 1
            Case TOPOINT
                _zigzagPosition = _zigzagPosition - 2.5
                If _zigzagPosition < 1 Then _zigzagPosition = 1
            Case TOEND
                _zigzagPosition = _zigzagPosition + 2.5
                If _zigzagPosition > 19 Then _zigzagPosition = 19
            Case CURVELESS
                _zigzagCurve = _zigzagCurve - 1
            Case CURVEMORE
                _zigzagCurve = _zigzagCurve + 1
            Case STYLE_STRAIGHT
                _thisArrow.category = arrowCategories.SingleArrow
                _thisArrow.style = arrowStyles.Straight
                SaveSetting("OR_DOR_dimensionArrows", "default", "category", _
                    arrowCategories.SingleArrow)
                SaveSetting("OR_DOR_dimensionArrows", "default", "style", _
                    arrowStyles.Straight)
            Case STYLE_LEADER
                _thisArrow.category = arrowCategories.SingleArrow
                _thisArrow.style = arrowStyles.Leader
                SaveSetting("OR_DOR_dimensionArrows", "default", "category", _
                    arrowCategories.SingleArrow)
                SaveSetting("OR_DOR_dimensionArrows", "default", "style", _
                    arrowStyles.Leader)
            Case STYLE_ZIGZAG
                _thisArrow.category = arrowCategories.SingleArrow
                _thisArrow.style = arrowStyles.Zigzag
                SaveSetting("OR_DOR_dimensionArrows", "default", "category", _
                    arrowCategories.SingleArrow)
                SaveSetting("OR_DOR_dimensionArrows", "default", "style", _
                    arrowStyles.Zigzag)
            Case STYLE_FREEFORM
                _thisArrow.category = arrowCategories.SingleArrow
                _thisArrow.style = arrowStyles.Freeform
                SaveSetting("OR_DOR_dimensionArrows", "default", "category", _
                    arrowCategories.SingleArrow)
                SaveSetting("OR_DOR_dimensionArrows", "default", "style", _
                    arrowStyles.Freeform)
            Case SAVEDEFAULT
                SaveSetting("OR_DOR_dimensionArrows", "default", "zigzagWidth", _zigzagWidth)
                SaveSetting("OR_DOR_dimensionArrows", "default", "zigzagCurve", _zigzagCurve)
                SaveSetting("OR_DOR_dimensionArrows", "default", "zigzagPosition", _
                    _zigzagPosition)
            Case FINISH
                placeFreeformArrow()
            Case HELP
                showHelp()
        End Select

        If _pointNumber <> 1 Then
            Windows.Forms.Cursor.Position = screenPosition
            showLineFeedback(_lastPoint)
        End If
    End Sub

    Public Function OnPopup(ByVal hook As Object) As Integer _
        Implements IMultiItem.OnPopup
        screenPosition = Windows.Forms.Cursor.Position
        Return UBound(_menuitems) + 1
    End Function
End Class



