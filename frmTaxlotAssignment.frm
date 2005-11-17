VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTaxlotAssignment 
   Caption         =   "Taxlot Assignment"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIncrementValue 
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAssign 
      Caption         =   "Assign"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox cmbTaxlotNum 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtTaxlotNum 
      Height          =   315
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin MSForms.ToggleButton tglBy100 
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   1080
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "100"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton tglBy10 
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   1080
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "10"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton tglBy1 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   1080
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "1"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton tglAutoNo 
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   600
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "No"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton tglAutoYes 
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   600
      Width           =   855
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "Yes"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label3 
      Caption         =   "Increment:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Autoincrement:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Taxlot:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmTaxlotAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FORM USED TO CAPTURE UNDERLYING MAPINDEX ATTRIBUTES FOR THE PURPOSE OF POPULATING TAXLOT

' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "C:\active\ModelingWorkshop_01-05-05\CustomCode\ormap\frmTaxlotAssignment.frm"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument

Private Sub cmbTaxlotNum_Click()
    'If one of the generic taxlot numbers is chosen, disable the taxlot textbox
11:     If Me.cmbTaxlotNum.Text <> "NUMBER" Then
12:         Me.txtTaxlotNum.Enabled = False
13:         Me.txtTaxlotNum.Text = ""
14:     Else
15:         Me.txtTaxlotNum.Enabled = True
16:     End If
End Sub

Private Sub cmdAssign_Click()
  On Error GoTo ErrorHandler
  'Must be a number that is 5 characters long
22:   If Me.cmbTaxlotNum.Text = "NUMBER" Then
23:     If Not IsNumeric(Me.txtTaxlotNum.Text) Then
24:       MsgBox "Invalid Start Value.  Please enter a valid number", vbOKOnly, "Error"
25:       Me.txtTaxlotNum.SetFocus
      Exit Sub
27:     End If
28:   End If

30:   Me.Hide
  
  Exit Sub
ErrorHandler:
  HandleError True, "cmdAssign_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub



Private Sub cmdHelp_Click()
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
43:     sFilePath = app.Path & "\" & "Assignment_help.rtf"
44:     If modUtils.FileExists(sFilePath) Then
45:         If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
46:             Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
47:         End If
48:     Else
49:         MsgBox "No help available"
50:     End If
End Sub

Private Sub cmdQuit_Click()
54:     Me.Hide
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
    'Populate drop down combobox and set default settings
60:     Set m_pApp = modUtils.GetAppRef    'New AppRef
61:     Set m_pMxDoc = m_pApp.Document
    'Populate with preset values
63:     cmbTaxlotNum.AddItem "NUMBER"
64:     cmbTaxlotNum.AddItem "0ROAD"
65:     cmbTaxlotNum.AddItem "WATER"
66:     cmbTaxlotNum.AddItem "0RLRD"
67:     cmbTaxlotNum.AddItem "00GAP"
68:     cmbTaxlotNum.AddItem "00LAP"
69:     cmbTaxlotNum.Text = "NUMBER" 'By default
    
71:     tglAutoYes.Value = False
72:     tglAutoNo.Value = True
73:     tglBy1.Value = False
74:     tglBy10.Value = False
75:     tglBy100.Value = True
    
    

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub tglAutoNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

87:     tglAutoYes.Value = False

  Exit Sub
ErrorHandler:
  HandleError True, "tglAutoNo_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub tglAutoYes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

97:     tglAutoNo.Value = False

  Exit Sub
ErrorHandler:
  HandleError True, "tglAutoYes_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub tglBy1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

107:     tglBy10.Value = False
108:     tglBy100.Value = False

  Exit Sub
ErrorHandler:
  HandleError True, "tglBy1_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


Private Sub tglBy10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

119:     tglBy1.Value = False
120:     tglBy100.Value = False

  Exit Sub
ErrorHandler:
  HandleError True, "tglBy10_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub tglBy100_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

130:     tglBy1.Value = False
131:     tglBy10.Value = False

  Exit Sub
ErrorHandler:
  HandleError True, "tglBy100_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


