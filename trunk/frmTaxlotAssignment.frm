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
'
' File name:            frmTaxlotAssignment
'
' Initial Author:       Type your name here
'
' Date Created:
'
' Description: FORM USED TO CAPTURE UNDERLYING MAPINDEX ATTRIBUTES FOR THE PURPOSE OF POPULATING TAXLOT
'
'
' Entry points:
'       List the public variables and their purposes.
'       List the properties and routines that the module exposes to the rest of the program.
'
' Dependencies:
'       How does this file depend or relate to other files?
'
' Issues:
'       What are unsolved bugs, bottlenecks,
'       possible future enhancements, and
'       descriptions of other issues.
'
' Method:
'       Describe any complex details that make sense on the file level.  This includes explanations
'       of complex algorithms, how different routines within the module interact, and a description
'       of a data structure used in the module.
'
' Updates:
'
'JWM 10/11/2006 added this comment section

Option Explicit
'******************************
' Global/Public Definitions
'------------------------------
' Public API Declarations
'------------------------------

'------------------------------
' Public Enums and Constants
'------------------------------

'------------------------------
' Public variables
'------------------------------

'------------------------------
' Public Types
'------------------------------

'------------------------------
' Public loop variables
'------------------------------

'******************************
' Private Definitions
'------------------------------
' Private API declarations
'------------------------------

'------------------------------
' Private Variables
'------------------------------
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument

'------------------------------
'Private Constants and Enums
'------------------------------
' Variables used by the Error handler function - DO NOT REMOVE
Private Const c_sModuleFileName As String = "frmTaxlotAssignment.frm"

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------

Private Sub cmbTaxlotNum_Click()
    'If one of the generic taxlot numbers is chosen, disable the taxlot textbox
    '++  JWM 10/11/2006 using strcomp function
85:     If StrComp(Me.cmbTaxlotNum.Text, "NUMBER", vbTextCompare) <> 0 Then
86:         Me.txtTaxlotNum.Enabled = False
87:         Me.txtTaxlotNum.Text = ""
88:     Else
89:         Me.txtTaxlotNum.Enabled = True
90:     End If
End Sub

Private Sub cmdAssign_Click()
  On Error GoTo ErrorHandler
  'Must be a number that is 5 characters long
'++  JWM 10/11/2006 using strcomp function
97:   If StrComp(Me.cmbTaxlotNum.Text, "NUMBER", vbTextCompare) = 0 Then
98:     If Not IsNumeric(Me.txtTaxlotNum.Text) Then
99:       MsgBox "Invalid Start Value.  Please enter a valid number", vbOKOnly, "Error"
100:       Me.txtTaxlotNum.SetFocus
      Exit Sub
102:     End If
103:   End If

105:   Me.Hide
  
  Exit Sub
ErrorHandler:
  HandleError True, "cmdAssign_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub



Private Sub cmdHelp_Click()
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
118:     sFilePath = app.Path & "\" & "Assignment_help.rtf"
119:     If modUtils.FileExists(sFilePath) Then
120:     Debug.Assert True 'need a better way to open files
121:         If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
122:             Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
123:         End If
124:     Else
125:         MsgBox "No help available"
126:     End If
End Sub

Private Sub cmdQuit_Click()
130:     Me.Hide
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
    'Populate drop down combobox and set default settings
136:     Set m_pApp = modUtils.GetAppRef    'New AppRef
137:     Set m_pMxDoc = m_pApp.Document
    'Populate with preset values
139:     cmbTaxlotNum.AddItem "NUMBER"
140:     cmbTaxlotNum.AddItem "0ROAD"
141:     cmbTaxlotNum.AddItem "WATER"
142:     cmbTaxlotNum.AddItem "0RLRD"
143:     cmbTaxlotNum.AddItem "00GAP"
144:     cmbTaxlotNum.AddItem "00LAP"
145:     cmbTaxlotNum.Text = "NUMBER" 'By default
    
147:     tglAutoYes.Value = False
148:     tglAutoNo.Value = True
149:     tglBy1.Value = False
150:     tglBy10.Value = False
151:     tglBy100.Value = True
    
    

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub tglAutoNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

163:     tglAutoYes.Value = False

  Exit Sub
ErrorHandler:
  HandleError True, "tglAutoNo_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub tglAutoYes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

173:     tglAutoNo.Value = False

  Exit Sub
ErrorHandler:
  HandleError True, "tglAutoYes_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub tglBy1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

183:     tglBy10.Value = False
184:     tglBy100.Value = False

  Exit Sub
ErrorHandler:
  HandleError True, "tglBy1_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


Private Sub tglBy10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

195:     tglBy1.Value = False
196:     tglBy100.Value = False

  Exit Sub
ErrorHandler:
  HandleError True, "tglBy10_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub tglBy100_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

206:     tglBy1.Value = False
207:     tglBy10.Value = False

  Exit Sub
ErrorHandler:
  HandleError True, "tglBy100_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


