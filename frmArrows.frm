VERSION 5.00
Begin VB.Form frmArrows 
   Caption         =   "Add"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   2925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdHook 
      Caption         =   "Hook"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdDimension 
      Caption         =   "Dimension Arrow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdArrow 
      Caption         =   "Arrow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblCurrentTool 
      Caption         =   "none"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmArrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' File name:            frmArrows
'
' Initial Author:       Type your name here
'
' Date Created:     10/11/2006
'
' Description:  FORM USED TO GENERATE HOOKS AND ARROWS
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
'               None

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
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument
'------------------------------
'Private Constants and Enums
'------------------------------

'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------

Private Sub cmdArrow_Click()
78:     Me.lblCurrentTool.Caption = "arrow"
79:     Me.Hide
End Sub

Private Sub cmdDimension_Click()
83:     Me.lblCurrentTool.Caption = "dimension"
84:     Me.Hide
End Sub

Private Sub cmdHelp_Click()
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
91:     sFilePath = app.Path & "\" & "Arrows_help.rtf"
92:     If modUtils.FileExists(sFilePath) Then
93:     Debug.Assert True ' need a different way to open the help file
94:         If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
95:             Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
96:         End If
97:     Else
98:         MsgBox "No help available"
99:     End If
End Sub

Private Sub cmdHook_Click()
103:     Me.lblCurrentTool.Caption = "hook"
104:     Me.Hide
    
End Sub

