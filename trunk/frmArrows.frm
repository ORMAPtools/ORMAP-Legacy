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
'FORM USED TO GENERATE HOOKS AND ARROWS
Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument

Private Sub cmdArrow_Click()
6:     Me.lblCurrentTool.Caption = "arrow"
7:     Me.Hide
End Sub

Private Sub cmdDimension_Click()
11:     Me.lblCurrentTool.Caption = "dimension"
12:     Me.Hide
End Sub

Private Sub cmdHelp_Click()
    'Open a custom help file in Internet Explorer
    'Requires a file called help.htm in the same dir as the application dll
    Dim sFilePath As String
19:     sFilePath = app.Path & "\" & "Arrows_help.rtf"
20:     If modUtils.FileExists(sFilePath) Then
21:         If FileExists("C:\Program Files\Windows NT\Accessories\wordpad.exe") Then
22:             Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & sFilePath, 1
23:         End If
24:     Else
25:         MsgBox "No help available"
26:     End If
End Sub

Private Sub cmdHook_Click()
30:     Me.lblCurrentTool.Caption = "hook"
31:     Me.Hide
    
End Sub

