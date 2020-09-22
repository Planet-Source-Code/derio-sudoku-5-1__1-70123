VERSION 5.00
Begin VB.Form FHowTo1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " How to - The Rule"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrTransparent 
      Interval        =   1
      Left            =   4080
      Top             =   3120
   End
   Begin Sudoku.Button cmdCommand 
      Height          =   390
      Index           =   0
      Left            =   4740
      Top             =   3180
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "Close"
      Enabled         =   -1  'True
   End
   Begin Sudoku.Button cmdCommand 
      Height          =   390
      Index           =   1
      Left            =   5880
      Top             =   3180
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "Next"
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FHowTo1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2340
      Width           =   6735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   5
      Left            =   60
      Top             =   2700
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00808080&
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   3180
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   60
      Top             =   2880
      Width           =   195
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Index           =   2
      Left            =   120
      Top             =   3000
      Width           =   315
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   1
      Left            =   300
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   135
      Index           =   0
      Left            =   -60
      Top             =   3360
      Width           =   7335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FHowTo1.frx":00C0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   2100
      TabIndex        =   1
      Top             =   1020
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FHowTo1.frx":019F
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   0
      Left            =   2100
      TabIndex        =   0
      Top             =   120
      Width           =   4875
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   180
      Picture         =   "FHowTo1.frx":0239
      Top             =   240
      Width           =   1680
   End
End
Attribute VB_Name = "FHowTo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************
'* Title: FHowTo1                        *
'* Stamp: 3 July 2007                    *
'* Auth : Derio                          *
'* Desc : Show general information about *
'*        the Sudoku rules               *
'*****************************************


Private Sub cmdCommand_Click(Index As Integer)
'** Capture and send the state and then fade out the form

  Tag = Me.cmdCommand(Index).Caption
  MSupport.FadeOut Me
End Sub

Private Sub tmrTransparent_Timer()
'** Show the form with fade in animation

  Me.tmrTransparent.Enabled = False
  DoEvents
  MSupport.FadeIn Me
End Sub
