VERSION 5.00
Begin VB.Form FQuestion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrTransparent 
      Interval        =   1
      Left            =   840
      Top             =   120
   End
   Begin Sudoku.Button btnCommand 
      Height          =   390
      Index           =   0
      Left            =   3420
      Top             =   1500
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "No"
      Enabled         =   -1  'True
   End
   Begin Sudoku.Button btnCommand 
      Height          =   390
      Index           =   1
      Left            =   2280
      Top             =   1500
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "Yes"
      Enabled         =   -1  'True
   End
   Begin VB.Shape shpAccessories 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   5
      Left            =   60
      Top             =   1020
      Width           =   135
   End
   Begin VB.Shape shpAccessories 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00808080&
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   1500
      Width           =   135
   End
   Begin VB.Shape shpAccessories 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   60
      Top             =   1200
      Width           =   195
   End
   Begin VB.Shape shpAccessories 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Index           =   2
      Left            =   120
      Top             =   1320
      Width           =   315
   End
   Begin VB.Shape shpAccessories 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   1
      Left            =   300
      Top             =   1440
      Width           =   255
   End
   Begin VB.Shape shpAccessories 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   135
      Index           =   0
      Left            =   -120
      Top             =   1680
      Width           =   7335
   End
   Begin VB.Image imgLogo 
      Height          =   765
      Index           =   2
      Left            =   60
      Picture         =   "FQuestion.frx":0000
      Top             =   60
      Width           =   795
   End
   Begin VB.Image imgLogo 
      Height          =   765
      Index           =   1
      Left            =   60
      Picture         =   "FQuestion.frx":2022
      Top             =   60
      Width           =   795
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3435
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgLogo 
      Height          =   765
      Index           =   0
      Left            =   60
      Picture         =   "FQuestion.frx":4044
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "FQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************
'* Title: FQuestion                    *
'* Stamp: 4 July 2007                  *
'* Auth : Derio                        *
'* Desc : Replace standard Message Box *
'***************************************


Private Sub btnCommand_Click(Index As Integer)
'** Capture the state and hide this form with fade out effect

  Tag = Me.btnCommand(Index).Caption
  MSupport.FadeOut Me
End Sub

Private Sub tmrTransparent_Timer()
'** Show the form with fade in effect

  Me.tmrTransparent.Enabled = False
  DoEvents
  MSupport.FadeIn Me
End Sub
