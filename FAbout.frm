VERSION 5.00
Begin VB.Form FAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2460
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3825
   ControlBox      =   0   'False
   Icon            =   "FAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTransparent 
      Interval        =   1
      Left            =   180
      Top             =   660
   End
   Begin Sudoku.Button cmdOK 
      Height          =   390
      Left            =   2580
      Top             =   1980
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "OK"
      Enabled         =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1680
      Left            =   720
      TabIndex        =   0
      Top             =   60
      Width           =   2955
      Begin VB.Label lblMotto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Hot Puzzle Craze"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1920
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sudoku 5.1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   165
         Width           =   2640
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) Derio 2006 - 2008"
         Height          =   195
         Left            =   540
         TabIndex        =   1
         Top             =   1380
         Width           =   2220
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sudoku 5.1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   645
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Top             =   180
         Width           =   2640
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   5
      Left            =   60
      Top             =   1500
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00808080&
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   1980
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   60
      Top             =   1680
      Width           =   195
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Index           =   2
      Left            =   120
      Top             =   1800
      Width           =   315
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   1
      Left            =   300
      Top             =   1920
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   135
      Index           =   0
      Left            =   -180
      Top             =   2160
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FAbout.frx":030A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************
'* Title: FAbout                            *
'* Stamp: 15 June 2007                      *
'* Auth : Derio                             *
'* Desc : Show the Information about me :-) *
'********************************************

Private Sub cmdOK_Click()
'** Hide the form with fade out animation

  MSupport.FadeOut Me
End Sub

Private Sub tmrTransparent_Timer()
'** Show the form with fade in animation
'   This timer active after 1 ms

Static Opacity As Integer

  If Tag = "Opening" Then
    Opacity = Opacity + 3
    If Not OpeningComplete Or Opacity < MAX_OPACITY Then
      If Opacity < MAX_OPACITY Then MSupport.MakeTransparent Me.hWnd, Opacity
    End If
    If OpeningComplete Then
      Me.tmrTransparent.Interval = 2000
      Me.Tag = "Closing"
      DoEvents
    End If
    
  ElseIf Tag = "Closing" Then
    Me.tmrTransparent.Enabled = False
    DoEvents
    MSupport.FadeOut Me
    Unload Me
  
  Else
    Me.tmrTransparent.Enabled = False
    DoEvents
    MSupport.FadeIn Me
  End If
End Sub
