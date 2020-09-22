VERSION 5.00
Begin VB.Form FRegister 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Top Scorer Registration"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   10
      Left            =   1260
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "K"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   12
      Left            =   2220
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "M"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   0
      Left            =   780
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "A"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   2
      Left            =   1740
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "C"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   4
      Left            =   2700
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "E"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   6
      Left            =   3660
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "G"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   8
      Left            =   4620
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "I"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   16
      Left            =   4140
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "Q"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   18
      Left            =   780
      Top             =   1140
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "S"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   20
      Left            =   1740
      Top             =   1140
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "U"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   22
      Left            =   2700
      Top             =   1140
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "W"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   24
      Left            =   3660
      Top             =   1140
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "Y"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   26
      Left            =   4620
      Top             =   1140
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   " "
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   28
      Left            =   1260
      Top             =   1620
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "'"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   30
      Left            =   2220
      Top             =   1620
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "*"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   32
      Left            =   3180
      Top             =   1620
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "@"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   34
      Left            =   4140
      Top             =   1620
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "0"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   44
      Left            =   4620
      Top             =   2100
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "<"
      ForeColor       =   255
      HiLightColor    =   255
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   36
      Left            =   780
      Top             =   2100
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "2"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   38
      Left            =   1740
      Top             =   2100
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "4"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   39
      Left            =   2700
      Top             =   2100
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "6"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   40
      Left            =   3660
      Top             =   2100
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "8"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.Cell idName 
      Height          =   720
      Index           =   0
      Left            =   900
      Top             =   2700
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   " "
      Mode            =   0
      ForeColor       =   65535
      ProtectedColor  =   65535
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   7
      Left            =   4140
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "H"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   14
      Left            =   3180
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "O"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   5
      Left            =   3180
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "F"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   1
      Left            =   1260
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "B"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   3
      Left            =   2220
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "D"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   17
      Left            =   4620
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "R"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   37
      Left            =   1260
      Top             =   2100
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "3"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin VB.Timer tmrTransparent 
      Interval        =   1
      Left            =   120
      Top             =   1980
   End
   Begin Sudoku.Button btnOK 
      Height          =   390
      Left            =   4080
      Top             =   3720
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "OK"
      Enabled         =   0   'False
   End
   Begin Sudoku.Button btnCancel 
      Height          =   390
      Left            =   2940
      Top             =   3720
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "Cancel"
      Enabled         =   -1  'True
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   9
      Left            =   780
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "J"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   11
      Left            =   1740
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "L"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   13
      Left            =   2700
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "N"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   15
      Left            =   3660
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "P"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.Cell idName 
      Height          =   720
      Index           =   1
      Left            =   1740
      Top             =   2700
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   " "
      Mode            =   0
      ForeColor       =   65535
      ProtectedColor  =   65535
   End
   Begin Sudoku.Cell idName 
      Height          =   720
      Index           =   2
      Left            =   2580
      Top             =   2700
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   " "
      Mode            =   0
      ForeColor       =   65535
      ProtectedColor  =   65535
   End
   Begin Sudoku.Cell idName 
      Height          =   720
      Index           =   3
      Left            =   3420
      Top             =   2700
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   " "
      Mode            =   0
      ForeColor       =   65535
      ProtectedColor  =   65535
   End
   Begin Sudoku.Cell idName 
      Height          =   720
      Index           =   4
      Left            =   4260
      Top             =   2700
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   " "
      Mode            =   0
      ForeColor       =   65535
      ProtectedColor  =   65535
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   41
      Left            =   2220
      Top             =   2100
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "5"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   42
      Left            =   3180
      Top             =   2100
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "7"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   43
      Left            =   4140
      Top             =   2100
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "9"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   35
      Left            =   4620
      Top             =   1620
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "1"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   33
      Left            =   3660
      Top             =   1620
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "~"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   31
      Left            =   2700
      Top             =   1620
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "#"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   29
      Left            =   1740
      Top             =   1620
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "^"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   27
      Left            =   780
      Top             =   1620
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "-"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   25
      Left            =   4140
      Top             =   1140
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "Z"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   23
      Left            =   3180
      Top             =   1140
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "X"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   21
      Left            =   2220
      Top             =   1140
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "V"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin Sudoku.ID idChar 
      Height          =   495
      Index           =   19
      Left            =   1260
      Top             =   1140
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "T"
      ForeColor       =   16777215
      HiLightColor    =   16777215
      Locked          =   0   'False
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   5
      Left            =   60
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00808080&
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   3720
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   60
      Top             =   3420
      Width           =   195
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Index           =   2
      Left            =   120
      Top             =   3540
      Width           =   315
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   1
      Left            =   300
      Top             =   3660
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   135
      Index           =   0
      Left            =   -180
      Top             =   3900
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   720
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   720
      Top             =   2940
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   720
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   2655
      Index           =   0
      Left            =   720
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "FRegister.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************
'* Title: FRegister                  *
'* Stamp: 18 July 2007               *
'* Auth : Derio                      *
'* Desc : Capture the hight score ID *
'*************************************

Dim CharIndex As Integer

Private Sub btnCancel_Click()
'** Cancel registering hiscore

  If MSupport.IsYes("Are you sure not to register your score?") Then
    Me.Tag = ""
    MSupport.FadeOut Me
  End If
End Sub

Private Sub btnOK_Click()
'** Finish registering

Dim I As Integer

  'capture the id
  Me.Tag = ""
  For I = 0 To Me.idName.Count - 1
    Me.Tag = Me.Tag & Me.idName(I).Caption
  Next I
  Me.Tag = Trim(Me.Tag)
  
  MSupport.FadeOut Me
End Sub

Private Sub Form_Load()
  CharIndex = -1
End Sub

Private Sub idChar_Click(Index As Integer)
  If idChar(Index).Caption <> "<" Then
    If idChar(Index).Caption = " " And CharIndex = -1 Then Exit Sub
    If CharIndex < Me.idName.Count - 1 Then
      CharIndex = CharIndex + 1
      Me.idName(CharIndex).Caption = idChar(Index).Caption
    End If
    If CharIndex <> -1 Then
      If Not Me.btnOK.Enabled Then Me.btnOK.Enabled = True
    End If
    
  Else
    If CharIndex >= 0 Then
      Me.idName(CharIndex).Caption = " "
      CharIndex = CharIndex - 1
      If CharIndex = -1 Then
        If Me.btnOK.Enabled Then Me.btnOK.Enabled = False
      End If
    End If
  End If
End Sub

Private Sub tmrTransparent_Timer()
  Me.tmrTransparent.Enabled = False
  FadeIn Me
End Sub
