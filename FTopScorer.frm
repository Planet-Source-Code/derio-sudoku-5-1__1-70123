VERSION 5.00
Begin VB.Form FTopScorer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Top Five Sudoku Mania"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrTransparent 
      Interval        =   1
      Left            =   120
      Top             =   2520
   End
   Begin Sudoku.ID idNo 
      Height          =   495
      Index           =   0
      Left            =   720
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   "1"
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.Button btnOK 
      Height          =   390
      Left            =   4980
      Top             =   3660
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "OK"
      Enabled         =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   17
      Left            =   4620
      Top             =   2820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   19
      Left            =   5580
      Top             =   2820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   12
      Left            =   4140
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   14
      Left            =   5100
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   9
      Left            =   4620
      Top             =   1500
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   11
      Left            =   5580
      Top             =   1500
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   4
      Left            =   4140
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   6
      Left            =   5100
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   3
      Left            =   5580
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   1
      Left            =   4620
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   0
      Left            =   4140
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   21
      Left            =   1920
      Top             =   2820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   23
      Left            =   2880
      Top             =   2820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   15
      Left            =   1440
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   17
      Left            =   2400
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   19
      Left            =   3360
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   11
      Left            =   1920
      Top             =   1500
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   13
      Left            =   2880
      Top             =   1500
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   5
      Left            =   1440
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   7
      Left            =   2400
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   9
      Left            =   3360
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   1
      Left            =   1920
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   3
      Left            =   2880
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   0
      Left            =   1440
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   2
      Left            =   2400
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   4
      Left            =   3360
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   6
      Left            =   1920
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   8
      Left            =   2880
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   10
      Left            =   1440
      Top             =   1500
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   12
      Left            =   2400
      Top             =   1500
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   14
      Left            =   3360
      Top             =   1500
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   16
      Left            =   1920
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   18
      Left            =   2880
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   20
      Left            =   1440
      Top             =   2820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   22
      Left            =   2400
      Top             =   2820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idName 
      Height          =   495
      Index           =   24
      Left            =   3360
      Top             =   2820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   2
      Left            =   5100
      Top             =   180
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   5
      Left            =   4620
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   7
      Left            =   5580
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   8
      Left            =   4140
      Top             =   1500
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   10
      Left            =   5100
      Top             =   1500
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   13
      Left            =   4620
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   15
      Left            =   5580
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   16
      Left            =   4140
      Top             =   2820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idScore 
      Height          =   495
      Index           =   18
      Left            =   5100
      Top             =   2820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   " "
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idNo 
      Height          =   495
      Index           =   1
      Left            =   720
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   "2"
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idNo 
      Height          =   495
      Index           =   2
      Left            =   720
      Top             =   1500
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   "3"
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idNo 
      Height          =   495
      Index           =   3
      Left            =   720
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   "4"
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin Sudoku.ID idNo 
      Height          =   495
      Index           =   4
      Left            =   720
      Top             =   2820
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   0   'False
      Caption         =   "5"
      ForeColor       =   65535
      HiLightColor    =   16777215
      DisabledColor   =   16777215
      Locked          =   -1  'True
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   5
      Left            =   60
      Top             =   3180
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00808080&
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   3660
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00E0E0E0&
      Height          =   195
      Index           =   4
      Left            =   60
      Top             =   3360
      Width           =   195
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Index           =   2
      Left            =   120
      Top             =   3480
      Width           =   315
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   1
      Left            =   300
      Top             =   3600
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   135
      Index           =   0
      Left            =   -540
      Top             =   3840
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   195
      Index           =   19
      Left            =   720
      Top             =   3180
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   18
      Left            =   720
      Top             =   3060
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   17
      Left            =   720
      Top             =   2940
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   435
      Index           =   15
      Left            =   720
      Top             =   2520
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   14
      Left            =   720
      Top             =   2400
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   13
      Left            =   720
      Top             =   2280
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   435
      Index           =   11
      Left            =   720
      Top             =   1860
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   10
      Left            =   720
      Top             =   1740
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   9
      Left            =   720
      Top             =   1620
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   435
      Index           =   7
      Left            =   720
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   6
      Left            =   720
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   5
      Left            =   720
      Top             =   960
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   435
      Index           =   2
      Left            =   720
      Top             =   540
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   1
      Left            =   720
      Top             =   420
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Index           =   0
      Left            =   720
      Top             =   300
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   195
      Index           =   3
      Left            =   720
      Top             =   120
      Width           =   5415
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "FTopScorer.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FTopScorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************
'* Title: FTopScorer                 *
'* Stamp: 19 July 2007               *
'* Auth : Derio                      *
'* Desc : Show Top Five Sudoku Mania *
'*************************************


Private Sub btnOK_Click()
'** Hide the form

  MSupport.FadeOut Me
End Sub

Private Sub tmrTransparent_Timer()
  Me.tmrTransparent.Enabled = False
  MSupport.FadeIn Me
End Sub

Public Sub InsertScore(ByVal Index As Integer, _
                       ByVal Name As String, _
                       ByVal Score As Integer)
'** insert score into score board

Dim I As Integer
Dim strTemp As String

  'insert the id (name)
  Name = Trim(Name)
  For I = 1 To 5
    Me.idName(I - 1 + (Index - 1) * 5).Caption = Mid(Name, I, 1)
  Next I
  
  'insert the score
  strTemp = Score
  For I = Len(strTemp) To 1 Step -1
    Me.idScore(3 - Len(strTemp) + I + (Index - 1) * 4).Caption = Mid(strTemp, I, 1)
  Next I
End Sub

Public Sub SetupHiLightIndex(ByVal Index As Integer)
'** Setup hilight index

Dim I As Integer

  Me.idNo(Index - 1).Enabled = True
  
  For I = 1 To 5
    Me.idName(I - 1 + (Index - 1) * 5).Enabled = True
  Next I

  For I = 1 To 4
    Me.idScore(I - 1 + (Index - 1) * 4).Enabled = True
  Next I
End Sub
