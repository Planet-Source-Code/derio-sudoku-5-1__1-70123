VERSION 5.00
Begin VB.Form FMain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sudoku - The Hot Puzzle Craze"
   ClientHeight    =   6825
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6840
   Icon            =   "FMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6840
   Visible         =   0   'False
   Begin VB.PictureBox pctOption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   7020
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   3000
      Width           =   750
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   8
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   7
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   7
         Top             =   480
         Width           =   240
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   5
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   3
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   240
      End
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   20
      Left            =   1560
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   18
      Left            =   120
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   10
      Left            =   840
      Top             =   840
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   2
      Left            =   1560
      Top             =   120
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   0
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin VB.Timer tmrTransparent 
      Interval        =   1000
      Left            =   7020
      Top             =   1740
   End
   Begin VB.PictureBox pctMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   1680
      ScaleHeight     =   1005
      ScaleWidth      =   3465
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   1
         Left            =   30
         TabIndex        =   12
         Top             =   210
         Width           =   3315
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   855
         Left            =   60
         Top             =   60
         Width           =   3315
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   0
         Left            =   50
         TabIndex        =   11
         Top             =   240
         Width           =   3315
      End
   End
   Begin VB.Timer tmrOpening 
      Interval        =   100
      Left            =   7020
      Top             =   2400
   End
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6960
      Top             =   1020
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   3
      Left            =   2340
      Top             =   120
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   4
      Left            =   3060
      Top             =   120
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   5
      Left            =   3780
      Top             =   120
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   6
      Left            =   4560
      Top             =   120
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   7
      Left            =   5280
      Top             =   120
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   8
      Left            =   6000
      Top             =   120
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   9
      Left            =   120
      Top             =   840
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   11
      Left            =   1560
      Top             =   840
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   12
      Left            =   2340
      Top             =   840
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   13
      Left            =   3060
      Top             =   840
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   14
      Left            =   3780
      Top             =   840
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   15
      Left            =   4560
      Top             =   840
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   16
      Left            =   5280
      Top             =   840
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   17
      Left            =   6000
      Top             =   840
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   19
      Left            =   840
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   21
      Left            =   2340
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   22
      Left            =   3060
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   23
      Left            =   3780
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   24
      Left            =   4560
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   25
      Left            =   5280
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   26
      Left            =   6000
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   27
      Left            =   120
      Top             =   2340
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   28
      Left            =   840
      Top             =   2340
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   29
      Left            =   1560
      Top             =   2340
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   30
      Left            =   2340
      Top             =   2340
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   31
      Left            =   3060
      Top             =   2340
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   32
      Left            =   3780
      Top             =   2340
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   33
      Left            =   4560
      Top             =   2340
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   34
      Left            =   5280
      Top             =   2340
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   35
      Left            =   6000
      Top             =   2340
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   36
      Left            =   120
      Top             =   3060
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   37
      Left            =   840
      Top             =   3060
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   38
      Left            =   1560
      Top             =   3060
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   39
      Left            =   2340
      Top             =   3060
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   40
      Left            =   3060
      Top             =   3060
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   41
      Left            =   3780
      Top             =   3060
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   42
      Left            =   4560
      Top             =   3060
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   43
      Left            =   5280
      Top             =   3060
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   44
      Left            =   6000
      Top             =   3060
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   45
      Left            =   120
      Top             =   3780
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   46
      Left            =   840
      Top             =   3780
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   47
      Left            =   1560
      Top             =   3780
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   48
      Left            =   2340
      Top             =   3780
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   49
      Left            =   3060
      Top             =   3780
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   50
      Left            =   3780
      Top             =   3780
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   51
      Left            =   4560
      Top             =   3780
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   52
      Left            =   5280
      Top             =   3780
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   53
      Left            =   6000
      Top             =   3780
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   54
      Left            =   120
      Top             =   4560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   55
      Left            =   840
      Top             =   4560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   56
      Left            =   1560
      Top             =   4560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   57
      Left            =   2340
      Top             =   4560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   58
      Left            =   3060
      Top             =   4560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   59
      Left            =   3780
      Top             =   4560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   60
      Left            =   4560
      Top             =   4560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   61
      Left            =   5280
      Top             =   4560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   62
      Left            =   6000
      Top             =   4560
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   63
      Left            =   120
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   64
      Left            =   840
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   65
      Left            =   1560
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   66
      Left            =   2340
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   67
      Left            =   3060
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   68
      Left            =   3780
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   69
      Left            =   4560
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   70
      Left            =   5280
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   71
      Left            =   6000
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   72
      Left            =   120
      Top             =   6000
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   73
      Left            =   840
      Top             =   6000
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   74
      Left            =   1560
      Top             =   6000
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   75
      Left            =   2340
      Top             =   6000
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   76
      Left            =   3060
      Top             =   6000
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   77
      Left            =   3780
      Top             =   6000
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   2
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   78
      Left            =   4560
      Top             =   6000
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   79
      Left            =   5280
      Top             =   6000
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   80
      Left            =   6000
      Top             =   6000
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin Sudoku.Cell sucCell 
      Height          =   720
      Index           =   1
      Left            =   840
      Top             =   120
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Caption         =   ""
      Mode            =   1
      ForeColor       =   8454143
      ProtectedColor  =   4210752
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuTesting 
         Caption         =   "Testing"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuResolvePuzzle 
         Caption         =   "Resolve"
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo Move"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighScore 
         Caption         =   "Top Five Sudoku Mania"
      End
      Begin VB.Menu mnuRule 
         Caption         =   "How to ..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About ..."
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************
'* Title  : Sudoku 5.0       *
'* Author : Derio            *
'* Type   : Puzzle Game      *
'* Stamp  : 15 June 2007     *
'*****************************

Private Const MaxSudokuOption = 9
Private Const MaxTime = 3600 'one hour

Private CurrentIndex As Integer
Private CurrentSelection As Integer
Private ClearSudokuCell As Integer
Private Playing As Boolean
Private UndoStack As Collection
Private SudokuLib(4) As String

Private Enum GameLevelType
  Testing = 0
  Beginer = 1
  Intermediate = 2
  Advance = 3
  Professional = 4
End Enum
Private ActiveGameLevel As GameLevelType

Private Enum TimeKeeperType
  Starting = 0
  Continue = 1
  Pending = 2
  GetTime = 3
End Enum

Private Enum UndoEventType
  AddPenciling = 10
  AddPencilingWithCaption = 11
  ChoosePenciling = 12
  ClearPenciling = 13
  DirectChoose = 21
  DirectChooseWithNote = 22
  DirectClear = 23
End Enum

Private FormHeight As Integer


Private Sub Form_Load()
'** Start the game
  
Dim fTemp As FAbout

  'Show About Form as splash screen
  Set fTemp = New FAbout
  fTemp.Tag = "Opening"
  MSupport.HideForm fTemp, False, Me
  
  Select Case UCase(Command$)
  Case "DEBUG", "D"
    InitSudoku Testing
    
  Case "BEGINER", "B"
    InitSudoku Beginer
    
  Case "INTERMEDIATE", "I"
    InitSudoku Intermediate
    
  Case "ADVANCE", "A"
    InitSudoku Advance
    
  Case "PRO", "P"
    InitSudoku Professional
    
  Case Else
    InitSudoku Intermediate
  End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'** Check for user key pressed

  If KeyAscii = vbKeyEscape Then
    Me.WindowState = vbMinimized
  End If
End Sub

Private Sub Form_Resize()
'** Setup time line when windows state change
Dim I As Integer

  If Me.Tag = "Resize" Then Exit Sub
  If Playing Then
    If WindowState = vbMinimized Then
      TimeKeeper Pending
      
    ElseIf WindowState = vbNormal Then
      Me.Tag = "Resize"
      Me.Height = Me.Height - Me.ScaleHeight
      DoEvents
      Me.Height = Me.Height + 120
      DoEvents
      
      For I = 1 To MaxSudokuOption
        Me.Height = Me.Height + Me.sucCell(0).Height
        DoEvents
      Next I
      
      Me.Height = Me.Height + 240
      
      Me.Tag = ""
      TimeKeeper Continue
      Me.tmrGame.Enabled = True
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'** Confirm for Exit

  If Playing Then
    If Not MSupport.IsYes("Are you sure to exit?") Then
      Cancel = True
      Exit Sub
    End If
  End If
  
  MTop5SM.CloseTopScorerFile
  MSupport.FadeOut Me, 254
  MSupport.Sleep 100
  Unload FAbout
End Sub

Private Sub lblOption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'** Select option value

Dim I As Integer

  If Button = vbLeftButton Then
    '* apply selected item
    If Me.sucCell(CurrentIndex).Caption <> "" Then
      AddUndo DirectChoose, CurrentIndex, Me.sucCell(CurrentIndex).Caption
    Else
      AddUndo DirectChooseWithNote, CurrentIndex, Me.sucCell(CurrentIndex).GetNoteList()
    End If
    Me.sucCell(CurrentIndex).Caption = Val(Me.lblOption(Index).Caption)
    Me.pctOption.Tag = ""
    Me.pctOption.Visible = False
    
    If ClearSudokuCell = 0 Then
      If IsSudokuSolved() Then
        ShowSolvedMessage
      End If
    End If
  
  Else
    '* make a note
    Me.sucCell(CurrentIndex).AddNote Me.lblOption(Index).Caption
    If Me.sucCell(CurrentIndex).AddNoteSuccess Then
      AddUndo AddPenciling, CurrentIndex, Me.lblOption(Index).Caption
    End If
    
    If Me.pctOption.Tag <> "" Then
      GetUndo
      AddUndo AddPencilingWithCaption, CurrentIndex, Me.pctOption.Tag
      Me.pctOption.Tag = ""
    End If
    Me.pctOption.Visible = False
  End If
End Sub

Private Sub ShowSolvedMessage()
'** Show Solving Message

Dim fTemp As FRegister
Dim Score As Integer
Dim Name As String
Dim Index As Integer

  Me.tmrGame.Enabled = False
  Playing = False
  ProtectSudoku
  Score = TimeKeeper(GetTime)
  MSupport.ShowInfo "Congratulation, you just solve the puzzle " & _
                    "in " & ShowTotalTime(Score)
                    
  Score = MaxTime - Score
  If Score < 0 Then Exit Sub
  
  'Record the score to top five list
  Index = MTop5SM.GetTopPos(Score)
  If Index <> 0 Then
    Set fTemp = New FRegister
    MSupport.HideForm fTemp, True
    Name = fTemp.Tag
    Unload fTemp
    Set fTemp = Nothing
    
    If Name <> "" Then
      MTop5SM.InsertNewScore Name, Score, Index
      MTop5SM.CreateTopScorerFile
    End If
  End If
  
  MTop5SM.ShowTopFiveSudokuMania Index
End Sub

Private Sub mnuAbout_Click()
'** Show something about me :-)

Dim fTemp As FAbout

  Set fTemp = New FAbout
  With fTemp
    .Caption = " About ..."
    .cmdOK.Enabled = True
    MSupport.HideForm fTemp
  End With

  Unload fTemp
  Set fTemp = Nothing
End Sub

Private Sub mnuExit_Click()
'** Exit the game

  Unload Me
End Sub

Private Sub ClearSudokuBoard()
'** Clear all of the sudoku item

Dim BoxIndex As Integer
Dim ItemIndex As Integer
Dim CellIndex As Integer
Dim L As Integer
Dim ArrBox(1 To MaxSudokuOption) As Integer

  If Me.pctOption.Visible Then Me.pctOption.Visible = False
 
  ScrambleIndex ArrBox()

  For BoxIndex = 1 To MaxSudokuOption
    For ItemIndex = 1 To MaxSudokuOption
      CellIndex = GetCellBaseOfBox(ArrBox(BoxIndex), ArrBox(ItemIndex))
      With Me.sucCell(CellIndex)
        .Visible = False
        .Caption = ""
        If ArrBox(BoxIndex) Mod 2 = 0 Then
          .Mode = DarkButton
        Else
          .Mode = LightButton
        End If
        Sleep 10
        DoEvents
      End With
    Next ItemIndex
  Next BoxIndex
End Sub

Private Sub ShowSudokuBoard(Optional Opening As Boolean = False)
'** Show sudoku board item

Dim BoxIndex As Integer
Dim ItemIndex As Integer
Dim CellIndex As Integer
Dim ArrBox(1 To MaxSudokuOption) As Integer

  ScrambleIndex ArrBox()
  
  'show all sudoku item with animation
  For BoxIndex = 1 To MaxSudokuOption
    For ItemIndex = 1 To MaxSudokuOption
      CellIndex = GetCellBaseOfBox(ArrBox(BoxIndex), ArrBox(ItemIndex))
      With Me.sucCell(CellIndex)
        .Visible = True
        Sleep 10
        DoEvents
      End With
    Next ItemIndex
  Next BoxIndex

  If Not Opening Then ProtectSudoku
End Sub

Private Sub ScrambleIndex(ArrBox() As Integer)
'** Scaramble array for index

Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim L As Integer

  '* init box index
  For I = 1 To MaxSudokuOption
    ArrBox(I) = I
  Next I
  
  'scramble
  For I = 1 To MaxSudokuOption
    J = 1 + Int(Rnd * MaxSudokuOption)
    K = 1 + Int(Rnd * MaxSudokuOption)
    
    L = ArrBox(J)
    ArrBox(J) = ArrBox(K)
    ArrBox(K) = L
  Next I

End Sub

Private Sub ProtectSudoku()
'protect the cell that has number

Dim BoxIndex As Integer
Dim ItemIndex As Integer
Dim CellIndex As Integer
Dim ArrBox(1 To MaxSudokuOption) As Integer

  ScrambleIndex ArrBox()
  
  'protect sudoku cell if not empty
  For BoxIndex = MaxSudokuOption To 1 Step -1
    For ItemIndex = MaxSudokuOption To 1 Step -1
      CellIndex = GetCellBaseOfBox(ArrBox(BoxIndex), ArrBox(ItemIndex))
      With Me.sucCell(CellIndex)
        If .Caption <> "" Then
          .Mode = Protected
        End If
        Sleep 10
        DoEvents
      End With
    Next ItemIndex
  Next BoxIndex
End Sub

Private Function IsSudokuSolved() As Boolean
'** Check sudoku board status

Dim I As Integer
Dim Row As Integer
Dim Col As Integer
Dim Box As Integer
Dim Index As Integer
Dim strTemp As String

  IsSudokuSolved = False
  
  'check for the row rule
  For Row = 1 To MaxSudokuOption
    strTemp = ""
    For Col = 1 To MaxSudokuOption
      Index = GetCell(Col, Row)
      strTemp = strTemp & Me.sucCell(Index).Caption
    Next Col
    
    'check if number 1 to 9 exist on the row
    For I = 1 To MaxSudokuOption
      If InStr(strTemp, I) = 0 Then
        Exit Function
      End If
    Next I
  Next Row
  
  'check for the col rule
  For Col = 1 To MaxSudokuOption
    strTemp = ""
    For Row = 1 To MaxSudokuOption
      Index = GetCell(Col, Row)
      strTemp = strTemp & Me.sucCell(Index).Caption
    Next Row
    
    'check if number 1 to 9 exist on the col
    For I = 1 To MaxSudokuOption
      If InStr(strTemp, I) = 0 Then
        Exit Function
      End If
    Next I
  Next Col
  
  'check for the box rule
  For Box = 1 To MaxSudokuOption
    strTemp = ""
    For I = 1 To MaxSudokuOption
      Index = GetCellBaseOfBox(Box, I)
      strTemp = strTemp & Me.sucCell(Index).Caption
    Next I
    
    'check if number 1 to 9 exist on the col
    For I = 1 To MaxSudokuOption
      If InStr(strTemp, I) = 0 Then
        Exit Function
      End If
    Next I
  Next Box
    
  IsSudokuSolved = True
End Function

Private Sub LoadLibrary(ByVal Level As GameLevelType)
'** Load selected library problem base on given index

Dim strTemp As String
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim Col As Integer
Dim Row As Integer
Dim RotationCode As Integer
Dim MirroringCode As Integer
Dim HSwitchCode As Integer
Dim VSwitchCode As Integer
Dim NumberOfShift As Integer
Dim tmpNumber As Integer
Dim fTemp As FLayer

  'create fade effect when removing sudoku item
  Set fTemp = New FLayer
  With fTemp
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight - 150
    .Left = Me.Left + (Me.Width - Me.ScaleWidth) \ 2
    .Top = Me.Top + Me.Height - Me.ScaleHeight
    MSupport.MakeTransparent fTemp.hWnd, 0
    .Show , Me
    .tmrShow.Enabled = True
  End With
  
  Me.Show
  ClearSudokuBoard
  Do
    DoEvents
  Loop Until Not fTemp.tmrShow.Enabled
  
  'init scrambles value
  RotationCode = Int(Rnd * 4)
  MirroringCode = Int(Rnd * 3)
  NumberOfShift = Int(Rnd * MaxSudokuOption)
  HSwitchCode = Int(Rnd * 7)
  VSwitchCode = Int(Rnd * 7)
  ClearSudokuCell = MaxSudokuOption * MaxSudokuOption
  
  strTemp = SudokuLib(Level)
  For J = 1 To Len(strTemp)
  
    'rotating
    Select Case RotationCode
    Case 1 'rotate 90 deg clock wise
      Col = GetColumn(J - 1)
      Row = (MaxSudokuOption + 1) - GetRow(J - 1)
      I = GetCell(Row, Col) + 1
      
    Case 2 'rotate 180 deg
      Col = (MaxSudokuOption + 1) - GetColumn(J - 1)
      Row = (MaxSudokuOption + 1) - GetRow(J - 1)
      I = GetCell(Col, Row) + 1
    
    Case 3 'rotate 90 deg counter clock wise
      Col = (MaxSudokuOption + 1) - GetColumn(J - 1)
      Row = GetRow(J - 1)
      I = GetCell(Row, Col) + 1
      
    Case Else 'no rotation
      I = J
    End Select
    
    'mirroring
    Select Case MirroringCode
      Case 1 'vertical mirror
        Col = (MaxSudokuOption + 1) - GetColumn(I - 1)
        Row = GetRow(I - 1)
        I = GetCell(Col, Row) + 1
        
      Case 2 'horizontal mirror
        Col = GetColumn(I - 1)
        Row = (MaxSudokuOption + 1) - GetRow(I - 1)
        I = GetCell(Col, Row) + 1
      
      Case Else 'no mirrror
    End Select
    
    'shift number
    tmpNumber = Val(Mid(strTemp, J, 1))
    If tmpNumber <> 0 Then
      tmpNumber = tmpNumber + NumberOfShift
      If tmpNumber > MaxSudokuOption Then
        tmpNumber = tmpNumber - MaxSudokuOption
      End If
    End If
    
    If tmpNumber <> 0 Then
      Me.sucCell(I - 1).Caption = tmpNumber
    Else
      Me.sucCell(I - 1).Caption = ""
    End If
  Next J
  
  'switch row
  Select Case HSwitchCode
  Case 1 'switch row 1 and 2
    SwitchSudokuRow 1, 2
    SwitchSudokuRow 9, 8
    
  Case 2 'switch row 1 and 3
    SwitchSudokuRow 1, 3
    SwitchSudokuRow 9, 7
    
  Case 3 'switch row 2 and 3
    SwitchSudokuRow 2, 3
    SwitchSudokuRow 8, 7
    
  Case 4 'switch row 1 and 2 and 3
    SwitchSudokuRow 1, 2
    SwitchSudokuRow 9, 8
    SwitchSudokuRow 1, 3
    SwitchSudokuRow 9, 7
    
  Case 5 'switch row 3 and 2 and 1
    SwitchSudokuRow 1, 2
    SwitchSudokuRow 9, 8
    SwitchSudokuRow 2, 3
    SwitchSudokuRow 8, 7
    
  Case 6 'switch row 4 and 6
    SwitchSudokuRow 4, 6
  End Select
    
  'switch columns
  Select Case VSwitchCode
  Case 1 'switch col 1 and 2
    SwitchSudokuCol 1, 2
    SwitchSudokuCol 9, 8
    
  Case 2 'switch col 1 and 3
    SwitchSudokuCol 1, 3
    SwitchSudokuCol 9, 7
    
  Case 3 'switch col 2 and 3
    SwitchSudokuCol 2, 3
    SwitchSudokuCol 8, 7
    
  Case 4 'switch col 1 and 2 and 3
    SwitchSudokuCol 1, 2
    SwitchSudokuCol 9, 8
    SwitchSudokuCol 1, 3
    SwitchSudokuCol 9, 7
  
  Case 5 'switch col 3 and 2 and 1
    SwitchSudokuCol 1, 2
    SwitchSudokuCol 9, 8
    SwitchSudokuCol 2, 3
    SwitchSudokuCol 8, 7
  
  Case 6 'switch col 4 and 6
    SwitchSudokuCol 4, 6
  End Select

  fTemp.tmrHide.Enabled = True
  ShowSudokuBoard
End Sub

Private Sub SwitchSudokuRow(ByVal Row1 As Integer, ByVal Row2 As Integer)
'** Switch sudoku row, change Row1 with Row2 and vise versa

Dim I As Integer
Dim tmpNumber As String
Dim tmpMode As SUDOKU_MODE

  For I = 1 To MaxSudokuOption
    tmpNumber = Me.sucCell(GetCell(I, Row1)).Caption
    tmpMode = Me.sucCell(GetCell(I, Row1)).Mode
    Me.sucCell(GetCell(I, Row1)).Caption = Me.sucCell(GetCell(I, Row2)).Caption
    Me.sucCell(GetCell(I, Row1)).Mode = Me.sucCell(GetCell(I, Row2)).Mode
    Me.sucCell(GetCell(I, Row2)).Caption = tmpNumber
    Me.sucCell(GetCell(I, Row2)).Mode = tmpMode
  Next I
  DoEvents
End Sub

Private Sub SwitchSudokuCol(ByVal Col1 As Integer, ByVal Col2 As Integer)
'** Switch sudoku row, change Col1 with Col2 and vise versa

Dim I As Integer
Dim tmpNumber As String
Dim tmpMode As SUDOKU_MODE

  For I = 1 To MaxSudokuOption
    tmpNumber = Me.sucCell(GetCell(Col1, I)).Caption
    tmpMode = Me.sucCell(GetCell(Col1, I)).Mode
    Me.sucCell(GetCell(Col1, I)).Caption = Me.sucCell(GetCell(Col2, I)).Caption
    Me.sucCell(GetCell(Col1, I)).Mode = Me.sucCell(GetCell(Col2, I)).Mode
    Me.sucCell(GetCell(Col2, I)).Caption = tmpNumber
    Me.sucCell(GetCell(Col2, I)).Mode = tmpMode
  Next I
  DoEvents
End Sub

Private Sub mnuHighScore_Click()
'** Show top five sudoku mania

  ShowTopFiveSudokuMania
End Sub

Private Sub mnuNewGame_Click()
'** Play sudoku

Dim I As Integer

  If Playing Then
    If Not MSupport.IsYes("Are you sure to play a new Sudoku Puzzle " & _
                          "and cancel the current one?") Then Exit Sub
            
  End If
  
  Me.Enabled = False
  Me.tmrGame.Enabled = False
  Me.Caption = App.Title
  DoEvents
  
  Playing = True
  ClearUndoStack
  LoadLibrary ActiveGameLevel
  
  'show some thing to keep your passion
  For I = 5 To 1 Step -1
    ShowMessage "Countdown: " & I
    MSupport.Sleep 1000
  Next I
  ShowMessage "GO !!!"
  Sleep 1000
  TimeKeeper Starting
  
  HideMessage
  Me.Enabled = True
End Sub

Private Sub mnuRule_Click()
'** Show how to play Sudoku

Dim fTemp As Form
Dim State As Integer

  State = 0
  Do
    Select Case State
    Case 0
      Set fTemp = New FHowTo1 'general rule
    Case 1
      Set fTemp = New FHowTo2 'filling the box
    Case 2
      Set fTemp = New FHowTo3 'penciling
    Case 3
      Set fTemp = New FHowTo4 'using penciling notes
    End Select
    HideForm fTemp
    
    Select Case fTemp.Tag
    Case "Next"
      State = State + 1
    Case "Prev"
      State = State - 1
    Case Else
      State = 9
    End Select
    Unload fTemp
    Set fTemp = Nothing
  Loop Until State = 9
End Sub

Private Sub mnuUndo_Click()
'** Undo last action

  UndoLastAction
End Sub

Private Sub UndoLastAction()
'** Execute undo

Dim strTemp As String
Dim LastCommand As UndoEventType
Dim Index As Integer

  strTemp = GetUndo()
  LastCommand = CInt(Left(strTemp, 2))
  Index = CInt(Mid(strTemp, 4, 2))
  strTemp = Mid(strTemp, 7)
  Select Case LastCommand
    Case UndoEventType.DirectChoose
      Me.sucCell(Index).Caption = strTemp
    
    Case UndoEventType.DirectChooseWithNote
      With Me.sucCell(Index)
        .ClearNote
        .Caption = ""
        While strTemp <> ""
          Me.sucCell(Index).AddNote Left(strTemp, 1)
          strTemp = Mid(strTemp, 2)
        Wend
      End With
      
    Case UndoEventType.DirectClear
      Me.sucCell(Index).Caption = strTemp
    
    Case UndoEventType.AddPenciling
      Me.sucCell(Index).RemoveNote strTemp
      
    Case UndoEventType.AddPencilingWithCaption
      Me.sucCell(Index).Caption = strTemp
      
    Case UndoEventType.ChoosePenciling
      With Me.sucCell(Index)
        .Caption = ""
        While strTemp <> ""
          .AddNote Left(strTemp, 1)
          strTemp = Mid(strTemp, 2)
        Wend
      End With
      
    Case UndoEventType.ClearPenciling
      Me.sucCell(Index).AddNote strTemp
  End Select
End Sub

Private Sub sucCell_CaptionChange(Index As Integer, ByVal LastCaption As String)
'** Mark the cell if some thing changed

  If LastCaption = "" Then
    ClearSudokuCell = ClearSudokuCell - 1
  ElseIf sucCell(Index).Caption = "" Then
    ClearSudokuCell = ClearSudokuCell + 1
  End If
End Sub

Private Sub sucCell_LeftClick(Index As Integer)
'** Show the option to fill in the box

Dim ArrSudoku(MaxSudokuOption) As Integer
Dim I As Integer

  If Not Playing Then Exit Sub
  
  If Me.pctOption.Visible Then
    If Me.pctOption.Tag <> "" Then
      Me.sucCell(CurrentIndex).Caption = Val(Me.pctOption.Tag)
    End If
  End If
  CurrentIndex = Index
  
  With pctOption
    CurrentSelection = 0
    .Left = Me.sucCell(Index).Left
    .Top = Me.sucCell(Index).Top
    .Tag = ""
    
    ' enable all of the options
    For I = 1 To MaxSudokuOption
      With lblOption(I - 1)
        If .BackColor = vbYellow Then
          .BackColor = .ForeColor
          .ForeColor = vbYellow
        End If
        .Enabled = True
        .FontBold = .Enabled
      End With
    Next I
    
    '* Hilight the default (some thing you chose before)
    If Me.sucCell(Index).Caption <> "" Then
      With lblOption(Me.sucCell(Index).Caption - 1)
        .Enabled = True
        .ForeColor = .BackColor
        .BackColor = vbYellow
        .FontBold = True
      End With
      Me.pctOption.Tag = Me.sucCell(Index).Caption
    End If
  
    If Not .Visible Then .Visible = True
  End With
    
End Sub

Private Sub sucCell_NoteClick(Index As Integer, ByVal LastCaption As String, ByVal NoteList As String)
'** Apply selected note for Sudoku Cell

Dim MinIndex As Integer

  AddUndo ChoosePenciling, Index, LastCaption & NoteList
  If Me.pctOption.Visible Then Me.pctOption.Visible = False
  
  If ClearSudokuCell = 0 Then
    If IsSudokuSolved() Then
      ShowSolvedMessage
    End If
  End If
End Sub

Private Sub sucCell_NoteRemove(Index As Integer, ByVal Note As String)
'** remove the selected note --> Keep the history for undo purpose

  AddUndo ClearPenciling, Index, Note
End Sub

Private Sub sucCell_RightClick(Index As Integer)
'** Clear sudoku cell

Dim MinIndex As Integer

  If Not Playing Then Exit Sub
  
  If Me.sucCell(Index).Caption <> "" Then
    AddUndo DirectClear, Index, Me.sucCell(Index).Caption
    Me.sucCell(Index).Caption = ""
  End If
End Sub

Private Sub tmrOpening_Timer()
'** Show animation Sudoku item for opening screen
'   This timer active after 10 ms (design property setup)

  tmrOpening.Enabled = False
  
  'Setup Form Main dimension
  Me.Height = Me.Height - Me.ScaleHeight + _
              Me.sucCell(Me.sucCell.Count - 1).Top + Me.sucCell(Me.sucCell.Count - 1).Height _
              + 120
  Me.Width = Me.Width - Me.ScaleWidth + _
             Me.sucCell(Me.sucCell.Count - 1).Left + Me.sucCell(Me.sucCell.Count - 1).Width _
             + 120
  Me.Refresh
  DoEvents
  FormHeight = Me.Height
  
  'Setup Form Main position
  Me.Left = (Screen.Width - Me.Width) \ 2
  Me.Top = (Screen.Height - Me.Height) \ 2
  DoEvents
  
  ShowSudokuBoard True
  OpeningComplete = True
End Sub

Private Sub tmrGame_Timer()
'** Show the progress time when playing random puzzle
  
Dim TimeLeft As Integer
Dim CellClear As Integer

  TimeLeft = MaxTime - TimeKeeper(GetTime)
  CellClear = MaxSudokuOption * MaxSudokuOption - ClearSudokuCell
  
  Caption = " Sudoku - " & _
            CellClear & " Cell" & IIf(CellClear > 1, "s", "") & " clear" & _
            ", " & IIf(TimeLeft <= 0, "<time out>", "time left: " & ShowTotalTime(TimeLeft))
End Sub

Private Function TimeKeeper(ByVal Mode As TimeKeeperType) As Integer
'** Get the total duration playing time

Static Duration As Integer
Static tmpTime As Single

Dim H As Integer
Dim M As Integer
Dim S As Integer
Dim strTemp As String

  Select Case Mode
  Case TimeKeeperType.Starting  'start game
    Duration = 0
    tmpTime = Timer
    Me.tmrGame.Enabled = True

  Case TimeKeeperType.Continue  'continue game after minimize
    tmpTime = Timer
    If Not Me.tmrGame.Enabled Then Me.tmrGame.Enabled = True
  
  Case TimeKeeperType.Pending
    If Me.tmrGame.Enabled Then Me.tmrGame.Enabled = False

  Case Else 'playing
    Duration = Duration + Timer - tmpTime
    tmpTime = Timer
    If Playing Then
      If Not Me.tmrGame.Enabled Then Me.tmrGame.Enabled = True
    End If
  End Select
  
  TimeKeeper = Duration
End Function

Private Function ShowTotalTime(ByVal Duration As Integer) As String
'** Show the total duration playing time

Dim H As Integer
Dim M As Integer
Dim S As Integer
Dim strTemp As String

  If Duration <= 0 Then Duration = 0
  
  H = Duration \ 3600
  M = (Duration - H * 3600) \ 60
  S = Duration - H * 3600 - M * 60
    
  strTemp = ""
  If H > 0 Then
    If H > 1 Then
      strTemp = H & " hours "
    Else
      strTemp = H & " hour "
    End If
  End If
  
  If M > 0 Then
    If M > 1 Then
      strTemp = strTemp & M & " mins "
    Else
      strTemp = strTemp & M & " min "
    End If
  End If
  
  If S > 0 Then
    If S > 1 Then
      strTemp = strTemp & S & " secs"
    Else
      strTemp = strTemp & S & " sec"
    End If
  End If
    
  ShowTotalTime = strTemp
End Function

Private Sub InitSudoku(ByVal InitialLevel As GameLevelType)
'** Init sudoku game

  Randomize
  ClearUndoStack
  
  '*Testing
  SudokuLib(GameLevelType.Testing) = _
              "002906400" & _
              "003857600" & _
              "760000095" & _
              "381405769" & _
              "070389040" & _
              "249601853" & _
              "130000076" & _
              "006714200" & _
              "007503900"
  
  '*Beginer
  SudokuLib(GameLevelType.Beginer) = _
              "002000300" & _
              "063120080" & _
              "700030060" & _
              "290700000" & _
              "400000006" & _
              "000006059" & _
              "040010005" & _
              "070053410" & _
              "005000600"

  '* Intermediate
  SudokuLib(GameLevelType.Intermediate) = _
              "700000800" & _
              "060030010" & _
              "000409350" & _
              "004008100" & _
              "380000096" & _
              "007600400" & _
              "079501000" & _
              "030070080" & _
              "006000002"
  
  '* Advance
  SudokuLib(GameLevelType.Advance) = _
              "605001020" & _
              "004050000" & _
              "009200010" & _
              "801700050" & _
              "003000700" & _
              "060003901" & _
              "090005200" & _
              "000040800" & _
              "040800105"
  
  '* Professional
  SudokuLib(GameLevelType.Professional) = _
              "200500470" & _
              "040008001" & _
              "300704800" & _
              "000002700" & _
              "600000002" & _
              "009600000" & _
              "001403007" & _
              "400200050" & _
              "052009003"

  ActiveGameLevel = InitialLevel
End Sub

Private Function GetBox(ByVal CellIndex As Integer) As Integer
'** Get the box index base on cell index

  GetBox = ((GetRow(CellIndex) - 1) \ 3) * 3 + ((GetColumn(CellIndex) - 1) \ 3) + 1
End Function

Private Function GetRow(ByVal CellIndex As Integer) As Integer
'** Get the row base on cell index

  GetRow = (CellIndex \ MaxSudokuOption) + 1
End Function

Private Function GetColumn(ByVal CellIndex As Integer) As Integer
'** Get the column base on cell index

  GetColumn = (CellIndex Mod MaxSudokuOption) + 1
End Function

Private Function GetCell(ByVal Column As Integer, ByVal Row As Integer) As Integer
'** Get the cell index base on Column and Row

  GetCell = (Row - 1) * MaxSudokuOption + Column - 1
End Function

Private Function GetCellBaseOfBox(ByVal BoxIndex As Integer, ByVal ItemIndex As Integer) As Integer
'** Get the CellIndex base on BoxIndex and ItemIndex
'   Every Box has nine item

  GetCellBaseOfBox = ItemIndex + _
                     ((ItemIndex - 1) \ 3) * 6 + _
                     (BoxIndex - 1) * 3 + _
                     ((BoxIndex - 1) \ 3) * 18 - 1
End Function

Private Sub ShowMessage(ByVal Message As String)
'** Show message

  Me.lblMessage(0).Caption = Message
  Me.lblMessage(1).Caption = Me.lblMessage(0).Caption
  If Not Me.pctMessage.Visible Then Me.pctMessage.Visible = True
  DoEvents
End Sub

Private Sub HideMessage()
'** Hide message

  Me.pctMessage.Visible = False
  DoEvents
End Sub

Private Sub ClearUndoStack()
'** Clear Undo Stack

  Set UndoStack = Nothing
  Set UndoStack = New Collection
  Me.mnuUndo.Enabled = False
End Sub

Private Sub AddUndo(MyCommand As UndoEventType, Index As Integer, Info As String)
'** Add undo history

Dim strTemp As String

  strTemp = MyCommand & "-" & Format(Index, "00") & "-" & Info
  UndoStack.Add strTemp
  If Not Me.mnuUndo.Enabled Then Me.mnuUndo.Enabled = True
End Sub

Private Function GetUndo() As String
'** Get undo information from stack

  GetUndo = UndoStack.Item(UndoStack.Count)
  UndoStack.Remove UndoStack.Count
  If UndoStack.Count = 0 Then Me.mnuUndo.Enabled = False
End Function

Private Sub tmrTransparent_Timer()
'** Show fade in effect

Static Opacity As Integer

  tmrTransparent.Interval = 1
  Opacity = Opacity + 2
  If Opacity < 255 Then
    MSupport.MakeTransparent Me.hWnd, Opacity
  Else
    tmrTransparent.Enabled = False
  End If
End Sub








'** These functions were designed for testing Sudoku puzzle

Private Sub mnuResolvePuzzle_Click()
Dim Finished As Boolean
Dim StepCount As Long
Dim I As Integer

  If Not MSupport.IsYes("Are you sure to let me solve this puzzle?") Then Exit Sub
  
  Playing = False
  Caption = App.Title
  Me.tmrGame.Enabled = False
  
  Finished = False
  ResolveUsingMin Finished, StepCount
  
  If Finished Then
    MSupport.ShowInfo "Puzzle resolved with " & StepCount & " steps!"
    
  Else
    If MSupport.IsYes("I'm sorry, I can't resove this puzzle" & vbCrLf & _
                      "Do you want me to try resolve this puzzle from begining?") Then
      While UndoStack.Count > 0
        UndoLastAction
        MSupport.Sleep 100
        DoEvents
      Wend
      
      Finished = False
      StepCount = 0
      ResolveUsingMin Finished, StepCount
      
      If Finished Then
        MSupport.ShowInfo "Puzzle resolved with " & StepCount & " steps!"
      
      Else
        MSupport.ShowInfo "I'm sorry, I can't resolve this puzzle", True
      End If
    End If
  End If
  
End Sub

Private Function CheckSudoku(MinIndex As Integer) As Boolean
'** Check sudoku board
'   if MinIndex = -1 --> Thre's some sudoku were invisible
'   if MinIndex = -2 --> The game is over

Dim I As Integer
Dim arrTemp(1 To MaxSudokuOption) As Integer
Dim MinOption As Integer
Dim CurrentOption As Integer
Dim ClearCount As Integer

  CheckSudoku = True
  MinIndex = -1
  MinOption = MaxSudokuOption + 1
  For I = 1 To Me.sucCell.Count
    If Me.sucCell(I - 1) = "" Then
      CurrentOption = GetSudokuItemList(I - 1, arrTemp())
      If CurrentOption = 0 Then
        Me.sucCell(I - 1).Visible = False
        MinIndex = -1
        CheckSudoku = False
        
      Else
        Me.sucCell(I - 1).Visible = True
        If CurrentOption < MinOption Then
          MinOption = CurrentOption
          MinIndex = I - 1
        End If
      End If
      
    Else
      ClearCount = ClearCount + 1
    End If
  Next I
  
  '* Mark if no more option on Sudoku Table
  If ClearCount = Me.sucCell.Count Then
    MinIndex = -2
  End If
End Function

Private Sub ResolveUsingMin(Finished As Boolean, StepCount As Long, Optional RollBack As Boolean)
'** Resolve the puzzle using minimum option method (recursive)
'   Just select the item with minimum option and make a choice

Dim I As Integer
Dim arrTemp(1 To MaxSudokuOption) As Integer
Dim MinIndex As Integer
Dim tmpIndex As Integer

  Sleep 100
  DoEvents
  StepCount = StepCount + 1
  '* get the indext that has minimum option
  CheckSudoku MinIndex
  If MinIndex < 0 Then
    If MinIndex = -2 Then Finished = True
    Exit Sub
  End If
  
  '* get the option of the seleted index
  GetSudokuItemList MinIndex, arrTemp()
  For I = 1 To MaxSudokuOption
    If arrTemp(I) = 1 Then
      '* guesing the item is the right one
      Me.sucCell(MinIndex) = I
      DoEvents
      
      '* resolve the rest using the same method
      ResolveUsingMin Finished, StepCount, RollBack
      
      '* undo the choice if not finished the sudoku puzzle
      If Not Finished Or RollBack Then
        Me.sucCell(MinIndex) = ""
        
        If Finished And RollBack Then Exit Sub
        
      Else
        Exit Sub
      End If
    
      '* clear sudoku state (before changing)
      CheckSudoku tmpIndex
    End If
  Next I
End Sub

Private Function GetSudokuItemList(ByVal Index As Integer, ArraySudoku() As Integer) As Integer
'** Get the possible Sudoku item

Dim I As Integer
Dim J As Integer
Dim strTemp As String
Dim EnableItemCount As Integer

  '* init array sudoku
  EnableItemCount = MaxSudokuOption
  For I = 1 To MaxSudokuOption
    ArraySudoku(I) = 1
  Next I
  
  '* disable the option if the same one at the same column has choosen
  J = GetColumn(Index)
  For I = 1 To MaxSudokuOption
    strTemp = Me.sucCell(GetCell(J, I))
    If strTemp <> "" Then
      ArraySudoku(Val(strTemp)) = 0
      EnableItemCount = EnableItemCount - 1
    End If
  Next I
  
  '* disable the option if the same one at the same row has choosen
  J = GetRow(Index)
  For I = 1 To MaxSudokuOption
    strTemp = Me.sucCell(GetCell(I, J))
    If strTemp <> "" Then
      If ArraySudoku(Val(strTemp)) <> 0 Then
        ArraySudoku(Val(strTemp)) = 0
        EnableItemCount = EnableItemCount - 1
      End If
    End If
  Next I
  
  '* disable the option if the same one at the same box has choosen
  J = GetBox(Index)
  For I = 1 To MaxSudokuOption
    strTemp = Me.sucCell(GetCellBaseOfBox(J, I))
    If strTemp <> "" Then
      If ArraySudoku(Val(strTemp)) <> 0 Then
        ArraySudoku(Val(strTemp)) = 0
        EnableItemCount = EnableItemCount - 1
      End If
    End If
  Next I
  
  GetSudokuItemList = EnableItemCount
End Function
