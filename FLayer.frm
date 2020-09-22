VERSION 5.00
Begin VB.Form FLayer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   1680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   300
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   180
      Top             =   240
   End
End
Attribute VB_Name = "FLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'****************************************
'* Title: FLayer                        *
'* Stamp: 5 July 2007                   *
'* Auth : Derio                         *
'* Desc : Cover the main screen to make *
'*        fade in / out effects         *
'****************************************


Private Sub tmrHide_Timer()
'** Hide the form (so you can see the other one behind)
'   with fade out effect

Static Opacity As Integer

  Opacity = Opacity + 2
  If Opacity < 255 Then
    MSupport.MakeTransparent Me.hWnd, 255 - Opacity
    
  Else
    Me.tmrHide.Enabled = False
    Unload Me
  End If
End Sub

Private Sub tmrShow_Timer()
'** Show the form with fade in effect

Static Opacity As Integer

  Opacity = Opacity + 2
  If Opacity < 255 Then
    MSupport.MakeTransparent Me.hWnd, Opacity
    
  Else
    Me.tmrShow.Enabled = False
  End If
End Sub
