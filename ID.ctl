VERSION 5.00
Begin VB.UserControl ID 
   CanGetFocus     =   0   'False
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   585
   ScaleHeight     =   540
   ScaleWidth      =   585
   ToolboxBitmap   =   "ID.ctx":0000
   Begin VB.Timer tmrHiLight 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   420
      Top             =   540
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   -15
      TabIndex        =   1
      Top             =   -15
      Width           =   495
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image imgBack 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "ID.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgBack 
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   0
      Picture         =   "ID.ctx":2606
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgBack 
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   0
      Picture         =   "ID.ctx":497C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "ID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'******************************
'* Title  : ID                *
'* Type   : ActiveX OCX       *
'* Author : Derio             *
'* Stamp  : 18 July 2007      *
'* Desc   : UI for Scorer ID  *
'******************************

Private Const CharTable = "ABCDEFGHIJKLMNOPQRSTUVWXYZ -'^*#@~&1234567890<"

Private MouseOver As Boolean

Private vEnabled As Boolean
Private vForeColor As OLE_COLOR
Private vHiLightColor As OLE_COLOR
Private vLocked As Boolean
Private vDisabledColor As OLE_COLOR

Public Event Click()


Private Sub lblCaption_Click(Index As Integer)
  If vEnabled And Not vLocked Then
    RaiseEvent Click
  End If
End Sub

Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not MouseOver Then
    If vEnabled And Not vLocked Then HiLight
  End If
End Sub

Private Sub tmrHiLight_Timer()
Dim Cur As POINT

  GetCursorPos Cur
  Cur.X = Cur.X - (Extender.Parent.Left + (Extender.Parent.Width - Extender.Parent.ScaleWidth) / 2) \ Screen.TwipsPerPixelX
  Cur.Y = Cur.Y - (Extender.Parent.Top + Extender.Parent.Height - Extender.Parent.ScaleHeight - 30) \ Screen.TwipsPerPixelY
  
  Cur.X = Cur.X * Screen.TwipsPerPixelX - Extender.Left
  Cur.Y = Cur.Y * Screen.TwipsPerPixelY - Extender.Top
  If Not (Cur.X >= 0 And Cur.X <= UserControl.Width _
          And Cur.Y >= 0 And Cur.Y <= UserControl.Height) Then
    tmrHiLight.Enabled = False
    HiLight False
    MouseOver = False
  End If
End Sub

Private Sub UserControl_Initialize()
  vForeColor = vbWhite
  vHiLightColor = vbWhite
  vDisabledColor = RGB(&H40, &H40, &H40)
  vEnabled = True
  vLocked = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    Me.Enabled = .ReadProperty("Enabled", True)
    Me.Caption = .ReadProperty("Caption", " ")
    Me.ForeColor = .ReadProperty("ForeColor", vbWhite)
    Me.HiLightColor = .ReadProperty("HiLightColor", vbWhite)
    Me.DisabledColor = .ReadProperty("DisabledColor", RGB(&H40, &H40, &H40))
    Me.Locked = .ReadProperty("Locked", False)
  End With
End Sub

Private Sub UserControl_Resize()
  With UserControl
    .Width = .imgBack(0).Width
    .Height = .imgBack(0).Height
  End With
End Sub

Private Sub HiLight(Optional DoHiLight As Boolean = True)
  With UserControl
    If DoHiLight Then
      MouseOver = True
      .imgBack(0).Visible = True
      .imgBack(1).Visible = False
      .imgBack(2).Visible = False
      .lblCaption(1).ForeColor = vHiLightColor
      If Not UserControl.tmrHiLight.Enabled Then UserControl.tmrHiLight.Enabled = True
         
    Else
      .imgBack(0).Visible = False
      .imgBack(1).Visible = True
      .imgBack(2).Visible = False
      .lblCaption(1).ForeColor = vForeColor
    End If
  End With
  
  PropertyChanged "HiLight"
End Sub

Public Property Get Enabled() As Boolean
  Enabled = vEnabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
  vEnabled = vNewValue
  With UserControl
    If Not vEnabled Then
      .imgBack(2).Visible = True
      .imgBack(0).Visible = False
      .imgBack(1).Visible = False
      .lblCaption(1).Visible = False
      .lblCaption(0).ForeColor = vDisabledColor
      
    Else
      .imgBack(0).Visible = False
      .imgBack(1).Visible = True
      .imgBack(2).Visible = False
      With .lblCaption(1)
        .ForeColor = vForeColor
        If Not .Visible Then .Visible = True
      End With
      .lblCaption(0).ForeColor = RGB(&H40, &H40, &H40)
    End If
  End With
  
  PropertyChanged "Enabled"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "Enabled", vEnabled
    .WriteProperty "Caption", UserControl.lblCaption(1).Caption
    .WriteProperty "ForeColor", vForeColor
    .WriteProperty "HiLightColor", vHiLightColor
    .WriteProperty "DisabledColor", vDisabledColor
    .WriteProperty "Locked", vLocked
  End With
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0
  Caption = UserControl.lblCaption(0).Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
  If Len(vNewValue) <> 1 Then Exit Property
  If InStr(CharTable, UCase(vNewValue)) = 0 Then Exit Property
  
  With UserControl
    .lblCaption(0).Caption = UCase(vNewValue)
    .lblCaption(1).Caption = UCase(vNewValue)
  End With
  
  PropertyChanged "Caption"
End Property

Public Property Get HiLightColor() As OLE_COLOR
  HiLightColor = vHiLightColor
End Property

Public Property Let HiLightColor(ByVal vNewValue As OLE_COLOR)
  vHiLightColor = vNewValue
  PropertyChanged "HiLightColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = vForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
  vForeColor = vNewValue
  If vEnabled Then
    UserControl.lblCaption(1).ForeColor = vForeColor
  End If
  
  PropertyChanged "ForeColor"
End Property

Public Property Get Locked() As Boolean
  Locked = vLocked
End Property

Public Property Let Locked(ByVal vNewValue As Boolean)
  vLocked = vNewValue
  
  PropertyChanged "Locked"
End Property

Public Property Get DisabledColor() As OLE_COLOR
  DisabledColor = vDisabledColor
End Property

Public Property Let DisabledColor(ByVal vNewValue As OLE_COLOR)
  vDisabledColor = vNewValue
  If Not Enabled Then
    UserControl.lblCaption(0).ForeColor = vDisabledColor
  End If
  PropertyChanged "DisabledColor"
End Property
