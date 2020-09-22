Attribute VB_Name = "MSupport"
Option Explicit
'****************************
'* Title: MSupport          *
'* Stamp: 15 June 2007      *
'* Auth : Derio             *
'* Desc : Supporting module *
'****************************


'** Basic supporting function
Public Type POINT
  X As Long
  Y As Long
End Type

Public Declare Sub Sleep _
  Lib "kernel32" (ByVal Milliseconds As Long)

Public Declare Function GetTempPath _
  Lib "kernel32" _
  Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                        ByVal lpBuffer As String) As Long
                        
Public Declare Function GetCursorPos _
  Lib "user32" (lpPoint As POINT) As Long

Private Declare Function GetWindowLong _
  Lib "user32" _
  Alias "GetWindowLongA" (ByVal hWnd As Long, _
                          ByVal nIndex As Long) As Long
                          
Private Declare Function SetLayeredWindowAttributes _
  Lib "user32" (ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
                
Private Declare Function SetWindowLong _
  Lib "user32" _
  Alias "SetWindowLongA" (ByVal hWnd As Long, _
                          ByVal nIndex As Long, _
                          ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2

Public Const MAX_OPACITY = 225





Public Function IsYes(ByVal Message As String, _
                      Optional CaptionYes As String = "Yes", _
                      Optional CaptionNo As String = "No") As Boolean
'** Show message box with question Yes or No

Dim fTemp As FQuestion
Dim actWidth As Single

  Set fTemp = New FQuestion
  With fTemp
    'setup the logo
    .Caption = " " & App.Title
    .lblMessage = Message
    .imgLogo(0).Visible = True
    .imgLogo(1).Visible = False
    .imgLogo(2).Visible = False
    
    'setup the button
    With .btnCommand(0)
      .Caption = CaptionNo
      .Top = fTemp.lblMessage.Top + fTemp.lblMessage.Height
      If .Top < fTemp.imgLogo(0).Top + fTemp.imgLogo(0).Height Then
        .Top = fTemp.imgLogo(0).Top + fTemp.imgLogo(0).Height
      End If
      
      .Top = .Top + 300
      .Visible = True
    End With
    
    With .btnCommand(1)
      .Caption = CaptionYes
      .Top = fTemp.btnCommand(0).Top
      .Visible = True
    End With
    
    'setup the message
    actWidth = .lblMessage.Width
    If .TextWidth(Message) < .lblMessage.Width Then
      .lblMessage.Width = .TextWidth(Message) + 60
      If .lblMessage.Width < 2 * .btnCommand(0).Width + 30 Then
        .lblMessage.Width = 2 * .btnCommand(0).Width + 30
      End If
    End If
    
    'setup acessories pos
    .shpAccessories(0).Top = .btnCommand(0).Top + 180
    .shpAccessories(1).Top = .shpAccessories(0).Top - 240
    .shpAccessories(2).Top = .shpAccessories(0).Top - 360
    .shpAccessories(3).Top = .shpAccessories(0).Top - 180
    .shpAccessories(4).Top = .shpAccessories(0).Top - 480
    .shpAccessories(5).Top = .shpAccessories(0).Top - 660
    
    'setup button pos
    actWidth = actWidth - .lblMessage.Width
    .btnCommand(0).Left = .btnCommand(0).Left - actWidth
    .btnCommand(1).Left = .btnCommand(1).Left - actWidth
    
    'setup the form
    .Width = .Width - actWidth
    .Height = .Height - .ScaleHeight + .btnCommand(0).Top + .btnCommand(0).Height + 60
    
    'show the form with fade-in effect
    HideForm fTemp
    
    'capture user act
    IsYes = (.Tag = "Yes")
  End With
  Unload fTemp
  Set fTemp = Nothing
End Function

Public Sub ShowInfo(ByVal Message As String, _
                    Optional Warning As Boolean = False)
'** Show message box just for information

Dim fTemp As FQuestion
Dim actWidth As Single

  Set fTemp = New FQuestion
  With fTemp
    'setup the logo
    .Caption = " " & App.Title
    .imgLogo(0).Visible = False
    If Not Warning Then
      .imgLogo(1).Visible = True
      .imgLogo(2).Visible = False
    Else
      .imgLogo(1).Visible = False
      .imgLogo(2).Visible = True
    End If
    
    .lblMessage = Message
    
    'setup the button
    With .btnCommand(0)
      .Caption = "OK"
      .Top = fTemp.lblMessage.Top + fTemp.lblMessage.Height
      If .Top < fTemp.imgLogo(0).Top + fTemp.imgLogo(0).Height Then
        .Top = fTemp.imgLogo(0).Top + fTemp.imgLogo(0).Height
      End If
      
      .Top = .Top + 300
      .Visible = True
    End With
    
    'setup the messege dimention
    actWidth = .lblMessage.Width
    If .TextWidth(Message) < .lblMessage.Width Then
      .lblMessage.Width = .TextWidth(Message) + 60
      If .lblMessage.Width < .btnCommand(0).Width + 30 Then
        .lblMessage.Width = .btnCommand(0).Width + 30
      End If
    End If
    
    'setup the form dimention
    actWidth = actWidth - .lblMessage.Width
    .Width = .Width - actWidth
    .btnCommand(0).Left = .btnCommand(0).Left - actWidth
    .btnCommand(1).Left = .btnCommand(1).Left - actWidth
    .btnCommand(1).Visible = False
    .Height = .Height - .ScaleHeight + .btnCommand(0).Top + .btnCommand(0).Height + 60
    
    'setup acessories pos
    .shpAccessories(0).Top = .btnCommand(0).Top + 180
    .shpAccessories(1).Top = .shpAccessories(0).Top - 240
    .shpAccessories(2).Top = .shpAccessories(0).Top - 360
    .shpAccessories(3).Top = .shpAccessories(0).Top - 180
    .shpAccessories(4).Top = .shpAccessories(0).Top - 480
    .shpAccessories(5).Top = .shpAccessories(0).Top - 660
    
    'show the form with fade-in effect
    HideForm fTemp
  End With
  
  Unload fTemp
  Set fTemp = Nothing
End Sub

Public Sub HideForm(MyForm As Form, _
                    Optional Modal As Boolean = True, _
                    Optional ParentForm As Form)
'** Show form with transparent effect

  MakeTransparent MyForm.hWnd, 0
  If Modal Then
    MyForm.Show vbModal
  Else
    If ParentForm Is Nothing Then
      MyForm.Show
    Else
      MyForm.Show , ParentForm
    End If
  End If
End Sub

Public Sub MakeTransparent(ByVal hWnd As Long, _
                           Opacity As Integer)
'** Make form transparent base on Opacity

Dim Msg As Long
    
  Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong hWnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes hWnd, 0, Opacity, LWA_ALPHA
End Sub

Public Sub FadeIn(MyForm As Form, Optional MaxOpacity = MAX_OPACITY)
'** Make form transparent from 0 to MAX_OPACITY

Dim Opacity As Integer

  MyForm.Enabled = False
  For Opacity = 0 To MaxOpacity
    MakeTransparent MyForm.hWnd, Opacity
    DoEvents
  Next Opacity
  MyForm.Enabled = True
End Sub

Public Sub FadeOut(MyForm As Form, _
                   Optional StartingPoint As Integer = MAX_OPACITY)
'** Make form transparent from MAX_OPACITY (base on starting point) to 0

Dim Opacity As Integer
  
  MyForm.Enabled = False
  For Opacity = StartingPoint To 0 Step -1
    MakeTransparent MyForm.hWnd, Opacity
    DoEvents
  Next Opacity
  MyForm.Hide
  MyForm.Enabled = True
End Sub

