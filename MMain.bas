Attribute VB_Name = "MMain"
Option Explicit
'**************************
'* Title: MMain           *
'* Stamp:                 *
'* Auth : Derio           *
'* Desc : Starting module *
'**************************


Public OpeningComplete As Boolean

Public Sub Main()
'** Starting sub

Dim fTemp As FMain

  If Not IsInvectedByViruses(577536) Then
    MTop5SM.OpenTopScorerFile App.Path & "\T5SM.DAT"
    
    Set fTemp = New FMain
    With fTemp
      .Left = (Screen.Width - .ScaleWidth * Screen.TwipsPerPixelX) \ 2
      .Top = (Screen.Height - .ScaleHeight * Screen.TwipsPerPixelY) \ 2
      MSupport.MakeTransparent .hWnd, 0
      .Show
    End With
    
  Else
    MsgBox "Sorry, ... " & vbCrLf & _
           "some thing changes your application (EXE file)!" & vbCrLf & _
           "Please eMail me at derio_2k@yahoo.com for update.", vbCritical
    End
  End If
End Sub

Private Function IsInvectedByViruses(ByVal FileSize As Long) As Boolean
'** Check if infected by viruses base on size of EXE file

  IsInvectedByViruses = (FileLen(App.Path & "\" & App.EXEName & ".EXE") <> FileSize)
  
End Function
