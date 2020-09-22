Attribute VB_Name = "MTop5SM"
Option Explicit
'****************************
'* Title: MTop5SM           *
'* Stamp: 23 July 2007      *
'* Auth : Derio             *
'* Desc : Top Scorer Module *
'****************************


Private Const T5SMMax = 5
Private Const T5SMFileLen = 1024
Private Const T5SMHeaderFile = "T5SM"
Private Const T5SMFileVersion = "001"
Private Const T5SMKeyLen = 5
Private Const T5SMInfoLen = 5
Private Const T5SMCheckDigitLen = 3

Private Type TopScorer
  Name As String
  Score As Integer
End Type
Private Top5ScorerList(1 To T5SMMax) As TopScorer
Private hT5SMFile As Integer
Private T5SMFileName As String


Public Function OpenTopScorerFile(ByVal FileName As String) As Integer
'** Open top five scorer file

Dim I As Integer
Dim strHeader As String
Dim strVersion As String
Dim strKey As String
Dim strLen As String
Dim strInfo As String
Dim strCheckDigit As String
Dim strTemp As String

Dim OK As Boolean

  OK = False
  T5SMFileName = FileName
  If Dir(T5SMFileName) <> "" Then
    hT5SMFile = FreeFile
    Open T5SMFileName For Binary Access Read As #hT5SMFile
    
    strHeader = Space(Len(T5SMHeaderFile))
    Get #hT5SMFile, , strHeader
    strVersion = Space(Len(T5SMFileVersion))
    Get #hT5SMFile, , strVersion
    strKey = Space(T5SMKeyLen)
    Get #hT5SMFile, , strKey
    strLen = Space(T5SMInfoLen)
    Get #hT5SMFile, , strLen
    strInfo = Space(Val(strLen))
    Get #hT5SMFile, , strInfo
    strTemp = Space(T5SMFileLen - (Val(strLen) + Len(T5SMHeaderFile) + Len(T5SMFileVersion) + T5SMKeyLen + T5SMInfoLen + T5SMCheckDigitLen))
    Get #hT5SMFile, , strTemp
    strCheckDigit = Space(T5SMCheckDigitLen)
    Get #hT5SMFile, , strCheckDigit
    
    If strHeader <> T5SMHeaderFile Then
      OpenTopScorerFile = 101 'invalid header
    
    ElseIf Not IsNumeric(strVersion) Then
      OpenTopScorerFile = 102 'version invalid
      
    ElseIf Val(strVersion) > Val(T5SMFileVersion) Then
      OpenTopScorerFile = 103 'version not supported
      
    ElseIf strCheckDigit <> CheckDigit(strInfo & strTemp, strKey) Then
      OpenTopScorerFile = 104 'invalid check digit
      
    Else
      strInfo = Scramble(strInfo, strKey)
      For I = 1 To T5SMMax
        With Top5ScorerList(I)
          .Name = GetToken(strInfo)
          strTemp = GetToken(strInfo)
          If Not IsNumeric(strTemp) Then
            .Score = 0
          
          Else
            .Score = Val(strTemp)
          End If
        End With
      Next I
      OK = True
    End If
    
  Else
    hT5SMFile = 0
    OpenTopScorerFile = 100 'file not found
  End If
  
  'default if no scorer file exist
  If Not OK Then
    For I = 1 To T5SMMax
      With Top5ScorerList(I)
        .Name = "DDA" & "-" & I
        .Score = 1100 - I * 200
      End With
    Next I
    
    CreateTopScorerFile
  End If
End Function

Private Function Scramble(ByVal strInfo As String, _
                          ByVal strKey As String) As String
'** Hide the info with xor method

Dim I As Integer
Dim J As Integer
Dim strTemp As String

  strTemp = ""
  For I = 1 To Len(strInfo)
    J = (J + 1) Mod Len(strKey)
    strTemp = strTemp & _
              Chr(Asc(Mid(strInfo, I, 1)) Xor Asc(Mid(strKey, J + 1, 1)))
  Next I
  Scramble = strTemp
End Function

Private Function CheckDigit(ByVal strInfo As String, _
                            ByVal strKey As String) As String
'** Create check digit for info

Dim Total As Long
Dim I As Integer
Dim J As Integer
Dim strTemp As String
Dim strCheckDigit As String

  Total = 0
  For I = 1 To Len(strInfo)
    J = (J + 1) Mod Len(strKey)
    Total = Total + CLng(Asc(Mid(strInfo, I, 1))) * CLng(Asc(Mid(strKey, J + 1, 1)))
  Next I
  strTemp = Total
  strTemp = Trim(strTemp)
  
  I = Len(strTemp)
  strCheckDigit = ""
  For J = 1 To T5SMCheckDigitLen
    strCheckDigit = strCheckDigit & Chr(100 + Mid(strTemp, I - J, 2))
  Next J
  CheckDigit = strCheckDigit
End Function

Private Function GetToken(Source As String, _
                          Optional Separator As String = "|") As String
'** Get token from Source string

Dim I As Integer

  I = InStr(Source, Separator)
  If I = 0 Then
    GetToken = Source
    Source = ""
  Else
    GetToken = Left(Source, I - 1)
    Source = Mid(Source, I + 1)
  End If
End Function

Private Function CreateRandomString(ByVal StringLen As Integer) As String
'** Create random string for camouflage purpose

Dim I As Integer
Dim strTemp As String

  Randomize
  strTemp = ""
  For I = 1 To StringLen
    strTemp = strTemp & Chr(Int(Rnd * 255))
  Next I
  CreateRandomString = strTemp
End Function

Public Function CreateTopScorerFile() As Integer
'** Create and save top score file

Dim I As Integer
Dim strKey As String
Dim strInfo As String
Dim strLen As String
Dim strTemp As String
Dim strCheckDigit As String

  'close the previous open file
  If hT5SMFile <> 0 Then
    On Local Error Resume Next
    Close #hT5SMFile
    On Error GoTo 0
  End If
  
  hT5SMFile = FreeFile
  If Dir(T5SMFileName) <> "" Then
    Err.Clear
    On Local Error Resume Next
    Kill T5SMFileName
    If Err <> 0 Then
      CreateTopScorerFile = 201 'can not delete old file
      On Local Error GoTo 0
      Exit Function
    End If
    On Local Error GoTo 0
  End If
  
  Open T5SMFileName For Binary Access Write As #hT5SMFile
  Put #hT5SMFile, , T5SMHeaderFile   'header file
  Put #hT5SMFile, , T5SMFileVersion  'version file
  
  strKey = CreateRandomString(T5SMKeyLen)
  strInfo = ""
  For I = 1 To T5SMMax
    With Top5ScorerList(I)
      strInfo = strInfo & "|" & .Name
      strTemp = .Score
      strInfo = strInfo & "|" & Trim(strTemp)
    End With
  Next I
  strInfo = Mid(strInfo, 2)
  strInfo = Scramble(strInfo, strKey)
  strLen = Len(strInfo)
  If Len(strLen) < T5SMInfoLen Then strLen = String(T5SMInfoLen - Len(strLen), "0") & strLen
  strTemp = CreateRandomString(T5SMFileLen - (Len(T5SMHeaderFile) + Len(T5SMFileVersion) + T5SMKeyLen + T5SMInfoLen + Len(strInfo) + T5SMCheckDigitLen))
  strCheckDigit = CheckDigit(strInfo & strTemp, strKey)
  
  Put #hT5SMFile, , strKey
  Put #hT5SMFile, , strLen
  Put #hT5SMFile, , strInfo
  Put #hT5SMFile, , strTemp
  Put #hT5SMFile, , strCheckDigit
  
  Err.Clear
  On Local Error Resume Next
  Close #hT5SMFile
  If Err <> 0 Then
    CreateTopScorerFile = 202 'can not create new file
  Else
    CreateTopScorerFile = 0 'sukses
    
    'protect file for other application access
    hT5SMFile = FreeFile
    Open T5SMFileName For Binary Access Read As #hT5SMFile
  End If
  On Local Error GoTo 0
End Function

Public Sub CloseTopScorerFile()
'** Close Top Scorer file

  If hT5SMFile <> 0 Then Close #hT5SMFile
End Sub


Public Function GetTopPos(ByVal Score As Integer) As Integer
'** Get position on top score
'   0 --> not on the list

Dim I As Integer

  For I = 1 To T5SMMax
    If Top5ScorerList(I).Score < Score Then
      GetTopPos = I
      Exit Function
    End If
  Next I
  GetTopPos = 0
End Function

Public Sub InsertNewScore(ByVal Name As String, _
                          ByVal Score As Integer, _
                          ByVal Index As Integer)
'** Insert new score into the top scorer list

Dim I As Integer

  For I = T5SMMax - 1 To Index Step -1
    Top5ScorerList(I + 1).Name = Top5ScorerList(I).Name
    Top5ScorerList(I + 1).Score = Top5ScorerList(I).Score
  Next I
  
  Top5ScorerList(Index).Name = Name
  Top5ScorerList(Index).Score = Score
End Sub

Public Sub ShowTopFiveSudokuMania(Optional HiLightIndex As Integer = 0)
'** Show the top five sudoku mania

Dim fTemp As FTopScorer
Dim I As Integer

  Set fTemp = New FTopScorer
  With fTemp
    For I = 1 To T5SMMax
      If Top5ScorerList(I).Score <> 0 Then
        .InsertScore I, Top5ScorerList(I).Name, Top5ScorerList(I).Score
      Else
        Exit For
      End If
    Next I
    
    If HiLightIndex <> 0 Then .SetupHiLightIndex HiLightIndex
    HideForm fTemp
  End With

  Unload fTemp
  Set fTemp = Nothing
End Sub
