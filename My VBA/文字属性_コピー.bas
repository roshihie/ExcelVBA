Option Explicit
'*******************************************************************************
'        文字属性  コピー
'*******************************************************************************
'      < 処理概要 >
'        コピー元で指定されたセル範囲の文字フォント属性を
'        コピー先で指定されたセル範囲へコピーする
'*******************************************************************************
Public Sub 文字属性_コピー()
  
  Const sSOCPRPT As String = "コピー元"
  Const sSOCTITL As String = "コピー元"
  Const sTRGPRPT As String = "コピー先"
  Const sTRGTITL As String = "コピー先"
  Dim oSocRnges  As Range
  Dim oTrgRnges  As Range
  Dim aoSocRng() As Range
  Dim aoTrgRng() As Range
  Dim nSocCnt    As Integer
  Dim nTrgCnt    As Integer
  Dim nxTrg      As Integer
  Dim nxMod      As Integer
  Dim nxChar     As Integer
  
  On Error Resume Next
  
OrgCell:
  If fnRange_Inp(oSocRnges, sSOCPRPT, sSOCTITL) = 1 Then
    Exit Sub
  End If
    
  ReDim aoSocRng(oSocRnges.Count)
  Call sbTable_Set(aoSocRng, oSocRnges, nSocCnt)
  
  If fnRange_Inp(oTrgRnges, sTRGPRPT, sTRGTITL) = 1 Then
    GoTo OrgCell
  End If
  
  ReDim aoTrgRng(oTrgRnges.Count)
  Call sbTable_Set(aoTrgRng, oTrgRnges, nTrgCnt)
  
  Application.ScreenUpdating = False
  
  For nxTrg = 0 To (nTrgCnt - 1)
  
    nxMod = nxTrg Mod nSocCnt
    
    For nxChar = 1 To aoTrgRng(nxTrg).Characters.Count
      With aoTrgRng(nxTrg).Characters(Start:=nxChar, Length:=1).Font
        .ColorIndex = aoSocRng(nxMod).Characters(Start:=nxChar, Length:=1).Font.ColorIndex
        .Bold = aoSocRng(nxMod).Characters(Start:=nxChar, Length:=1).Font.Bold
        .Underline = aoSocRng(nxMod).Characters(Start:=nxChar, Length:=1).Font.Underline
      End With
    Next
      
  Next
  
End Sub

'*******************************************************************************
'        セル範囲入力
'*******************************************************************************
Private Function fnRange_Inp(oRnges As Range, sPrpt As String, sTitl As String)

  fnRange_Inp = 0
  Err.Number = 0
  Set oRnges = Application.InputBox(Prompt:=sPrpt, Title:=sTitl, Type:=8)

  If Err.Number > 0 Then
    fnRange_Inp = 1
  End If

End Function

'*******************************************************************************
'        対象セル範囲  テーブル設定
'*******************************************************************************
Private Sub sbTable_Set(aoRng() As Range, oRnges As Range, nMaxCnt As Integer)

  Const sJYOGAI As String = ",JYOGAI"
  Dim asJYOGAI() As String
  Dim oRng       As Range
  Dim nxRng      As Integer
  Dim nxJgi      As Integer
  
  asJYOGAI = Split(sJYOGAI, ",")
  
  nxRng = 0
  For Each oRng In oRnges
    
    For nxJgi = 0 To UBound(asJYOGAI)
      If oRng.Value = asJYOGAI(nxJgi) Then
        Exit For
      End If
    Next
    If nxJgi > UBound(asJYOGAI) Then
      Set aoRng(nxRng) = oRng
      nxRng = nxRng + 1
    End If
  Next
  
  nMaxCnt = nxRng

End Sub
