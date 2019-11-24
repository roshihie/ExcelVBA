Option Explicit
'*******************************************************************************
'        ��������  �R�s�[
'*******************************************************************************
'      < �����T�v >
'        �R�s�[���Ŏw�肳�ꂽ�Z���͈͂̕����t�H���g������
'        �R�s�[��Ŏw�肳�ꂽ�Z���͈͂փR�s�[����
'*******************************************************************************
Public Sub ��������_�R�s�[()
  
  Const sSOCPRPT As String = "�R�s�[��"
  Const sSOCTITL As String = "�R�s�[��"
  Const sTRGPRPT As String = "�R�s�[��"
  Const sTRGTITL As String = "�R�s�[��"
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
'        �Z���͈͓���
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
'        �ΏۃZ���͈�  �e�[�u���ݒ�
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
