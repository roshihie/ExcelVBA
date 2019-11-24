Option Explicit
'*******************************************************************************
'        �t�@�C�����e�o��  �f�[�^��(�f�[�^�J���}��؂蕔)���o
'*******************************************************************************
'      < �����T�v >
'        EASYPLUS �ŏo�͂����t�@�C�����e�̃f�[�^��(�f�[�^�J���}��؂蕔)��
'        ���o���A���̑��̕������폜����B
'*******************************************************************************
Public Sub �_��REC_CSV�f�[�^�����o  ()

  Const s�_��DATA_HDR As String = "1*** ճ�Ͻ�� ��Ը REC"
  Dim oRow()    As tRowSet
  Dim oFirstRng As Range

  Application.ScreenUpdating = False
  
  Call sb�f�[�^����(oRow, oFirstRng, s�_��DATA_HDR)
  If oFirstRng Is Nothing Then
    Exit Sub
  End If
  Call sb�f�[�^���o(oRow)
  
  Application.ScreenUpdating = True
  Cells(1, 1).Activate
  
End Sub

'*******************************************************************************
'        �f�[�^��(�f�[�^�J���}��؂蕔)����
'*******************************************************************************
'      < �����T�v >
'        �f�[�^��(�f�[�^�J���}��؂蕔)���������A�ʒu����肷��B
'*******************************************************************************
Private Sub sb�f�[�^����(oRow() As tRowSet, oFirstRng As Range, sDATA_HDR As String)

  Const sDATA_BODY    As String = "^([^,]+,)+$"
  Dim oRegExp  As RegExp
  Dim oFoudRng As Range
  Dim bFound   As Boolean
  Dim bBegin   As Boolean
  Dim bEnd     As Boolean
  Dim nCurRow  As Long
  Dim nidx     As Long

  Set oRegExp = CreateObject("VBScript.RegExp")
  oRegExp.Pattern = sDATA_BODY
  oRegExp.Global = False
  bFound = True

  Set oFoudRng = Cells.Find(What:=sDATA_HDR, LookAt:=xlPart)
  If oFoudRng Is Nothing Then
    bFound = False
  Else
    Set oFirstRng = oFoudRng
  End If
  
  nidx = 0
  Do While (bFound = True)
    ReDim Preserve oRow(nidx)
    nCurRow = oFoudRng.Row + 1
    bBegin = False
    Do While (bBegin = False)
      If oRegExp.Test(Cells(nCurRow, 1).Value) Then
        oRow(nidx).nBgnRow = nCurRow
        oRow(nidx).nEndRow = nCurRow
        bBegin = True
      End If
      nCurRow = nCurRow + 1
    Loop
        
    bEnd = False
    Do While (bEnd = False)
      If oRegExp.Test(Cells(nCurRow, 1).Value) Then
        oRow(nidx).nEndRow = nCurRow
      Else
        bEnd = True
      End If
      nCurRow = nCurRow + 1
    Loop
    
    Set oFoudRng = Cells.FindNext(oFoudRng)
    If oFoudRng.Address = oFirstRng.Address Then
      bFound = False
    Else
      nidx = nidx + 1
    End If
  Loop
  
End Sub

'*******************************************************************************
'        �f�[�^��(�f�[�^�J���}��؂蕔)���o
'*******************************************************************************
'      < �����T�v >
'        �f�[�^��(�f�[�^�J���}��؂蕔)�̓��肳�ꂽ�ʒu����ɒ��o����B
'*******************************************************************************
Private Sub sb�f�[�^���o(oRow() As tRowSet)

  Dim nidx     As Long
  Dim nLastRow As Long
  Dim nDelBgn  As Long
  Dim nDelEnd  As Long

  nidx = UBound(oRow)
  nDelBgn = oRow(nidx).nEndRow + 1
  nLastRow = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
  If nDelBgn < nLastRow Then
    Range(Rows(nDelBgn), Rows(nLastRow)).Delete
  End If

  Do While (nidx > LBound(oRow))
    nDelEnd = oRow(nidx).nBgnRow - 1
    nDelBgn = oRow(nidx - 1).nEndRow + 1
    Range(Rows(nDelBgn), Rows(nDelEnd)).Delete
    nidx = nidx - 1
  Loop

  nDelEnd = oRow(nidx).nBgnRow - 1
  If nDelEnd > 1 Then
    Range(Rows(1), Rows(nDelEnd)).Delete
  End If

End Sub

