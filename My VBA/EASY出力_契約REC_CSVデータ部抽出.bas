Option Explicit
'*******************************************************************************
'        ファイル内容出力  データ部(データカンマ区切り部)抽出
'*******************************************************************************
'      < 処理概要 >
'        EASYPLUS で出力したファイル内容のデータ部(データカンマ区切り部)を
'        抽出し、その他の部分を削除する。
'*******************************************************************************
Public Sub 契約REC_CSVデータ部抽出  ()

  Const s契約DATA_HDR As String = "1*** ﾕｳｼﾏｽﾀｰ ｹｲﾔｸ REC"
  Dim oRow()    As tRowSet
  Dim oFirstRng As Range

  Application.ScreenUpdating = False
  
  Call sbデータ検索(oRow, oFirstRng, s契約DATA_HDR)
  If oFirstRng Is Nothing Then
    Exit Sub
  End If
  Call sbデータ抽出(oRow)
  
  Application.ScreenUpdating = True
  Cells(1, 1).Activate
  
End Sub

'*******************************************************************************
'        データ部(データカンマ区切り部)検索
'*******************************************************************************
'      < 処理概要 >
'        データ部(データカンマ区切り部)を検索し、位置を特定する。
'*******************************************************************************
Private Sub sbデータ検索(oRow() As tRowSet, oFirstRng As Range, sDATA_HDR As String)

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
'        データ部(データカンマ区切り部)抽出
'*******************************************************************************
'      < 処理概要 >
'        データ部(データカンマ区切り部)の特定された位置を基に抽出する。
'*******************************************************************************
Private Sub sbデータ抽出(oRow() As tRowSet)

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

