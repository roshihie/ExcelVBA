Option Explicit
'*******************************************************************************
'        新規行挿入処理
'*******************************************************************************
'      < 処理概要 >
'        英単語_熟語.xlsm の現在行に新規行を追加し、フォントの設定，関数の設定を
'        行う。
'*******************************************************************************
Sub NewLineSet()
                                                           ' 新規挿入ラインを初期化する
  Dim lLastCol   As Integer
  Dim oFindCell  As Range
                                                           ' 列の最終Column 取得（ヘッダー行）
  Set oFindCell = Cells.Find(What:="№", after:=Range("A1"))
                                                           ' 挿入行の｢№｣列に関数を設定
  ActiveCell.EntireRow.Insert Shift:=xlDown
  Cells(ActiveCell.Row, oFindCell.Column).Formula = _
    "=IF(LEN(B" & ActiveCell.Row & ")>1,ROW()-ROW(A$1)-COUNTBLANK(A$2:A" & ActiveCell.Row - 1 & "),"""")"
                                                           ' 挿入行の｢単語｣列のフォントを標準に設定
  With Cells(ActiveCell.Row, oFindCell.Column).Offset(, 1)
    .Font.Bold = False
    .Font.Italic = False
  End With
                                                           ' 挿入行全体は塗りつぶしなし
  lLastCol = oFindCell.End(xlToRight).Column
  Range(Cells(ActiveCell.Row, oFindCell.Column), _
        Cells(ActiveCell.Row, lLastCol)).Interior.ColorIndex = xlNone
                       ' 挿入行＋１行の№列の関数 再設定 (∵COUNTBLANK(A$2:Annn) の nnn が１アップしない)
                                                           ' 上記事象は、挿入行＋１行目のみ
  Cells(ActiveCell.Row + 1, oFindCell.Column).Formula = _
    "=IF(LEN(B" & ActiveCell.Row + 1 & ")>1,ROW()-ROW(A$1)-COUNTBLANK(A$2:A" & ActiveCell.Row & "),"""")"

  Cells(ActiveCell.Row, oFindCell.Column).Offset(, 1).Select
    
End Sub

