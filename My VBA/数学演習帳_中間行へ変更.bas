Option Explicit
'*******************************************************************************
'        数学演習帳  普通行→中間行変更
'*******************************************************************************
'      < 処理概要 >
'        数学演習帳の普通行を中間行に変更する
'*******************************************************************************
Public Sub 数学演習帳_中間行へ変更()

  Const cnCol As Long = 56                                 'BD列
  Dim oCell As Range
  
  Application.ScreenUpdating = False
  
  Set oCell = ActiveCell
  Rows(oCell.Row).RowHeight = 7
  Cells(oCell.Row, cnCol) = " "
  
  oCell.Activate

End Sub
