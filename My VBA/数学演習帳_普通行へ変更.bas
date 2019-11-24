Option Explicit
'*******************************************************************************
'        数学演習帳  中間行→普通行変更
'*******************************************************************************
'      < 処理概要 >
'        数学演習帳の中間行を普通行を変更する
'*******************************************************************************
Public Sub 数学演習帳_普通行へ変更()

  Const cnCol As Long = 56                                 'BD列
  Dim oCell As Range
  
  Application.ScreenUpdating = False
  
  Set oCell = ActiveCell
  Rows(oCell.Row).RowHeight = 15
  Cells(oCell.Row, cnCol) = " "
  
  oCell.Activate

End Sub

