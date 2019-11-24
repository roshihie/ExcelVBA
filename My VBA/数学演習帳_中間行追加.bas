Option Explicit
'*******************************************************************************
'        数学演習帳  中間行追加
'*******************************************************************************
'      < 処理概要 >
'        数学演習帳に行間を空けるための中間行を追加する
'*******************************************************************************
Public Sub 数学演習帳_中間行追加()

  Const cnCol As Long = 56                                 'BD列
  Dim oCell As Range
  
  Application.ScreenUpdating = False
  
  Set oCell = ActiveCell
  Rows(oCell.Row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
  Rows(oCell.Row - 1).RowHeight = 7
  Cells(oCell.Row - 1, cnCol) = " "
  
  oCell.Activate

End Sub
