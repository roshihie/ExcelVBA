Option Explicit
'*******************************************************************************
'        wK   Ôs¨ÊsÏX
'*******************************************************************************
'      < Tv >
'        wK ÌÔsðÊsðÏX·é
'*******************************************************************************
Public Sub wK _ÊsÖÏX()

  Const cnCol As Long = 56                                 'BDñ
  Dim oCell As Range
  
  Application.ScreenUpdating = False
  
  Set oCell = ActiveCell
  Rows(oCell.Row).RowHeight = 15
  Cells(oCell.Row, cnCol) = " "
  
  oCell.Activate

End Sub

