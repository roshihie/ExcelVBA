Option Explicit
'*******************************************************************************
'        wK   Ês¨ÔsÏX
'*******************************************************************************
'      < Tv >
'        wK ÌÊsðÔsÉÏX·é
'*******************************************************************************
Public Sub wK _ÔsÖÏX()

  Const cnCol As Long = 56                                 'BDñ
  Dim oCell As Range
  
  Application.ScreenUpdating = False
  
  Set oCell = ActiveCell
  Rows(oCell.Row).RowHeight = 7
  Cells(oCell.Row, cnCol) = " "
  
  oCell.Activate

End Sub
