Option Explicit
'*******************************************************************************
'        ���w���K��  ���ʍs�����ԍs�ύX
'*******************************************************************************
'      < �����T�v >
'        ���w���K���̕��ʍs�𒆊ԍs�ɕύX����
'*******************************************************************************
Public Sub ���w���K��_���ԍs�֕ύX()

  Const cnCol As Long = 56                                 'BD��
  Dim oCell As Range
  
  Application.ScreenUpdating = False
  
  Set oCell = ActiveCell
  Rows(oCell.Row).RowHeight = 7
  Cells(oCell.Row, cnCol) = " "
  
  oCell.Activate

End Sub
