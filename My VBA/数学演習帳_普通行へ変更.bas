Option Explicit
'*******************************************************************************
'        ���w���K��  ���ԍs�����ʍs�ύX
'*******************************************************************************
'      < �����T�v >
'        ���w���K���̒��ԍs�𕁒ʍs��ύX����
'*******************************************************************************
Public Sub ���w���K��_���ʍs�֕ύX()

  Const cnCol As Long = 56                                 'BD��
  Dim oCell As Range
  
  Application.ScreenUpdating = False
  
  Set oCell = ActiveCell
  Rows(oCell.Row).RowHeight = 15
  Cells(oCell.Row, cnCol) = " "
  
  oCell.Activate

End Sub

