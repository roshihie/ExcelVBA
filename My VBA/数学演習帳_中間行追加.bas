Option Explicit
'*******************************************************************************
'        ���w���K��  ���ԍs�ǉ�
'*******************************************************************************
'      < �����T�v >
'        ���w���K���ɍs�Ԃ��󂯂邽�߂̒��ԍs��ǉ�����
'*******************************************************************************
Public Sub ���w���K��_���ԍs�ǉ�()

  Const cnCol As Long = 56                                 'BD��
  Dim oCell As Range
  
  Application.ScreenUpdating = False
  
  Set oCell = ActiveCell
  Rows(oCell.Row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
  Rows(oCell.Row - 1).RowHeight = 7
  Cells(oCell.Row - 1, cnCol) = " "
  
  oCell.Activate

End Sub
