Option Explicit
'*******************************************************************************
'        �V�K�s�}������
'*******************************************************************************
'      < �����T�v >
'        �p�P��_�n��.xlsm �̌��ݍs�ɐV�K�s��ǉ����A�t�H���g�̐ݒ�C�֐��̐ݒ��
'        �s���B
'*******************************************************************************
Sub NewLineSet()
                                                           ' �V�K�}�����C��������������
  Dim lLastCol   As Integer
  Dim oFindCell  As Range
                                                           ' ��̍ŏIColumn �擾�i�w�b�_�[�s�j
  Set oFindCell = Cells.Find(What:="��", after:=Range("A1"))
                                                           ' �}���s�̢�����Ɋ֐���ݒ�
  ActiveCell.EntireRow.Insert Shift:=xlDown
  Cells(ActiveCell.Row, oFindCell.Column).Formula = _
    "=IF(LEN(B" & ActiveCell.Row & ")>1,ROW()-ROW(A$1)-COUNTBLANK(A$2:A" & ActiveCell.Row - 1 & "),"""")"
                                                           ' �}���s�̢�P�꣗�̃t�H���g��W���ɐݒ�
  With Cells(ActiveCell.Row, oFindCell.Column).Offset(, 1)
    .Font.Bold = False
    .Font.Italic = False
  End With
                                                           ' �}���s�S�͓̂h��Ԃ��Ȃ�
  lLastCol = oFindCell.End(xlToRight).Column
  Range(Cells(ActiveCell.Row, oFindCell.Column), _
        Cells(ActiveCell.Row, lLastCol)).Interior.ColorIndex = xlNone
                       ' �}���s�{�P�s�̇���̊֐� �Đݒ� (��COUNTBLANK(A$2:Annn) �� nnn ���P�A�b�v���Ȃ�)
                                                           ' ��L���ۂ́A�}���s�{�P�s�ڂ̂�
  Cells(ActiveCell.Row + 1, oFindCell.Column).Formula = _
    "=IF(LEN(B" & ActiveCell.Row + 1 & ")>1,ROW()-ROW(A$1)-COUNTBLANK(A$2:A" & ActiveCell.Row & "),"""")"

  Cells(ActiveCell.Row, oFindCell.Column).Offset(, 1).Select
    
End Sub

