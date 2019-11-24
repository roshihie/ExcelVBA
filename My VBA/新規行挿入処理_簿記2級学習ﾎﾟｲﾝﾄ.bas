Option Explicit

Dim isERR As Integer                                                ' ERROR �t���O

Public Type typCellPos                                              ' �Z���|�W�V�����^
    lngRow As Long
    lngCol As Long
End Type

Dim conItm          As Variant                                      ' �w�b�_�[�s�̏����ݒ�K�v�� �z��
Const conHeadItm    As String = "is,�� �� �� �t,�� �� �� �t,"       ' �w�b�_�[�s�̏����ݒ�K�v�� ����

Const conArrayMax   As Integer = 10                                 ' �w�b�_�[�s�̏����ݒ�K�v�� MAX��
Const conDiffItm    As Integer = 2                                  ' �w�b�_�[�s�Ɩ��׍s�̍�(�s��)
Const conRGBMax     As Integer = 255                                '�Ԋ|���F(RGB White�p)

'*******************************************************************************
'        �V �K �s �} �� �� ��
'*******************************************************************************
'        �����T�v�F����L�Q���w�K�|�C���g��.xlsm �ɂ����āA�}�������V�K�s��
'                  �Œ�l �܂��� Excel�֐����ݒ肳��Ă����� �����ݒ���s��
'
'                  �Œ�l �ݒ��  �F"is"
'                  Excel�֐��ݒ��F"�������t", "�Ώ����t"
'
'*******************************************************************************
Public Sub �V�K�s�}������()

    Dim posItm(conArrayMax) As typCellPos                           ' �����ݒ�K�v��̃Z���ʒu �z��
    Dim posHead             As typCellPos                           ' �w�b�_�[�s�̃Z���ʒu
                                                                    ' (�s�� ActiveCell�s�Œ�C��͏����ݒ�K�v��)
    Dim i                   As Integer
    
    Application.ScreenUpdating = False                              ' ��ʍX�V ��~
    Call ProcInit(posHead, posItm)                                  ' ��������
    Call ProcNewLine(posHead, posItm)                               ' ���ڕ⊮����

End Sub

'*******************************************************************************
'        ���@���@���@��
'*******************************************************************************
'        �����T�v�FExcel�֐��ݒ���z��Ɋi�[����
'
'            �߂�l�@�F�Ȃ�
'            �����P  �F�w�b�_�[�|�W�V����  typCellPos
'            �����Q  �F���ڃ|�W�V����      typCellPos
'*******************************************************************************
Private Sub ProcInit(posHead As typCellPos, _
                     posItm() As typCellPos)
                     
    Dim rngFind   As Range                                          ' FIND�֐��̃��^�[���l
    Dim lngCol    As Long                                           ' ���[�N��NO
    Dim intRtn    As Integer                                        ' MSGBOX�֐��̃��^�[���l
    
    Dim i, j      As Integer
    
    conItm = Split(conHeadItm, ",")                                 ' �w�b�_�[�s�̗񖼏̂�z��ɂ���
    
    i = 0
    Do While conItm(i) <> ""                                        ' �w�b�_�[�s�̗񖼏̔z��̑S�f�[�^���������s��
    
       Set rngFind = Cells.Find(conItm(i))                          ' �w�b�_�[�s�̗񖼏̔z���FIND����
       If rngFind Is Nothing Then                                   ' �w�b�_�[�s�̗񖼏̂��Ȃ��Ƃ�
          posItm(i).lngRow = 0
          posItm(i).lngCol = 0
       Else                                                         ' �w�b�_�[�s�̗񖼏̂�����Ƃ�
          posHead.lngRow = rngFind.Row                                ' �w�b�_�[�|�W�V����.�sNO �Ƀw�b�_�[�s �ݒ�
          posHead.lngCol = rngFind.Column                             ' �w�b�_�[�|�W�V����.��NO �Ƀw�b�_�[�s�̗񖼏̗̂� �ݒ�
          posItm(i).lngRow = ActiveCell.Row                           ' ���ڃ|�W�V����.�sNO ��ActiveCell�s �ݒ�
          posItm(i).lngCol = rngFind.Column                           ' ���ڃ|�W�V����.��NO �Ƀw�b�_�[�s�̗񖼏̗̂� �ݒ�
       End If
       i = i + 1
    Loop
    
End Sub

'*******************************************************************************
'        �V�K�s�}������
'*******************************************************************************
'        �����T�v�FActiveCell�s�ɑ΂��āA�V�K�s��ǉ����A�Œ�l �܂��� Excel�֐��ݒ��ɂ�
'                  ���l�̐ݒ���s��
'
'            �߂�l�@�F�Ȃ�
'            �����P  �F�w�b�_�[�|�W�V����  typCellPos
'            �����Q  �F���ڃ|�W�V����      typCellPos
'*******************************************************************************
Private Sub ProcNewLine(posHead As typCellPos, _
                        posItm() As typCellPos)

    Dim i    As Integer

    ActiveCell.EntireRow.Insert

    i = 0                                                           ' �����ݒ�K�v��̑S�f�[�^���������s��
    Do While (i <= conArrayMax And _
              posItm(i).lngCol <> 0)
       
       Select Case Cells(posHead.lngRow, posItm(i).lngCol).Value
          Case "is"                                                 ' "is"��
             Cells(posItm(i).lngRow, posItm(i).lngCol).Value = "1"    ' �Œ�l "1" �ݒ�
          Case "�� �� �� �t"                                        ' "�������t"
                                                                      ' �j���Z�o�֐� �ݒ�
             Cells(posItm(i).lngRow, posItm(i).lngCol).Offset(, 1).Formula = _
                "=IF(AH" & ActiveCell.Row & "<>"""",""("" & CHOOSE(WEEKDAY(DATE(YEAR(AH" & ActiveCell.Row & _
                "),MONTH(AH" & ActiveCell.Row & "),DAY(AH" & ActiveCell.Row & _
                ")),1),""��"",""��"",""��"",""��"",""��"",""��"",""�y"") & "")"","""")"
          Case "�� �� �� �t"                                        ' "�Ώ����t"
                                                                      ' �j���Z�o�֐� �ݒ�
             Cells(posItm(i).lngRow, posItm(i).lngCol).Offset(, 1).Formula = _
                "=IF(BJ" & ActiveCell.Row & "<>"""",""("" & CHOOSE(WEEKDAY(DATE(YEAR(BJ" & ActiveCell.Row & _
                "),MONTH(BJ" & ActiveCell.Row & "),DAY(BJ" & ActiveCell.Row & _
                ")),1),""��"",""��"",""��"",""��"",""��"",""��"",""�y"") & "")"","""")"
          
       End Select
       
       i = i + 1
    Loop

End Sub
