Option Explicit

Dim isERR As Integer                                       ' ERROR �t���O

Public Type typCellPos                                     ' �Z���|�W�V�����^
    lngRow As Long
    lngCol As Long
End Type

Dim conItm          As Variant                             ' �w�b�_�[�s�̗�(���׍s���� �Ԋ|�F�ύX��)�z��
                                                           ' �w�b�_�[�s�̗�(���׍s���� �Ԋ|�F�ύX��)����
Const conHeadItm    As String = "��,��,��@���@��,���@���@��,�T�@�@�@ �v,�� �� �� �t,�� �� �� �t,"

Const conArrayMax   As Integer = 10                        ' �w�b�_�[�s�̗� MAX��
Const conDiffItm    As Integer = 2                         ' �w�b�_�[�s�Ɩ��׍s�̍�(�s��)
Const conRGBRed     As Integer = 191                       '�Ԋ|�F(RGB Red)
Const conRGBGreen   As Integer = 191                       '�Ԋ|�F(RGB Green)
Const conRGBBlue    As Integer = 191                       '�Ԋ|�F(RGB Blue)

'*******************************************************************************
'        �r�����q�������`�������� �Ԋ|�ݒ�
'*******************************************************************************
'        �����T�v�F�ꗗ�\�� ActiveCell�s ���܂� ���߂̏�r���C���r���ň͂܂ꂽ
'                  RectAngle ��Ԋ|����
'                  ���̂Ƃ��A�w�b�_�[�s�̎w���̖��׍s(���׍s���� �Ԋ|�F�ݒ��)�ɑ΂���
'                  ��r�������݂��Ȃ��s�̕����� �Ԋ|�F�ɐݒ肷��
'
'*******************************************************************************
Public Sub �r����RectAngle�Ԋ|�ݒ�()

    Dim posStart            As typCellPos                  ' �����Ώۂ�START�Z���ʒu (�s��START�s�C���"is"�s�Œ�)
    Dim posEnd              As typCellPos                  ' �����Ώۂ�END�Z���ʒu   (�s��END�s�C���"is"�s�Œ�)
    Dim posItm(conArrayMax) As typCellPos                  ' ���׍s���� �Ԋ|�F�ݒ�� �z��
                                                           ' (�s��START�s�C��͕����F�ݒ��)
    
    Dim i                   As Integer
    
    Application.ScreenUpdating = False                     ' ��ʍX�V ��~
    Call ProcInit(posStart, posEnd, posItm)                ' ��������
    
    Call ProcInteriorSet(posStart, posEnd)                 ' �Ԋ|�ݒ�
    
    Do While (posItm(0).lngRow <= posEnd.lngRow)           ' ���׍s���� �Ԋ|�F�ݒ�s �� �����Ώ�END�s �̊� �������s��
    
       Call ProcFontSet(posEnd, posItm)                    ' ���׍s���� �Ԋ|�F�ݒ�
    Loop

End Sub

'*******************************************************************************
'        ���@���@���@��
'*******************************************************************************
'        �����T�v�F�����Ώۂ�START�s�CEND�s���m�肵�A���׍s���� �Ԋ|�F�ݒ���z��Ɋi�[����
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�J�n�|�W�V����   typCellPos
'            �����Q�@�F�I���|�W�V����   typCellPos
'            �����R  �F���ڃ|�W�V����   typCellPos
'*******************************************************************************
Private Sub ProcInit(posStart As typCellPos, _
                     posEnd As typCellPos, _
                     posItm() As typCellPos)
                     
    Dim rngFind   As Range                                 ' FIND�֐��̃��^�[���l
    Dim lngCol    As Long                                  ' ���[�N��NO
    Dim intRtn    As Integer                               ' MSGBOX�֐��̃��^�[���l
    
    Dim i, j      As Integer
    
    conItm = Split(conHeadItm, ",")                        ' �w�b�_�[�s�̗񖼏̂�z��ɂ���
    
    Set rngFind = Cells.Find("is")                         ' is��(�S�s "1" ���ߍ��ݗ�) ����
    If rngFind Is Nothing Then                             ' is�� ���Ȃ��Ƃ�
       intRtn = MsgBox(prompt:="is��(�S�s ""1"" ���ߍ��ݗ�) ��������܂���", Buttons:=vbOKOnly + vbCritical)
       If intRtn = vbOK Then
          isERR = True
          Exit Sub
       End If
    End If
    
    posStart.lngCol = Cells(rngFind.Row, 1).End(xlToRight).Column   ' is��̍ō���NO ���擾
    posStart.lngRow = ActiveCell.Row                                ' ActiveCell�s�̒��߂̏�r���s �擾
    Do While (Cells(posStart.lngRow, posStart.lngCol).Borders(xlEdgeTop).LineStyle = xlLineStyleNone)
    
       posStart.lngRow = posStart.lngRow - 1
    Loop
                                                           ' is��̍ŉE��NO ���擾
    posEnd.lngCol = Cells(rngFind.Row, ActiveSheet.Columns.Count).End(xlToLeft).Column
    posEnd.lngRow = ActiveCell.Row                                  ' ActiveCell�s�̒��߂̉��r���s �擾
    Do While (Cells(posEnd.lngRow, posEnd.lngCol).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone)
    
       posEnd.lngRow = posEnd.lngRow + 1
    Loop
    
    i = 0
    Do While conItm(i) <> ""                               ' �w�b�_�[�s�̗񖼏̔z��̑S�f�[�^���������s��
    
       Set rngFind = Cells.Find(conItm(i))                 ' �w�b�_�[�s�̗񖼏̔z���FIND����
       If rngFind Is Nothing Then                          ' �w�b�_�[�s�̗񖼏̂��Ȃ��Ƃ�
          posItm(i).lngRow = 0
          posItm(i).lngCol = 0
       Else                                                ' �w�b�_�[�s�̗񖼏̂�����Ƃ�
          posItm(i).lngCol = rngFind.Column                ' ��NO �Ƀw�b�_�[�s�̗񖼏̗̂� �ݒ�
          posItm(i).lngRow = posStart.lngRow               ' �sNO �Ƀw�b�_�[�s�̗񖼏̖̂��׍s �ݒ�
       End If
       i = i + 1
    Loop
    
End Sub

'*******************************************************************************
'        �r�����q�������`�������� �Ԋ|�ݒ�
'*******************************************************************************
'        �����T�v�FActiveCell�s ���܂� ���߂̌r���ň͂܂ꂽ RectAngle �̖Ԋ|���s��
'                  �Ԋ|�J���[�FRGB(191, 191, 191)
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�J�n�|�W�V����   typCellPos
'            �����Q�@�F�I���|�W�V����   typCellPos
'*******************************************************************************
Private Sub ProcInteriorSet(posStart As typCellPos, _
                            posEnd As typCellPos)
    
    Range(Cells(posStart.lngRow, posStart.lngCol), _
          Cells(posEnd.lngRow, posEnd.lngCol)).Interior.Color = RGB(conRGBRed, conRGBGreen, conRGBBlue)

End Sub

'*******************************************************************************
'        ���׍s���� �Ԋ|�F�ݒ菈��
'*******************************************************************************
'        �����T�v�F�w�b�_�[�s�̎w���̖��׍s(���׍s���� �Ԋ|�F�ݒ��)�ɑ΂���
'                  ��r�������݂��Ȃ��s�̕����� �Ԋ|�F�ɐݒ肷��
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�I���|�W�V����   typCellPos
'            �����Q�@�F���ڃ|�W�V����   typCellPos
'*******************************************************************************
Private Sub ProcFontSet(posEnd As typCellPos, _
                        posItm() As typCellPos)

    Dim i    As Integer
    
    i = 0                                                  ' (���׍s���� �Ԋ|�F�ݒ��̑S�f�[�^���������s��
    Do While (i <= conArrayMax And _
              posItm(i).lngCol <> 0)
              
       posItm(i).lngRow = posItm(i).lngRow + 1             ' ��r�������݂��閾�׍s�͏����Ȃ�
                                                           ' �����F�ݒ��̖��׍s���I���|�W�V�����̍s �̊ԁA�������s��
       Do While (posItm(i).lngRow <= posEnd.lngRow)
                                                           ' �����F�� �Ԋ|�F�ɐݒ肷��
          Cells(posItm(i).lngRow, posItm(i).lngCol).Font.Color = RGB(conRGBRed, conRGBGreen, conRGBBlue)
                                                           ' ���t��̂Ƃ��A�j����������F�� �Ԋ|�F�ɐݒ肷��
          If (Cells(2, posItm(i).lngCol).Value = "�� �� �� �t" Or _
              Cells(2, posItm(i).lngCol).Value = "�� �� �� �t") Then
              Cells(posItm(i).lngRow, posItm(i).lngCol).Offset(, 1).Font.Color = RGB(conRGBRed, conRGBGreen, conRGBBlue)
          End If
           
          posItm(i).lngRow = posItm(i).lngRow + 1
       Loop
          
       i = i + 1
    Loop

End Sub
