Option Explicit

Dim isERR As Integer                                                ' ERROR �t���O

Public Type typCellPos                                              ' �Z���|�W�V�����^
    lngRow As Long
    lngCol As Long
End Type

Dim conItm          As Variant                                      ' �w�b�_�[�s�̗�z��
                                                                    ' �w�b�_�[�s�̗񖼏�
Const conHeadItm    As String = "��,��,��@���@��,���@���@��,�T�@�@�@ �v,�� �� �� �t,�� �� �� �t,"

Const conArrayMax   As Integer = 10                                 ' �w�b�_�[�s�̃R�s�[�⊮�Ώۗ� MAX��
Const conDiffItm    As Integer = 2                                  ' �w�b�_�[�s�Ɩ��׍s�̍�(�s��)
Const conRGBMax     As Integer = 255                                '�Ԋ|���F(RGB White�p)

'*******************************************************************************
'        �� �� �� �� �� ��
'*******************************************************************************
'        �����T�v�F����L�Q���w�K�|�C���g��.xlsm �ɂ����āA�R�s�[�⊮�Ώۗ�̏�r�������݂���
'                  ���׍s���R�s�[���A���r�������݂���s�܂Ńy�[�X�g���ĕ⊮����
'
'                  �R�s�[�⊮�Ώۗ�F"��", "��", "�啪��", "������", "�T�v", "�������t", "�Ώ����t"
'
'                  �Ȃ��A�� �̓V���A���Ɏ������Ԃ��s��
'*******************************************************************************
Public Sub ���ڕ⊮����()

    Dim posStart            As typCellPos                           ' �����Ώۂ�START�Z���ʒu (�s��START�s�C���"is"�s�Œ�)
    Dim posEnd              As typCellPos                           ' �����Ώۂ�END�Z���ʒu   (�s��END�s�C���"is"�s�Œ�)
    Dim posItm(conArrayMax) As typCellPos                           ' �R�s�[�⊮�Ώۗ�̃Z���ʒu �z��
                                                                    ' (�s�͖��׍s�Œ�C��̓R�s�[�⊮�Ώۗ�)
    Dim intNo               As Integer                              ' "��"��̃V���A��NO
    
    Dim i                   As Integer
    
    Application.ScreenUpdating = False                              ' ��ʍX�V ��~
    Call ProcInit(posStart, posEnd, posItm)                         ' ��������
    
    intNo = 0
    Do While (posItm(0).lngRow <= posEnd.lngRow)                    ' �R�s�[�⊮�Ώۗ�̍s �� �����Ώ�END�s �̊� �������s��
    
       intNo = intNo + 1                                            ' ���V���A���l �C���N�������g
       Call ProcItmCopy(intNo, posItm)                              ' ���ڕ⊮����
    Loop

End Sub

'*******************************************************************************
'        ���@���@���@��
'*******************************************************************************
'        �����T�v�F�����Ώۂ�START�s�CEND�s���m�肵�A�R�s�[�⊮�Ώۗ��z��Ɋi�[����
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�J�n�|�W�V����   typCellPos
'            �����Q�@�F�I���|�W�V����   typCellPos
'            �����R  �F���ڃ|�W�V����   typCellPos
'*******************************************************************************
Private Sub ProcInit(posStart As typCellPos, _
                     posEnd As typCellPos, _
                     posItm() As typCellPos)
                     
    Dim rngFind   As Range                                          ' FIND�֐��̃��^�[���l
    Dim lngCol    As Long                                           ' ���[�N��NO
    Dim intRtn    As Integer                                        ' MSGBOX�֐��̃��^�[���l
    
    Dim i, j      As Integer
    
    conItm = Split(conHeadItm, ",")                                 ' �w�b�_�[�s�̗񖼏̂�z��ɂ���
    
    Set rngFind = Cells.Find("is")                                  ' is��(�S�s "1" ���ߍ��ݗ�) ����
    If rngFind Is Nothing Then                                      ' is�� ���Ȃ��Ƃ�
       intRtn = MsgBox(prompt:="is��(�S�s ""1"" ���ߍ��ݗ�) ��������܂���", Buttons:=vbOKOnly + vbCritical)
       If intRtn = vbOK Then
          isERR = True
          Exit Sub
       End If
    Else                                                            ' is�� ������Ƃ�
       lngCol = rngFind.Column                                          ' is�� �̗�NO��ݒ�
    End If
                                                                    ' �����Ώۂ�END�Z���ʒu
    posEnd.lngRow = Cells(ActiveSheet.Rows.Count, lngCol).End(xlUp).Row ' �sNO �Ɉꗗ�\���ו��̍ŏI�s�� �ݒ�
    posEnd.lngCol = lngCol                                              ' ��NO �� is�� �ݒ�
                                                                    ' �����Ώۂ�START�Z���ʒu
    posStart.lngRow = Cells(posEnd.lngRow, lngCol).End(xlUp).Row      ' �sNO �Ɉꗗ�\���ו��̐擪�s�� �ݒ�
    posStart.lngCol = lngCol                                          ' ��NO �� is�� �ݒ�
    
    i = 0
    Do While conItm(i) <> ""                                        ' �w�b�_�[�s�̗񖼏̔z��̑S�f�[�^���������s��
    
       Set rngFind = Cells.Find(conItm(i))                          ' �w�b�_�[�s�̗񖼏̔z���FIND����
       If rngFind Is Nothing Then                                   ' �w�b�_�[�s�̗񖼏̂��Ȃ��Ƃ�
          posItm(i).lngRow = 0
          posItm(i).lngCol = 0
       Else                                                         ' �w�b�_�[�s�̗񖼏̂�����Ƃ�
          posItm(i).lngRow = rngFind.Offset(conDiffItm).Row           ' �sNO �Ƀw�b�_�[�s�̗񖼏̖̂��׍s �ݒ�
          posItm(i).lngCol = rngFind.Column                           ' ��NO �Ƀw�b�_�[�s�̗񖼏̗̂� �ݒ�
       End If
       i = i + 1
    Loop
    
End Sub

'*******************************************************************************
'        �� �� �� �� �� ��
'*******************************************************************************
'        �����T�v�F�R�s�[�⊮�Ώۗ񂻂ꂼ��ɑ΂��āA��r���̂���s�̒l���R�s�[����
'                  ���r���̂���s�܂Ńy�[�X�g����
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�V���A��NO       Integer
'            �����Q�@�F���ڃ|�W�V����   typCellPos
'*******************************************************************************
Private Sub ProcItmCopy(intNo As Integer, _
                        posItm() As typCellPos)

    Dim i    As Integer

    If Cells(posItm(2).lngRow, posItm(2).lngCol) <> "" Then         ' �啪�ޗ�̖��׍s�̒l���u�����N�̂Ƃ�
       Cells(posItm(0).lngRow, posItm(0).lngCol) = intNo              ' ����̖��׍s�̒l�ɃV���A��NO �ݒ�
    End If
    
    i = 0                                                           ' �R�s�[�⊮�Ώۗ�̑S�f�[�^���������s��
    Do While (i <= conArrayMax And _
              posItm(i).lngCol <> 0)
              
       Cells(posItm(i).lngRow, posItm(i).lngCol).Copy               ' �R�s�[�⊮�Ώۗ�̏�r�������݂��閾�׍s���R�s�[
       posItm(i).lngRow = posItm(i).lngRow + 1
                                                                    ' �R�s�[�⊮�Ώۗ�̎��̏�r�����o�Ă���܂ŁA�������s��
       Do While (Cells(posItm(i).lngRow, posItm(i).lngCol).Borders(xlEdgeTop).LineStyle = xlLineStyleNone)

          Cells(posItm(i).lngRow, posItm(i).lngCol).PasteSpecial (xlPasteAllExceptBorders)            ' �y�[�X�g(�r���������S��)
          Cells(posItm(i).lngRow, posItm(i).lngCol).Font.Color = RGB(conRGBMax, conRGBMax, conRGBMax) ' �����F�� WHITE �ɂ���
                                                                    ' ���t��̂Ƃ��A�j����������F�� WHITE �ɂ���
          If (Cells(2, posItm(i).lngCol).Value = "�� �� �� �t" Or _
              Cells(2, posItm(i).lngCol).Value = "�� �� �� �t") Then
              Cells(posItm(i).lngRow, posItm(i).lngCol).Offset(, 1).Font.Color = RGB(conRGBMax, conRGBMax, conRGBMax)
          End If
           
          posItm(i).lngRow = posItm(i).lngRow + 1
       Loop
          
       i = i + 1
    Loop

End Sub
