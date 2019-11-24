Option Explicit

Dim isERR As Integer                                       ' ERROR �t���O

Public Type typCellPos                                     ' �Z���|�W�V�����^
    lngRow As Long
    lngCol As Long
End Type

Const conStartStr   As String = "��"                       ' �ꗗ�\�̊J�n������
Dim conItm          As Variant                             ' �w�b�_�[�s�̗�z��
                                                           ' �w�b�_�[�s�̗񖼏�
Const conHeadItm    As String = "��,��,��@���@��,���@���@��,�T�@�@�@�@ �v,�T�v�i�����p�j,�o �� �� �t,"

Const conArrayMax   As Integer = 10                        ' �w�b�_�[�s�̃R�s�[�⊮�Ώۗ� MAX��
Const conDiffItm    As Integer = 2                         ' �w�b�_�[�s�Ɩ��׍s�̍�(�s��)

'*******************************************************************************
'        �� �� �� �� �� ��
'*******************************************************************************
'        �����T�v�F���d���������u�a�`�܂Ƃ߁�.xlsm �ɂ����āA�R�s�[�⊮�Ώۗ�̏�r�������݂���
'                  ���׍s���R�s�[���A���r�������݂���s�܂Ńy�[�X�g���ĕ⊮����
'
'                  �R�s�[�⊮�Ώۗ�F"��", "��", "�啪��", "������", "�T�v", "�o�͓��t"
'
'                  �Ȃ��A�� �̓V���A���Ɏ������Ԃ��s��
'*******************************************************************************
Public Sub ���ڕ⊮����()

    Dim posStart            As typCellPos                  ' �����Ώۂ�START�Z���ʒu (�s��START�s�C���"is"��Œ�)
    Dim posEnd              As typCellPos                  ' �����Ώۂ�END�Z���ʒu   (�s��END�s�C���"is"��Œ�)
    Dim posItm(conArrayMax) As typCellPos                  ' �R�s�[�⊮�Ώۗ�̃Z���ʒu �z��
                                                           ' (�s�͖��׍s�Œ�C��̓R�s�[�⊮�Ώۗ�)
    Dim intNo               As Integer                     ' "��"��̃V���A��NO
    
    Dim i                   As Integer
    
    Application.ScreenUpdating = False                     ' ��ʍX�V ��~
    Call ProcInit(posStart, posEnd, posItm)                ' ��������
    
    intNo = 0
    Do While (posItm(0).lngRow <= posEnd.lngRow)           ' �R�s�[�⊮�Ώۗ�̍s �� �����Ώ�END�s �̊� �������s��
    
       Call ProcItmPrep(posItm)                            ' �T�v(�����p) �P�s�ڂ��쐬
       intNo = intNo + 1                                   ' ���V���A���l �C���N�������g
       Call ProcItmCopy(intNo, posStart, posItm)           ' ���ڕ⊮����
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
                     
    Dim rngFind   As Range                                 ' FIND�֐��̃��^�[���l
    Dim intRtn    As Integer                               ' MSGBOX�֐��̃��^�[���l
    
    Dim i         As Integer
    Dim j         As Integer
    
    conItm = Split(conHeadItm, ",")                        ' �w�b�_�[�s�̗񖼏̂�z��ɂ���
    
    Set rngFind = Cells.Find(conStartStr)                  ' �J�n������ ����
    If rngFind Is Nothing Then                             ' �J�n������ ���Ȃ��Ƃ�
       intRtn = MsgBox(prompt:="�J�n������ ��������܂���", Buttons:=vbOKOnly + vbCritical)
       If intRtn = vbOK Then
          isERR = True
          Exit Sub
       End If
    End If
                                                           ' �����Ώۂ�START�Z��
    posStart.lngRow = rngFind.Row                          ' �sNO �ɊJ�n������̍s �ݒ�
    posStart.lngCol = rngFind.Column                       ' ��NO �ɊJ�n������̗� �ݒ�
                                                           ' �����Ώۂ�End�Z���ʒu
    posEnd.lngRow = Cells.SpecialCells(xlCellTypeLastCell).Row    ' �sNO �Ɏg�p�͈͓��ŏI�Z���̍s �ݒ�
    posEnd.lngCol = Cells.SpecialCells(xlCellTypeLastCell).Column ' �sNO �Ɏg�p�͈͓��ŏI�Z���̍s �ݒ�
    
    i = 0
    j = 0
    Do While conItm(i) <> ""                               ' �w�b�_�[�s�̗񖼏̔z��̑S�f�[�^���������s��
    
       Set rngFind = Cells.Find(conItm(i))                 ' �w�b�_�[�s�̗񖼏̔z���FIND����
       If rngFind Is Nothing Then                          ' �w�b�_�[�s�̗񖼏̂��Ȃ��Ƃ�
       Else                                                ' �w�b�_�[�s�̗񖼏̂�����Ƃ�
          posItm(j).lngRow = rngFind.Offset(conDiffItm).Row       ' �sNO �Ƀw�b�_�[�s�̗񖼏̖̂��׍s �ݒ�
          posItm(j).lngCol = rngFind.Column                ' ��NO �Ƀw�b�_�[�s�̗񖼏̗̂� �ݒ�
          
          j = j + 1
       End If
       
       i = i + 1
    Loop
    
End Sub

'*******************************************************************************
'        �� �� �� �� �� �� �� ��
'*******************************************************************************
'        �����T�v�F"�T�v"��ɕ����s�ݒ肳��Ă���Ƃ��A���������ׂČ�����
'                  �T�v(�����p)�̂P�s�ڂɐݒ肷��
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F���ڃ|�W�V����   typCellPos
'*******************************************************************************
Private Sub ProcItmPrep(posItm() As typCellPos)

    Dim intRow  As Integer
    Dim rngCell As Range
    Dim strComb As String
    Dim lngRGB  As Long

    strComb = ""
    intRow = 1                                             ' "�T�v"��̏�r���ݒ�s�̖Ԋ|�F��ۑ�
    lngRGB = Cells(posItm(4).lngRow, posItm(4).lngCol).Interior.Color
                                                           ' "�T�v"��̎��̉��r�����o�Ă��鍷���s�����Z�o
    Do While (Cells(posItm(4).lngRow, posItm(4).lngCol).Offset(intRow).Borders(xlEdgeTop).LineStyle = xlLineStyleNone)
                                                           ' "�T�v"��̃R�s�[�⊮�s��������
       If Cells(posItm(4).lngRow, posItm(4).lngCol).Offset(intRow).Font.Color = lngRGB Then
          Cells(posItm(4).lngRow, posItm(4).lngCol).Offset(intRow).Value = ""
       End If
       
       intRow = intRow + 1
    Loop
                                                           ' "�T�v"���r���`���r���ݒ�s�̂��ׂĂ̕����������
    For Each rngCell In Range(Cells(posItm(4).lngRow, posItm(4).lngCol), _
                              Cells(posItm(4).lngRow, posItm(4).lngCol).Offset(intRow - 1))
       
       strComb = strComb & rngCell.Value
    Next
                                                           '"�T�v(�����p)"�� �̏�r���ݒ�s�� "�T�v"��̌���������ݒ�
    Cells(posItm(5).lngRow, posItm(5).lngCol).Value = strComb  

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
                        posStart As typCellPos, _
                        posItm() As typCellPos)

    Dim i       As Integer
    Dim intRow  As Integer
    Dim lngRGB  As Long
    Dim rngCell As Range

    If Cells(posItm(2).lngRow, posItm(2).lngCol) <> "" Then    ' �啪�ޗ�̖��׍s�̒l���u�����N�̂Ƃ�
       Cells(posItm(0).lngRow, posItm(0).lngCol) = intNo       ' "��"��̖��׍s�̒l�ɃV���A��NO �ݒ�
    End If
    
    i = 0                                                  ' �R�s�[�⊮�Ώۗ�̑S�f�[�^���������s��
    Do While (i <= conArrayMax And _
              posItm(i).lngCol <> 0)
              
       Cells(posItm(i).lngRow, posItm(i).lngCol).Copy      ' �R�s�[�⊮�Ώۗ�̏�r�������݂��閾�׍s���R�s�[
       lngRGB = Cells(posItm(i).lngRow, posItm(i).lngCol).Interior.Color
       posItm(i).lngRow = posItm(i).lngRow + 1
                                                           ' �R�s�[�⊮�Ώۗ�̎��̏�r�����o�Ă���܂ŁA�������s��
       Do While (Cells(posItm(i).lngRow, posItm(i).lngCol).Borders(xlEdgeTop).LineStyle = xlLineStyleNone)

          If Cells(posItm(i).lngRow, posItm(i).lngCol).Value = "" Then
             Cells(posItm(i).lngRow, posItm(i).lngCol).PasteSpecial (xlPasteAllExceptBorders) ' �y�[�X�g(�r���������S��)
             Cells(posItm(i).lngRow, posItm(i).lngCol).Font.Color = lngRGB                    ' �����F�� WHITE �ɂ���
          Else
             Cells(posItm(i).lngRow, posItm(i).lngCol).Copy
          End If
                                                           ' ���t��̂Ƃ��A�j����������F�� WHITE �ɂ���
          If Cells(posStart.lngCol, posItm(i).lngCol).Value = "�o �� �� �t" Then
             Cells(posItm(i).lngRow, posItm(i).lngCol).Offset(, 1).Font.Color = lngRGB
          End If
           
          posItm(i).lngRow = posItm(i).lngRow + 1
       Loop
          
       i = i + 1
    Loop

End Sub
