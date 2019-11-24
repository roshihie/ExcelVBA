Option Explicit
'*******************************************************************************
'        �\�[�X��̓V�[�g  �V���R���y�A���V�V�[�g�֔��f
'*******************************************************************************
'      < �����T�v >
'        �\�[�X��͎��R����COBOL�\�[�X�̐V���V�[�g���R���y�A���A�V�V�[�g���
'        ���V�[�g�̃R�����g�𔽉f����B
'*******************************************************************************
Private Type tSocStrct                 '�\�[�X�\����
  bMatch     As Boolean                  '�}�b�`���O����
  nOpidx     As Long                     '���葤index
  sSeqno     As String                   '�\�[�XSEQNO
  sSourc     As String                   '�\�[�X�R�[�h
  sComnt(46) As String                   '�R�����g(C��`AV��)
End Type

Private Type tSocPropt                 '�\�[�X����
  oSheet     As Worksheet                '�V�[�g
  nSocMinRow As Long                     '�\�[�X�ŏ��s
  nSocMaxRow As Long                     '�\�[�X�ő�s
  nSocMaxidx As Long                     '�\�[�X�ő�index
End Type

Private Type tSocCompr                 '�R���y�Aindex
  nLimBegin  As Long                     '����idx Begin
  nLimEnd    As Long                     '����idx End
  nCmpBegin  As Long                     '�R���y�Aidx Begin
  nCmpEnd    As Long                     '�R���y�Aidx End
End Type

Private Type tSocBody                  '�\�[�X�G���A
  oSoc()     As tSocStrct                '�\�[�X�V�[�g
  oPrp       As tSocPropt                '�\�[�X����
  oidx       As tSocCompr                '�R���y�Aindex
End Type

Const g_cnSocCol       As Long = 2     '�\�[�X��
Const g_cnLeftComntCol As Long = 3     '�R�����g��(���[)
Const g_cnRigtComntCol As Long = 48    '�R�����g��(�E�[)
Const g_cnSocMinRow    As Long = 3     '�\�[�X�ŏ��s

Dim g_oFso   As New FileSystemObject   '##### Debug
Dim g_oTStrm As Object                 '##### Debug

'*******************************************************************************
'        ���C������
'*******************************************************************************
Public Sub COBOL��̓V�[�g�R���y�A()
  
  Const cnOutFile As String = "C:\Users\roshi_000\MyImp\MyOwn\Develop\Excel VBA\My VBA\Debug.txt"  '##### Debug

  Const cnStartBlk As Long = 45        '�}�b�`���O�u���b�N�����l
  Const cnStep     As Long = -1        '���Z�s
  
  Dim oNew    As tSocBody              '�V�\�[�X
  Dim oOld    As tSocBody              '���\�[�X
  Dim nBlock  As Long                  '�R���y�A�u���b�N
  
  Call sb��������(oNew, oOld)
  
  Set g_oTStrm = g_oFso.CreateTextFile(cnOutFile, True)  '##### Debug
  
  For nBlock = cnStartBlk To 1 Step cnStep
    Call sb�}�b�`���O����(nBlock, oNew, oOld)
  Next
  
  Call sb������Debug(oNew, oOld)                         '##### Debug
 'Call sb�V�V�[�g�ҏW(oNew, oOld)
 'Call sb�I������

  g_oTStrm.Close                                         '##### Debug

End Sub

'*******************************************************************************
'        ��������
'*******************************************************************************
Private Sub sb��������(oNew As tSocBody, oOld As tSocBody)
  
  Dim sNewSheet  As String
  Dim sOldSheet  As String
  
  sNewSheet = Application.InputBox(Prompt:="�V�\�[�g������͂��ĉ������B", _
                                   Title:="�V�\�[�X���w��", Type:=2)
  sOldSheet = Application.InputBox(Prompt:="���\�[�g������͂��ĉ������B", _
                                   Title:="���\�[�X���w��", Type:=2)
                                   
  Call sb�\�[�X�G���A�ݒ�(oNew, sNewSheet)
  Call sb�\�[�X�G���A�ݒ�(oOld, sOldSheet)
  
  oNew.oSoc(oNew.oPrp.nSocMaxidx).nOpidx = oOld.oPrp.nSocMaxidx
  oOld.oSoc(oOld.oPrp.nSocMaxidx).nOpidx = oNew.oPrp.nSocMaxidx
  
End Sub

'*******************************************************************************
'        �\�[�X�G���A�ݒ�
'*******************************************************************************
'      < �����T�v >
'        �\�[�X�����C�\�[�X�\���̂�ݒ肷��B
'*******************************************************************************
Private Sub sb�\�[�X�G���A�ݒ�(oCmn As tSocBody, sCmnSheet As String)
  
  Const cnLenSeq As Long = 6
  Const cnBgnSoc As Long = 7
  Const cnLenSoc As Long = 66
  
  Dim nRow       As Long
  Dim nidx       As Long
  Dim nComntCol  As Long
  Dim nComntidx  As Long

  With oCmn.oPrp
    Set .oSheet = Worksheets(sCmnSheet)
    .nSocMinRow = g_cnSocMinRow
    .nSocMaxRow = .oSheet.Cells(.oSheet.Rows.Count, g_cnSocCol).End(xlUp).Row
    .nSocMaxidx = .nSocMaxRow - g_cnSocMinRow
    ReDim oCmn.oSoc(.nSocMaxidx)
  End With

  For nRow = oCmn.oPrp.nSocMinRow To oCmn.oPrp.nSocMaxRow
    nidx = nRow - g_cnSocMinRow
    With oCmn.oSoc(nidx)
      .bMatch = False
      .nOpidx = 0
      .sSeqno = Left(oCmn.oPrp.oSheet.Cells(nRow, g_cnSocCol).Value, cnLenSeq)
      .sSourc = Trim(Mid(oCmn.oPrp.oSheet.Cells(nRow, g_cnSocCol).Value, cnBgnSoc, cnLenSoc))
      
      For nComntCol = g_cnLeftComntCol To g_cnRigtComntCol
        nComntidx = nComntCol - g_cnLeftComntCol
        .sComnt(nComntidx) = oCmn.oPrp.oSheet.Cells(nRow, nComntCol).Value
      Next
    End With
  Next
  
End Sub

'*******************************************************************************
'        �}�b�`���O����
'*******************************************************************************
Private Sub sb�}�b�`���O����(nBlock As Long, _
                             oNew As tSocBody, oOld As tSocBody)
  Dim nNewRept  As Long
  Dim nOldRept  As Long
  Dim nMchCond  As Long

  oOld.oidx.nLimBegin = 0                        'Old����idx �ݒ�
  oOld.oidx.nLimEnd = oOld.oPrp.nSocMaxidx

  oOld.oidx.nCmpBegin = oOld.oidx.nLimBegin      'Old�R���y�Aidx �ݒ�
  nOldRept = fn�R���y�Aidx�m��(nBlock, oOld)

  Do While (nOldRept = 0)                        'Old�R���y�Aidx Begin �擾�\�̊� �s��

    Call sb�V�\�[�X����idx�m��(oOld, oNew)         'New����idx �ݒ�
    oNew.oidx.nCmpBegin = oNew.oidx.nLimBegin      'New�R���y�Aidx �ݒ�
    nNewRept = fn�R���y�Aidx�m��(nBlock, oNew)
      
    nMchCond = 9                                   '�R���y�A���� �����{

    Do While (nNewRept = 0)                        'New�R���y�Aidx Begin, End �擾�ς̊� �s��
 
      nMchCond = 1                                   '�R���y�A���� �A���}�b�`
      nMchCond = fn�R���y�A���{(oNew, oOld)          '�R���y�A���{
        
      If nMchCond = 0 Then                           '�R���y�A���� �}�b�`
        Call sb�}�b�`������(oNew, oOld)
        Call sb�}�b�`Debug(nBlock, oNew, oOld)         '##### Debug #####
        Exit Do
      End If

      oNew.oidx.nCmpBegin = oNew.oidx.nCmpBegin + 1  'New�R���y�Aidx �ݒ�
      nNewRept = fn�R���y�Aidx�m��(nBlock, oNew)
    Loop

    If (nMchCond = 0) Then                       '�R���y�A���� �}�b�`
      oOld.oidx.nCmpBegin = oOld.oidx.nCmpBegin + nBlock
      
    ElseIf (nMchCond = 1) Then                   '�R���y�A���� �A���}�b�`
      oOld.oidx.nCmpBegin = oOld.oidx.nCmpBegin + 1
       
    ElseIf (nMchCond = 9) Then                   '�R���y�A���� �����{
      oOld.oidx.nCmpBegin = oOld.oidx.nCmpBegin + nBlock
        
    End If

    nOldRept = fn�R���y�Aidx�m��(nBlock, oOld)   'Old�R���y�Aidx �ݒ�
  Loop
    
End Sub

'*******************************************************************************
'        �R���y�Aidx �m��
'*******************************************************************************
'      < �����T�v >
'        �����Ƀu���b�N�s�C�\�[�X�{�̂��󂯎��A�R���y�Aidx Begin�CEnd ��
'        �擾����B
'        �R���y�Aidx Begin ���擾�s�\�̏ꍇ      �߂�l�� 99
'        �R���y�Aidx End   ������idx�𒴂����ꍇ �߂�l�� 90
'        �R���y�A�͈͂Ƀ}�b�`�m��s������ꍇ    �߂�l�� 10 (��ʂɕԂ��Ȃ�)
'        �R���y�Aidx Begin, End ���擾���ꂽ�ꍇ �߂�l�� 00 ��Ԃ��B
'*******************************************************************************
Private Function fn�R���y�Aidx�m��(nBlock As Long, oCmn As tSocBody) As Long

  Dim nCmnRept  As Long
  Dim nNextEnd  As Long

  nNextEnd = oCmn.oidx.nCmpBegin
  nCmnRept = 10
  
  Do While (nCmnRept = 10)
                                       '�R���y�Aidx Begin �擾
    oCmn.oidx.nCmpBegin = nNextEnd
    If (fn�R���y�Aidx_Begin�m��(oCmn) = False) Then
      fn�R���y�Aidx�m�� = 99
      Exit Function
    End If
                                       '�R���y�Aidx End   �擾
    oCmn.oidx.nCmpEnd = oCmn.oidx.nCmpBegin + nBlock - 1
    nCmnRept = fn�R���y�A�͈͍s�`�F�b�N(nBlock, oCmn)
    If nCmnRept = 90 Then
      fn�R���y�Aidx�m�� = nCmnRept
      Exit Function
    End If
    nNextEnd = oCmn.oidx.nCmpEnd + 1
  Loop
  
  fn�R���y�Aidx�m�� = 0

End Function

'*******************************************************************************
'        �R���y�Aidx Begin �m��
'*******************************************************************************
'      < �����T�v >
'        �����Ƀ\�[�X�{�̂��󂯎��A�\�[�X�{�̂��������āA
'        �}�b�`���m��̍ŏ��s���擾�\�̏ꍇ�A
'            ���ʂ��R���y�Aidx Begin�ɐݒ肵�A�߂�l�� True ��Ԃ��B
'        �}�b�`���m��s���擾�s�̏ꍇ�A
'            �߂�l�� False ��Ԃ��B
'*******************************************************************************
Private Function fn�R���y�Aidx_Begin�m��(oCmn As tSocBody) As Boolean

  Dim nidx  As Long
  Dim nEnd  As Long

  fn�R���y�Aidx_Begin�m�� = False
  nidx = oCmn.oidx.nCmpBegin
  
  Do While (nidx <= oCmn.oidx.nLimEnd)
    If oCmn.oSoc(nidx).bMatch = False Then
      oCmn.oidx.nCmpBegin = nidx
      fn�R���y�Aidx_Begin�m�� = True
      Exit Function
    End If
    nidx = nidx + 1
  Loop

End Function

'*******************************************************************************
'        �R���y�A�͈͍s�`�F�b�N
'*******************************************************************************
'      < �����T�v >
'        �����Ƀu���b�N�s�C�\�[�X�{�̂��󂯎��A�R���y�Aidx Begin�{�P�s����
'        �R���y�Aidx Begin�{�P�s����R���y�Aidx Begin�{�u���b�N�s�܂ŁA���ׂ�
'        �}�b�`���m��s�ł���`�F�b�N���s���B
'        ����idx End���R���y�Aidx End �̏ꍇ�A�߂�l�� 90
'        �R���y�A�͈͂Ƀ}�b�`�m��s������ꍇ �߂�l�� 10
'        ����ȊO�̏ꍇ                       �߂�l�� 00 ��Ԃ��B
'*******************************************************************************
Private Function fn�R���y�A�͈͍s�`�F�b�N(nBlock As Long, oCmn As tSocBody) As Long

  Dim nidx  As Long
  
  If oCmn.oidx.nLimEnd < oCmn.oidx.nCmpEnd Then
    fn�R���y�A�͈͍s�`�F�b�N = 90
    Exit Function
  End If
  
  nidx = oCmn.oidx.nCmpBegin + 1
  Do While (nidx <= oCmn.oidx.nCmpEnd)
    If oCmn.oSoc(nidx).bMatch = True Then
      fn�R���y�A�͈͍s�`�F�b�N = 10
      Exit Function
    End If
    nidx = nidx + 1
  Loop

  fn�R���y�A�͈͍s�`�F�b�N = 0

End Function

'*******************************************************************************
'        �V�\�[�X���� index �m��
'*******************************************************************************
'      < �����T�v >
'        �����ɋ��\�[�X�C�V�\�[�X���󂯎��A
'      �E���\�[�X�R���y�Aidx End ���牺���ɋ��\�[�X���������āA�}�b�`�m��ύs��
'        �ŏ��s���擾���A����ɊY�����鑊��(�V�\�[�X)�� index ����
'        �V�\�[�X�̃}�b�`���m��s��V�\�[�X����idx End �Ƃ���B
'      �E���\�[�X�R���y�Aidx Begin ������ ���\�[�X���������āA�}�b�`�m��ύs��
'        �ŏ��s���擾���A����ɊY�����鑊��(�V�\�[�X)�� index ����
'        �V�\�[�X�̃}�b�`���m��s��V�\�[�X����idx Begin �Ƃ���B
'*******************************************************************************
Private Sub sb�V�\�[�X����idx�m��(oOld As tSocBody, oNew As tSocBody)

  Dim nidx  As Long
  Dim nEnd  As Long
                                       '�V�\�[�X����idx End   �擾
  nEnd = oOld.oPrp.nSocMaxidx
  nidx = oOld.oidx.nCmpEnd
  oNew.oidx.nLimEnd = oNew.oPrp.nSocMaxidx
  
  Do While (nidx <= nEnd)
    If oOld.oSoc(nidx).bMatch = True Then
      oNew.oidx.nLimEnd = oOld.oSoc(nidx).nOpidx - 1
      Exit Do
    End If
    nidx = nidx + 1
  Loop
                                       '�V�\�[�X����idx Begin �擾
  nEnd = 0
  nidx = oOld.oidx.nCmpBegin
  oNew.oidx.nLimBegin = 0
  
  Do While (nidx >= nEnd)
    If oOld.oSoc(nidx).bMatch = True Then
      oNew.oidx.nLimBegin = oOld.oSoc(nidx).nOpidx + 1
      Exit Do
    End If
    nidx = nidx - 1
  Loop

End Sub

'*******************************************************************************
'        �R���y�A���{
'*******************************************************************************
'      < �����T�v >
'        �V���\�[�X�̃R���y�A�͈�index Begin�`End �̃R���y�A���s���B
'        �R���y�A�u���b�N�P�ʂŃ}�b�`���Ă���ꍇ 0 ��Ԃ��B
'        �A���}�b�`�̏ꍇ                         1 ��Ԃ��B
'*******************************************************************************
Private Function fn�R���y�A���{(oNew As tSocBody, oOld As tSocBody) As Long

  Dim nNewidx  As Long
  Dim nOldidx  As Long

  nNewidx = oNew.oidx.nCmpBegin
  nOldidx = oOld.oidx.nCmpBegin
  
  Do While (nNewidx <= oNew.oidx.nCmpEnd)
    If oNew.oSoc(nNewidx).sSourc <> oOld.oSoc(nOldidx).sSourc Then
      fn�R���y�A���{ = 1
      Exit Function
    End If
    nNewidx = nNewidx + 1
    nOldidx = nOldidx + 1
  Loop

  fn�R���y�A���{ = 0

End Function

'*******************************************************************************
'        �}�b�`������
'*******************************************************************************
'      < �����T�v >
'        �R���y�A�u���b�N�P�ʂŃ}�b�`�����ꍇ�A�R���y�A���ʂ� True, ����idx ��
'        ���\�[�X�C�V�\�[�X���ꂼ�ꑊ��̃}�b�`���� index ��ݒ肷��B
'*******************************************************************************
Private Sub sb�}�b�`������(oNew As tSocBody, oOld As tSocBody)

  Dim nNewidx  As Long
  Dim nOldidx  As Long

  nNewidx = oNew.oidx.nCmpBegin
  nOldidx = oOld.oidx.nCmpBegin

  Do While (nNewidx <= oNew.oidx.nCmpEnd)
    oNew.oSoc(nNewidx).bMatch = True
    oNew.oSoc(nNewidx).nOpidx = nOldidx
    oOld.oSoc(nOldidx).bMatch = True
    oOld.oSoc(nOldidx).nOpidx = nNewidx
    
    nNewidx = nNewidx + 1
    nOldidx = nOldidx + 1
  Loop

End Sub

'*******************************************************************************
'        �}�b�`Debug ����
'*******************************************************************************
Private Sub sb�}�b�`Debug(nBlock As Long, oNew As tSocBody, oOld As tSocBody)

  g_oTStrm.WriteLine "########### �}�b�`��� ###########"
  g_oTStrm.WriteLine "  ���u���b�N�s��        : " & nBlock
  g_oTStrm.WriteLine "    �V�\�[�X   CmpBegin : " & oNew.oidx.nCmpBegin
  g_oTStrm.WriteLine "               CmpEnd   : " & oNew.oidx.nCmpEnd
  g_oTStrm.WriteLine " "
  g_oTStrm.WriteLine "    ���\�[�X   CmpBegin : " & oOld.oidx.nCmpBegin
  g_oTStrm.WriteLine "               CmpEnd   : " & oOld.oidx.nCmpEnd
  
End Sub
  
'*******************************************************************************
'        ������Debug ����
'*******************************************************************************
Private Sub sb������Debug(oNew As tSocBody, oOld As tSocBody)

  Dim nidx  As Long
  
  g_oTStrm.WriteLine " "
  g_oTStrm.WriteLine " "
  g_oTStrm.WriteLine "####### New Source ############################################################"
  For nidx = 0 To oNew.oPrp.nSocMaxidx
    With oNew.oSoc(nidx)
      g_oTStrm.WriteLine .sSeqno & " " & .sSourc & " " & Space(75 - Len(.sSourc)) & .bMatch
    End With
  Next
  
  g_oTStrm.WriteLine " "
  g_oTStrm.WriteLine " "
  g_oTStrm.WriteLine "####### Old Source ############################################################"
  For nidx = 0 To oOld.oPrp.nSocMaxidx
    With oOld.oSoc(nidx)
      g_oTStrm.WriteLine .sSeqno & " " & .sSourc & " " & Space(75 - Len(.sSourc)) & .bMatch
    End With
  Next

End Sub

'*******************************************************************************
'        �V�V�[�g�ҏW����
'*******************************************************************************
Private Sub sb�V�V�[�g�ҏW(oNew As tSocBody, oOld As tSocBody)

  Dim nNewidx  As Long
  Dim nOldidx  As Long
  Dim nidx     As Long
  
  oOld.oidx.nLimBegin = 0                        'Old����idx �ݒ�
  oOld.oidx.nLimEnd = oOld.oPrp.nSocMaxidx
  oNew.oidx.nLimBegin = 0                        'New����idx �ݒ�
  oNew.oidx.nLimEnd = oNew.oPrp.nSocMaxidx
  
  nNewidx = 0
  nOldidx = 0
  Do While (nNewidx <= oNew.oidx.nLimEnd)

    oNew.oidx.nCmpBegin = fn�}�b�`�s�擾(oNew, nNewidx)
    nidx = oNew.oidx.nCmpBegin
    oOld.oidx.nCmpBegin = oOld.oSoc(nidx).nOpidx
    
    For nidx = nOldidx To oOld.oidx.nCmpBegin - 1
      '�A���}�b�` Old�\�[�X �o��
    Next
    
    For nidx = nNewidx To oNew.oidx.nCmpBegin - 1
      '�A���}�b�` New�\�[�X �o��
    Next
    
    nNewidx = oNew.oidx.nCmpBegin
    Do While (nidx <= oNew.oidx.nLimEnd And _
              oNew.oSoc(nidx).bMatch = False)
        Exit Do
      End
      '�}�b�` Old �\�[�X �o��
      
      nNewidx = nNewidx + 1
    Loop
    
  Loop

End Sub
