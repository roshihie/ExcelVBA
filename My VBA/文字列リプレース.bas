Option Explicit
'*******************************************************************************
'        �����񃊃v���[�X
'*******************************************************************************
'      < �����T�v >
'        �A�N�e�B�u�V�[�g�̃Z�����e�ɂ���
'        ������ϊ��ꗗ�i��p�u�b�N ��p�V�[�g�j�ɋL�ڂ��ꂽ�ύX�O������
'        �ύX�㕶���ɕύX����
'*******************************************************************************
Public Sub �����񃊃v���[�X()
   
    Dim oDic As Object
    Set oDic = CreateObject("Scripting.Dictionary")        'Dictionary�I�u�W�F�N�g�̐錾
    
    Dim i  As Integer
    With Workbooks("������ϊ���`.xlsx").Worksheets("�ϊ��ꗗ")
      For i = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
        oDic.Add .Cells(i, 1).Value, .Cells(i, 2).Value    'Dictionary�I�u�W�F�N�g�̏������A�v�f�̒ǉ�
      Next
    End With
    
    Dim bRslt        As Boolean
    Dim oSoc, oSoces As Range
    Dim sRep         As Variant
        
    Set oSoces = Range(Cells(1, 1), Cells(Rows.Count, 1).End(xlUp))
        
    For Each oSoc In oSoces
      For Each sRep In oDic                                'Dictionary�I�u�W�F�N�g���g�������������̒u��
        bRslt = oSoc.Replace(What:=sRep, Replacement:=oDic(sRep), LookAt:=xlPart)
      Next
    Next
End Sub
