Attribute VB_Name = "Module1"
Option Explicit
'*******************************************************************************
'        ������擾����
'*******************************************************************************
'      < �����T�v >
'        �w�肳�ꂽ�G���A�̕����񂩂�u�����N�ŋ�؂�ꂽ���ꂼ��̕������
'        �������A�w�肳�ꂽ�z��ԍ��̕������Ԃ��B
'*******************************************************************************
Public Function fn������擾(oRang As Range, sDlm As String, nPos As Integer) As String

  Dim asStr() As String
 
  asStr = Split(oRang, sDlm)
  fn������擾 = asStr(nPos)

End Function

