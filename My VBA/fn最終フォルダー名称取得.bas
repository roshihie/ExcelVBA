Option Explicit
'*******************************************************************************
'        �ŏI�t�H���_�[���̎擾
'*******************************************************************************
'      < �����T�v >
'        �p�X������ŏI�t�H���_�[���̂��擾����
'*******************************************************************************
Function fn�ŏI�t�H���_�[���̎擾(oRange As Range, sDlmt As String) As String

  Dim asFolder() As String
  Dim nToIdx     As Integer
  Dim i          As Integer
  
  asFolder = Split(oRange, sDlmt)
  fn�ŏI�t�H���_�[���̎擾 = asFolder(UBound(asFolder))

End Function

