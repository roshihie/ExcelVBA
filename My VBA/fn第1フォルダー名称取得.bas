Option Explicit
'*******************************************************************************
'        ��P�t�H���_�[���̎擾
'*******************************************************************************
'      < �����T�v >
'        �p�X�������P�t�H���_�[���̂��擾����
'*******************************************************************************
Function fn��1�t�H���_�[���̎擾(oRange As Range, sDlmt As String) As String

  Dim asFolder() As String
  Dim nToIdx     As Integer
  Dim i          As Integer
  
  asFolder = Split(oRange, sDlmt)
  
  If Left(oRange, 2) = "\\" Then
    nToIdx = LBound(asFolder) + 2
  Else
    nToIdx = LBound(asFolder)
  End If
  
  fn��1�t�H���_�[���̎擾 = asFolder(LBound(asFolder))
  For i = LBound(asFolder) + 1 To nToIdx
    fn��1�t�H���_�[���̎擾 = fn��1�t�H���_�[���̎擾 & "\" & asFolder(i)
  Next

End Function

