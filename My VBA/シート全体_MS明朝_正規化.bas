Option Explicit
'*******************************************************************************
'        �V�[�g�S�̃t�H���g���K������
'*******************************************************************************
'      < �����T�v >
'        �V�[�g�S�̂̕����t�H���g���u�l�r �����v�ɕύX����
'*******************************************************************************
Public Sub �V�[�g�S��_MS����_���K��()

  Dim nSize  As Long
  
  nSize = Application.InputBox(Prompt:="�t�H���g�T�C�Y����͂���", _
                               Title:="�l�r ���� �t�H���g�T�C�Y�w��", _
                               Default:=10, _
                               Type:=1)
  If nSize = 0 Then
    Exit Sub
  End If
  
  With Cells.Font
    .Name = "�l�r ����"
    .Size = nSize
    .Bold = False
  End With

End Sub
