Option Explicit
'*******************************************************************************
'        �I��̈�t�H���g �l�r���� �ύX
'*******************************************************************************
'      < �����T�v >
'        �I��̈�̕����t�H���g���u�l�r �����v�ɕύX����
'*******************************************************************************
Public Sub �I��̈�_MS����_�ύX()

  Dim nSize  As Long
  
  If TypeName(Selection) = "Range" Then
    nSize = Application.InputBox(Prompt:="�t�H���g�T�C�Y����͂���", _
                                 Title:="�l�r ���� �t�H���g�T�C�Y�w��", _
                                 Default:=11, _
                                 Type:=1)
    If nSize = 0 Then
      Exit Sub
    End If
    
    With Selection.Font
      .Name = "�l�r ����"
      .Size = nSize
      .Bold = False
    End With
  Else
    MsgBox "�Z���̈悪�I������Ă��܂���", vbCritical
  End If
  
End Sub
