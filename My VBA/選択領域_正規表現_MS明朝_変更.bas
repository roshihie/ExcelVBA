Option Explicit
'*******************************************************************************
'        ���K�\�������ɂ�� �I��̈�t�H���g �l�r���� �ύX
'*******************************************************************************
'      < �����T�v >
'        ���K�\���ɂ�芿�����������A�q�b�g���������̃t�H���g�� ��l�r �����
'        �ɕύX����
'*******************************************************************************
Public Sub �I��̈�_���K�\��_MS����_�ύX()

  Const cnKanji As String = "[��-�]+|[��-��]+|[�@-��]+|[��-���`-�y�O-�X]+|[�u�v��]+"
  Dim oRegExp  As RegExp
  Dim oMatchCl As MatchCollection
  Dim oMatch   As Match
  Dim nSize    As Long
  Dim oCell    As Range

  Application.ScreenUpdating = False
  
  If TypeName(ActiveCell) = "Range" Then
    nSize = Application.InputBox(Prompt:="�t�H���g�T�C�Y����͂���", _
                                 Title:="�l�r ���� �t�H���g�T�C�Y�w��", _
                                 Default:=11, _
                                 Type:=1)
    If nSize = 0 Then
      Exit Sub
    End If

    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = cnKanji
    oRegExp.Global = True
    
    For Each oCell In Selection
      Set oMatchCl = oRegExp.Execute(oCell.Value)
      For Each oMatch In oMatchCl
        With oCell.Characters(Start:=oMatch.FirstIndex + 1, _
                              Length:=oMatch.Length).Font
          .Name = "�l�r ����"
          .Size = nSize
        End With
      Next
    Next
  Else
    MsgBox "�Z���̈悪�I������Ă��܂���", vbCritical
  End If
  
End Sub
