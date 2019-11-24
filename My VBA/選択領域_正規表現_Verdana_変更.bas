Option Explicit
'*******************************************************************************
'        ���K�\�������ɂ�� �I��̈�t�H���g Verdana �ύX
'*******************************************************************************
'      < �����T�v >
'        ���K�\���ɂ��p�����������A�q�b�g�����p���̃t�H���g�� �Verdana�
'        �ɕύX����
'*******************************************************************************
Public Sub �I��̈�_���K�\��_Verdana_�ύX()

  'Const cnEnglish As String = "[A-Z,a-z,?,!,;,:,',.,""]+"
  Const cnEnglish As String = "[A-Z]+|[a-z]+|[?]+|[!]+|[;]+|[:]+|[']+|[.]+|[""]+"
  Dim oRegExp  As RegExp
  Dim oMatchCl As MatchCollection
  Dim oMatch   As Match
  Dim nSize    As Long
  Dim oCell    As Range
  
  Application.ScreenUpdating = False
  
  If TypeName(Selection) = "Range" Then
    nSize = Application.InputBox(Prompt:="�t�H���g�T�C�Y����͂���", _
                                 Title:="Verdana �t�H���g�T�C�Y�w��", _
                                 Default:=11, _
                                 Type:=1)
    If nSize = 0 Then
      Exit Sub
    End If
    
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = cnEnglish
    oRegExp.Global = True
    
    For Each oCell In Selection
      Set oMatchCl = oRegExp.Execute(oCell.Value)
      For Each oMatch In oMatchCl
        With oCell.Characters(Start:=oMatch.FirstIndex + 1, _
                              Length:=oMatch.Length).Font
          .Name = "Verdana"
          .Size = nSize
        End With
      Next
    Next
  Else
    MsgBox "�Z���̈悪�I������Ă��܂���", vbCritical
  End If
  
End Sub
