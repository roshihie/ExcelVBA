Option Explicit
'*******************************************************************************
'        ���K�\�������ɂ�� �I��̈�t�H���g ���w�p���� �ύX
'*******************************************************************************
'      < �����T�v >
'        ���K�\���ɂ��p���A����ѐ������������A�q�b�g����
'        �p���̃t�H���g�� �cmmib10�, �����̃t�H���g�� �HGS����B� �ɕύX����
'        �������t�H���g�T�C�Y�� 11 �Œ�Ƃ���
'*******************************************************************************
Public Sub �I��̈�_���K�\��_���w�p����_�ύX()

  Const cnNumeric  As String = "[0-9]+"
  Const cnAlphabet As String = "[A-Z]+|[a-z]+"
  Const cnApostro  As String = "[']+"
  Const cnKakko    As String = "[()]+"
  Dim oRegExp  As RegExp
  Dim oMatchCl As MatchCollection
  Dim oMatch   As Match
  Dim nSize    As Long
  Dim oCell    As Range
  
  Application.ScreenUpdating = False
  
  If TypeName(Selection) = "Range" Then
    'nSize = Application.InputBox(Prompt:="�t�H���g�T�C�Y����͂���", _
    '                             Title:="Verdana �t�H���g�T�C�Y�w��", _
    '                             Default:=11, _
    '                             Type:=1)
    'If nSize = 0 Then
    '  Exit Sub
    'End If

    Set oRegExp = CreateObject("VBScript.RegExp")
'                                                          �p���t�H���g�ύX
    oRegExp.Pattern = cnAlphabet
    oRegExp.Global = True
    
    For Each oCell In Selection
      Set oMatchCl = oRegExp.Execute(oCell.Value)
      For Each oMatch In oMatchCl
        With oCell.Characters(Start:=oMatch.FirstIndex + 1, _
                              Length:=oMatch.Length).Font
          .Name = "BKM-cmmi10"
          .Size = 12
        End With
      Next
    Next
'                                                          �����t�H���g�ύX
    oRegExp.Pattern = cnNumeric
    oRegExp.Global = True
    
    For Each oCell In Selection
      Set oMatchCl = oRegExp.Execute(oCell.Value)
      For Each oMatch In oMatchCl
        With oCell.Characters(Start:=oMatch.FirstIndex + 1, _
                              Length:=oMatch.Length).Font
          .Name = "HGS����B"
          .Size = 11.5
        End With
      Next
    Next
'                                                          �A�|�X�g���t�B �t�H���g�ύX
    oRegExp.Pattern = cnApostro
    oRegExp.Global = True
    
    For Each oCell In Selection
      Set oMatchCl = oRegExp.Execute(oCell.Value)
      For Each oMatch In oMatchCl
        With oCell.Characters(Start:=oMatch.FirstIndex + 1, _
                              Length:=oMatch.Length).Font
          .Name = "CR�b���f�s��04"
          .Size = 12.5
        End With
      Next
    Next
'                                                          �J�b�R �t�H���g�ύX
    oRegExp.Pattern = cnKakko
    oRegExp.Global = True
    
    For Each oCell In Selection
      Set oMatchCl = oRegExp.Execute(oCell.Value)
      For Each oMatch In oMatchCl
        With oCell.Characters(Start:=oMatch.FirstIndex + 1, _
                              Length:=oMatch.Length).Font
          .Name = "Constantia"
          .Size = 12.5
        End With
      Next
    Next
  
  Else
    MsgBox "�Z���̈悪�I������Ă��܂���", vbCritical
  End If
  
End Sub
