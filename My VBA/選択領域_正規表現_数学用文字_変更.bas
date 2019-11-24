Option Explicit
'*******************************************************************************
'        正規表現検索による 選択領域フォント 数学用文字 変更
'*******************************************************************************
'      < 処理概要 >
'        正規表現により英字、および数字を検索し、ヒットした
'        英字のフォントを ｢cmmib10｣, 数字のフォントを ｢HGS明朝B｣ に変更する
'        ただしフォントサイズは 11 固定とする
'*******************************************************************************
Public Sub 選択領域_正規表現_数学用文字_変更()

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
    'nSize = Application.InputBox(Prompt:="フォントサイズを入力する", _
    '                             Title:="Verdana フォントサイズ指定", _
    '                             Default:=11, _
    '                             Type:=1)
    'If nSize = 0 Then
    '  Exit Sub
    'End If

    Set oRegExp = CreateObject("VBScript.RegExp")
'                                                          英字フォント変更
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
'                                                          数字フォント変更
    oRegExp.Pattern = cnNumeric
    oRegExp.Global = True
    
    For Each oCell In Selection
      Set oMatchCl = oRegExp.Execute(oCell.Value)
      For Each oMatch In oMatchCl
        With oCell.Characters(Start:=oMatch.FirstIndex + 1, _
                              Length:=oMatch.Length).Font
          .Name = "HGS明朝B"
          .Size = 11.5
        End With
      Next
    Next
'                                                          アポストロフィ フォント変更
    oRegExp.Pattern = cnApostro
    oRegExp.Global = True
    
    For Each oCell In Selection
      Set oMatchCl = oRegExp.Execute(oCell.Value)
      For Each oMatch In oMatchCl
        With oCell.Characters(Start:=oMatch.FirstIndex + 1, _
                              Length:=oMatch.Length).Font
          .Name = "CRＣ＆Ｇ行刻04"
          .Size = 12.5
        End With
      Next
    Next
'                                                          カッコ フォント変更
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
    MsgBox "セル領域が選択されていません", vbCritical
  End If
  
End Sub
