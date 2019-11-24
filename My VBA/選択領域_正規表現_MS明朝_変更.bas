Option Explicit
'*******************************************************************************
'        正規表現検索による 選択領域フォント ＭＳ明朝 変更
'*******************************************************************************
'      < 処理概要 >
'        正規表現により漢字を検索し、ヒットした漢字のフォントを ｢ＭＳ 明朝｣
'        に変更する
'*******************************************************************************
Public Sub 選択領域_正規表現_MS明朝_変更()

  Const cnKanji As String = "[一-龠]+|[ぁ-ん]+|[ァ-ヴ]+|[ａ-ｚＡ-Ｚ０-９]+|[「」｢｣]+"
  Dim oRegExp  As RegExp
  Dim oMatchCl As MatchCollection
  Dim oMatch   As Match
  Dim nSize    As Long
  Dim oCell    As Range

  Application.ScreenUpdating = False
  
  If TypeName(ActiveCell) = "Range" Then
    nSize = Application.InputBox(Prompt:="フォントサイズを入力する", _
                                 Title:="ＭＳ 明朝 フォントサイズ指定", _
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
          .Name = "ＭＳ 明朝"
          .Size = nSize
        End With
      Next
    Next
  Else
    MsgBox "セル領域が選択されていません", vbCritical
  End If
  
End Sub
