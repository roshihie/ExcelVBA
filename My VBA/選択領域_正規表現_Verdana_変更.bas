Option Explicit
'*******************************************************************************
'        正規表現検索による 選択領域フォント Verdana 変更
'*******************************************************************************
'      < 処理概要 >
'        正規表現により英字を検索し、ヒットした英字のフォントを ｢Verdana｣
'        に変更する
'*******************************************************************************
Public Sub 選択領域_正規表現_Verdana_変更()

  'Const cnEnglish As String = "[A-Z,a-z,?,!,;,:,',.,""]+"
  Const cnEnglish As String = "[A-Z]+|[a-z]+|[?]+|[!]+|[;]+|[:]+|[']+|[.]+|[""]+"
  Dim oRegExp  As RegExp
  Dim oMatchCl As MatchCollection
  Dim oMatch   As Match
  Dim nSize    As Long
  Dim oCell    As Range
  
  Application.ScreenUpdating = False
  
  If TypeName(Selection) = "Range" Then
    nSize = Application.InputBox(Prompt:="フォントサイズを入力する", _
                                 Title:="Verdana フォントサイズ指定", _
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
    MsgBox "セル領域が選択されていません", vbCritical
  End If
  
End Sub
