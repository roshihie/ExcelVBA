Option Explicit
'*******************************************************************************
'        選択領域フォント ＭＳ明朝 変更
'*******************************************************************************
'      < 処理概要 >
'        選択領域の文字フォントを「ＭＳ 明朝」に変更する
'*******************************************************************************
Public Sub 選択領域_MS明朝_変更()

  Dim nSize  As Long
  
  If TypeName(Selection) = "Range" Then
    nSize = Application.InputBox(Prompt:="フォントサイズを入力する", _
                                 Title:="ＭＳ 明朝 フォントサイズ指定", _
                                 Default:=11, _
                                 Type:=1)
    If nSize = 0 Then
      Exit Sub
    End If
    
    With Selection.Font
      .Name = "ＭＳ 明朝"
      .Size = nSize
      .Bold = False
    End With
  Else
    MsgBox "セル領域が選択されていません", vbCritical
  End If
  
End Sub
