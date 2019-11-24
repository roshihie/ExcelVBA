Attribute VB_Name = "Module1"
Option Explicit
'*******************************************************************************
'        文字列取得処理
'*******************************************************************************
'      < 処理概要 >
'        指定されたエリアの文字列からブランクで区切られたそれぞれの文字列に
'        分解し、指定された配列番号の文字列を返す。
'*******************************************************************************
Public Function fn文字列取得(oRang As Range, sDlm As String, nPos As Integer) As String

  Dim asStr() As String
 
  asStr = Split(oRang, sDlm)
  fn文字列取得 = asStr(nPos)

End Function

