Option Explicit
'*******************************************************************************
'        最終フォルダー名称取得
'*******************************************************************************
'      < 処理概要 >
'        パス名から最終フォルダー名称を取得する
'*******************************************************************************
Function fn最終フォルダー名称取得(oRange As Range, sDlmt As String) As String

  Dim asFolder() As String
  Dim nToIdx     As Integer
  Dim i          As Integer
  
  asFolder = Split(oRange, sDlmt)
  fn最終フォルダー名称取得 = asFolder(UBound(asFolder))

End Function

