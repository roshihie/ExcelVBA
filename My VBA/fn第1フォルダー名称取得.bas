Option Explicit
'*******************************************************************************
'        第１フォルダー名称取得
'*******************************************************************************
'      < 処理概要 >
'        パス名から第１フォルダー名称を取得する
'*******************************************************************************
Function fn第1フォルダー名称取得(oRange As Range, sDlmt As String) As String

  Dim asFolder() As String
  Dim nToIdx     As Integer
  Dim i          As Integer
  
  asFolder = Split(oRange, sDlmt)
  
  If Left(oRange, 2) = "\\" Then
    nToIdx = LBound(asFolder) + 2
  Else
    nToIdx = LBound(asFolder)
  End If
  
  fn第1フォルダー名称取得 = asFolder(LBound(asFolder))
  For i = LBound(asFolder) + 1 To nToIdx
    fn第1フォルダー名称取得 = fn第1フォルダー名称取得 & "\" & asFolder(i)
  Next

End Function

