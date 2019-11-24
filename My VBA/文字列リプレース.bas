Option Explicit
'*******************************************************************************
'        文字列リプレース
'*******************************************************************************
'      < 処理概要 >
'        アクティブシートのセル内容について
'        文字列変換一覧（専用ブック 専用シート）に記載された変更前文字を
'        変更後文字に変更する
'*******************************************************************************
Public Sub 文字列リプレース()
   
    Dim oDic As Object
    Set oDic = CreateObject("Scripting.Dictionary")        'Dictionaryオブジェクトの宣言
    
    Dim i  As Integer
    With Workbooks("文字列変換定義.xlsx").Worksheets("変換一覧")
      For i = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
        oDic.Add .Cells(i, 1).Value, .Cells(i, 2).Value    'Dictionaryオブジェクトの初期化、要素の追加
      Next
    End With
    
    Dim bRslt        As Boolean
    Dim oSoc, oSoces As Range
    Dim sRep         As Variant
        
    Set oSoces = Range(Cells(1, 1), Cells(Rows.Count, 1).End(xlUp))
        
    For Each oSoc In oSoces
      For Each sRep In oDic                                'Dictionaryオブジェクトを使った複数条件の置換
        bRslt = oSoc.Replace(What:=sRep, Replacement:=oDic(sRep), LookAt:=xlPart)
      Next
    Next
End Sub
