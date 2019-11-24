Attribute VB_Name = "Module4"
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'        マクロ記録集
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'
Sub セル書式設定_文字列折返しなしMacro()
'
' セル書式設定_文字列折返しなし Macro
' マクロ記録日 : 2005/1/10  ユーザー名 :
'
    Range("A1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
    '   文字列折返しなし
        .WrapText = False
        
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub セル書式設定_文字列折返しありMacro()
'
' セル書式設定_文字列折返しあり Macro
' マクロ記録日 : 2005/1/10  ユーザー名 :
'
    Range("A1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
    '   文字列折返しあり
        .WrapText = True
        
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub セル書式設定_網掛け_解除Macro()
'
' セル書式設定_網掛け解除 Macro
' マクロ記録日 : 2005/1/10  ユーザー名 :
'
    Range("A1").Select
'   網掛け解除
    Selection.Interior.ColorIndex = xlNone
    
End Sub
Sub セル書式設定_網掛け_グレーMacro()
'
' セル書式設定_網掛け_グレー Macro
' マクロ記録日 : 2005/1/10  ユーザー名 :
'
    Range("A1").Select
    With Selection.Interior
    '   網掛け(グレー)
        .ColorIndex = 15
        
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
End Sub
Sub 関数Find使用方法Macro()
'
' Find関数使用方法 Macro
' マクロ記録日 : 2005/1/16  ユーザー名 :
'
' Find(What:=検索文字列,
'      After:=指定したセルの次から検索,
'      LookIn:=検索内容[数式[xlFormulas],
'                       値[xlValues],  ○
'                       コメント[xlComments]],
'      LookAt:=検索対象部分[一部分一致で検索[xlPart],  ○
'                           全て一致で検索[xlWhole]],
'      SearchOrder:=検索方向[列[xlByColumns],
'                            行[xlByRows]],  ○
'      SearchDirection:=行のとき左→右,列のとき上→下[xlNext],  ○
'                       行のとき右→左,列のとき下→上[xlPrevious]
'      MatchCase:=大文字･小文字を区別[True],
'                 大文字・小文字任意 [False]  ○
'      MatchByte:=全角･半角を区別[True],
'                 全角・半角任意 [False]  ○
    Workbooks("テストBook.xls").Activate
    Cells.Find(What:="取引金額", after:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, _
               SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
               MatchCase:=False, SearchFormat:=False).Activate
    Cells.FindNext(after:=ActiveCell).Activate
End Sub
Sub 行挿入Macro()
'
' 行挿入 Macro
' マクロ記録日 : 2005/2/27  ユーザー名 :
'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
End Sub
Sub 文字列→数値変換Macro()
'
' 文字列→数値変換 Macro
' マクロ記録日 : 2005/5/15  ユーザー名 :
'
    Range("K5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J7:J18").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
End Sub
Sub Sort処理Macro()
'
' Sort処理Macro
' マクロ記録日 : 2009/2/1  ユーザー名 :
'
    Range("A1:A27").sort Key1:=Range("A1"), _
                         Order1:=xlAscending, _
                         Header:=xlGuess, _
                         OrderCustom:=1, _
                         MatchCase:=False, _
                         Orientation:=xlTopToBottom, _
                         SortMethod:=xlPinYin, _
                         DataOption1:=xlSortNormal
End Sub

