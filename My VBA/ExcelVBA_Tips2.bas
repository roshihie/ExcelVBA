Attribute VB_Name = "Module3"
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'        ＥｘｃｅｌＶＢＡ  Ｐｒｏｇｒａｍ
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'*******************************************************************************
'        コピーシート，オリジナルシート同一行  網掛処理
'*******************************************************************************
'        処理概要：オリジナルシートの各行のステータス列＝'必要*'のとき
'                  コピーシートの各行を比較して項目列のなかに同一の見出しが
'                  存在した場合は コピーシートの該当行に網掛けを行う
'*******************************************************************************
Sub ListShadowCreate1()

    Dim orgSheet     As Worksheet
    Dim cpySheet     As Worksheet
    Dim cpyRange     As Range

    Dim orgIdxRow    As Integer
    Dim orgStartRow  As Integer
    Dim orgStatusCol As Integer
    Dim orgNoCol     As Integer
    Dim orgNameCol   As Integer    '   比較対象の見出し列

    Dim cpyIdxRow    As Integer
    Dim cpyStartRow  As Integer
    Dim cpyNoCol     As Integer
    Dim cpyNameCol   As Integer    '   比較対象の見出し列

    Set orgSheet = Workbooks("Book1").Worksheets("Sheet1")
    Set cpySheet = Workbooks("Book2").Worksheets("Sheet1")

    orgStartRow = 5
    orgStatusCol = 10
    orgNoCol = 1
    orgNameCol = 2

    cpyStartRow = 3
    cpyNoCol = 1
    cpyNameCol = 2

    orgIdxRow = orgStartRow
    '   オリジナルシート行参照ループ
    Do While orgSheet.Cells(orgIdxRow, orgNoCol).Value <> ""
        If orgSheet.Cells(orgIdxRow, orgStatusCol).Value Like "必要*" Then
            cpyIdxRow = cpyStartRow
            '   コピーシート行参照ループ
            Do While cpySheet.Cells(cpyIdxRow, cpyNoCol).Value <> ""
                '   コピーシートの見出し列の先頭から+10列までを比較する
                cpySheet.Activate
                For Each cpyRange In cpySheet.Range(Cells(cpyIdxRow, cpyNameCol), _
                                                    Cells(cpyIdxRow, cpyNameCol + 10))
                    If cpyRange.Value = orgSheet.Cells(orgIdxRow, orgNameCol).Value Then
                        cpySheet.Activate
                        Range(Cells(cpyIdxRow, cpyNameCol), Cells(cpyIdxRow, 12)).Select
                        Selection.Interior.ColorIndex = 15
                    End If
                Next
                cpyIdxRow = cpyIdxRow + 1
            Loop
        End If
        orgIdxRow = orgIdxRow + 1
    Loop

End Sub

'*******************************************************************************
'        ステータスによる網掛処理
'*******************************************************************************
'        処理概要：シートの各行のステータス列＝'完了'のとき、グレー網掛け
'                  シートの各行のステータス列＝'*検討中'のとき、黄色網掛け
'                  を行う
'*******************************************************************************
Sub ListShadowCreate2()

    Dim ixRow       As Integer
    Dim startRow    As Integer
    Dim noCol       As Integer
    Dim startCol    As Integer
    Dim endCol      As Integer
    Dim statusCol   As Integer
    
    Dim cellValue   As String

    Workbooks("Book1").Worksheets("Sheet1").Activate
    
    startRow = 5
    noCol = 1
    startCol = 1
    statusCol = 13
    
    ixRow = startRow
    '   シートの行参照ループ
    Do While Cells(ixRow, noCol).Value <> ""
        cellValue = Cells(ixRow, statusCol).Value
        If cellValue = "完了" Then
            '   シートの最右端からアクティブセル領域の最右端＋２列を取得して選択
            endCol = ActiveSheet.Columns.Count
            Range(Cells(ixRow, startCol), Cells(ixRow, endCol).End(xlToLeft).Offset(, 2)).Select
            With Selection.Interior
                .ColorIndex = 15
                .Pattern = xlSolid
            End With
        Else
            If cellValue Like "*検討中" Then
                '   シートの最右端からアクティブセル領域の最右端＋２列を取得して選択
                endCol = ActiveSheet.Columns.Count
                Range(Cells(ixRow, startCol), Cells(ixRow, endCol).End(xlToLeft).Offset(, 2)).Select
                With Selection.Interior
                    .ColorIndex = 27
                    .Pattern = xlSolid
                End With
            End If
        End If
        ixRow = ixRow + 1
    Loop

End Sub

'*******************************************************************************
'        見出し項目コピー処理
'*******************************************************************************
'        処理概要：オリジナルシートの各行のステータス列＝"Y"のとき
'                  見出し項目をコピーシートにコピーする
'*******************************************************************************
Sub ListItemCopy()

'
'
    Dim orgSheet    As Worksheet
    Dim cpySheet    As Worksheet
    
    Dim ixRow       As Integer         'オリジナルシートの項目
    Dim startRow    As Integer
    Dim noCol       As Integer
    Dim targetCol   As Integer
    Dim statusCol   As Integer

    Dim selRow      As Integer         'コピーシートの項目
    
    Set orgSheet = Workbooks("Book1").Worksheets("Sheet1")
    Set cpySheet = Workbooks("Book2").Worksheets("Sheet2")
    
    startRow = 5
    noCol = 1
    tagetCol = 2
    statusCol = 24
    
    selRow = 0
    
    ixRow = startRow
    Do While orgSheet.Cells(ixRow, noCol).Value <> ""
        If orgSheet.Cells(ixRow, statusCol).Value = "Y" Then
            selRow = selRow + 1
            cpySheet.Activate
            Cells(selRow + 2, 2).Select
            Selection.Value = orgSheet.Cells(ixRow, tagetCol).Value
            '   セル書式設定  文字折返しなし
            Selection.WrapText = False
        End If
        ixRow = ixRow + 1
    Loop

End Sub

'*******************************************************************************
'        新規挿入ラインの初期化
'*******************************************************************************
'        処理概要：新規挿入するラインを一定の方式で初期化する
'
'*******************************************************************************
Sub NewLineSet()

'   新規挿入ラインを初期化する
    Dim lastCol         As Integer
    Dim userSelectCell  As Range

    'ActiveCellのバックアップ
    Set userSelectCell = ActiveCell
    
    '横幅の最終Column 取得（ヘッダー行）
    Cells.Find(What:="№", after:=Range("A1")).Select
    lastCol = Selection.End(xlToRight).Column

    '挿入行の№列に関数を設定
    userSelectCell.Select
    Selection.EntireRow.Select
    Selection.Insert Shift:=xlDown
    Cells(ActiveCell.row, ActiveCell.Column).Formula = _
        "=IF(LEN(B" & ActiveCell.row & ")>1,ROW()-ROW(A$1)-COUNTBLANK(A$2:A" & ActiveCell.row - 1 & "),"""")"

    '挿入行は塗りつぶしなし
    Range(Cells(ActiveCell.row, ActiveCell.Column), Cells(ActiveCell.row, lastCol)).Select
    Selection.Interior.ColorIndex = xlNone
        
    '挿入行＋１行の№列の関数 再設定 (∵COUNTBLANK(A$2:Annn) の nnn が１アップしない)
    '上記事象は、挿入行＋１行目のみ
    Cells(ActiveCell.row + 1, ActiveCell.Column).Formula = _
        "=IF(LEN(B" & ActiveCell.row + 1 & ")>1,ROW()-ROW(A$1)-COUNTBLANK(A$2:A" & ActiveCell.row & "),"""")"
    
    Cells(ActiveCell.row, 1).Select
    
End Sub

'*******************************************************************************
'        指定色の変更
'*******************************************************************************
'        処理概要：セルの色が黄色だったら水色に変更する
'
'*******************************************************************************
Sub InteriorColorChange()

'   黄色から水色に変更
    Workbooks("Book1").Worksheets("Sheet1").Activate
    With Application
        .FindFormat.Interior.ColorIndex = 6            ' 黄色
        .ReplaceFormat.Interior.ColorIndex = 8         ' 水色
    End With
    
    ActiveSheet.UsedRange.Replace _
        What:="", replacement:="", SearchFormat:=True, ReplaceFormat:=True
        
    With Application
        .FindFormat.Clear
        .ReplaceFormat.Clear
    End With
    
End Sub

'*******************************************************************************
'        重複行の削除コントロール
'*******************************************************************************
'        処理概要：下記 重複行の削除を CALL する
'
'*******************************************************************************
Sub DuplicateRowsDelCall()

    Dim strColumn  As String
    
    Workbooks("Book1").Worksheets("Sheet1").Activate
    
    strColumn = "A"
    Call DuplicateRowsDelete3(strColumn)

End Sub

'*******************************************************************************
'        重複行の削除
'*******************************************************************************
'        処理概要：指定されたシートの列位置にてソート(行位置＝1 固定)し
'                  上下同一のとき上の行を削除する(データがなくなるまで処理を行う)
'
'            戻り値　：なし
'            引数１　：シート名称       String
'            引数２　：カラム位置       Integer
'*******************************************************************************
Sub DuplicateRowsDelete3(strColumn As String)

    Dim strColumnRange As String
    Dim rngSortArea    As Range
    Dim rngCurrentCell As Range
    Dim rngNextCell    As Range
    
    strColumnRange = strColumn & "1"
    Set rngSortArea = Range(strColumnRange).CurrentRegion
    
    rngSortArea.sort Key1:=Range(strColumnRange), _
                     Order1:=xlAscending, _
                     Header:=xlGuess, _
                     OrderCustom:=1, _
                     MatchCase:=False, _
                     Orientation:=xlTopToBottom, _
                     SortMethod:=xlPinYin, _
                     DataOption1:=xlSortNormal
       
    Set rngCurrentCell = Range(strColumnRange)
    
    Do While Not IsEmpty(rngCurrentCell)
    
        Set rngNextCell = rngCurrentCell.Offset(1)
        If rngNextCell.Value = rngCurrentCell.Value Then
            rngCurrentCell.EntireRow.Delete
        End If
        
        Set rngCurrentCell = rngNextCell
    Loop
    
End Sub

'*******************************************************************************
'        指定セルコピー
'*******************************************************************************
'        処理概要：A列セルの情報と最右端セルの情報を取得する
'
'*******************************************************************************
sub
