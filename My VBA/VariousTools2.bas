Attribute VB_Name = "Module2"
Sub セルコメント一覧作成()
    '選択セル内のコメント一覧を作成するマクロ
    '2004.10.29
    
    'ステータスバーの保存
    ORG_BAR = Application.DisplayStatusBar
    Application.StatusBar = True
    '選択セルのアドレスを取得
    開始列 = Selection.Cells.Column
    開始行 = Selection.Cells.Row
    Set myRange = Selection.Cells
    終了列 = myRange.Columns(myRange.Columns.Count).Column
    終了行 = myRange.Rows(myRange.Rows.Count).Row
    'アクティブブックの情報を保存
    Myworkbook = ActiveWorkbook.Name
    Myworksheet = ActiveSheet.Name
    '新規ブックの追加
    Workbooks.Add
    '新規ブックのシートを１枚だけにする
    Call シート削除
    '新規ブックのシート名を変更する
    Worksheets(1).Activate
    Worksheets(1).Name = Workbooks(Myworkbook).Worksheets(Myworksheet).Name
    '一覧ヘッダ編集とシートの書式設定
    ActiveSheet.Cells(1, 1) = "相対アドレス"
    ActiveSheet.Cells(1, 2) = "行番号"
    ActiveSheet.Cells(1, 3) = "列番号"
    ActiveSheet.Cells(1, 4) = "セルの内容"
    ActiveSheet.Cells(1, 5) = "コメントの内容"
    Columns("A:E").VerticalAlignment = xlTop
    '一覧編集開始行設定
    行 = 2
    For m = 開始行 To 終了行
        For n = 開始列 To 終了列
            'ステータスバーを更新
            Application.StatusBar = Cells(m, n).Address & "を処理中"
            'セルにコメントが含まれているか問合せ
            If TypeName(Workbooks(Myworkbook).Worksheets(Myworksheet).Cells(m, n).Comment) = "Nothing" Then
            Else
                'セルにコメントが含まれている場合、一覧を編集する
                絶対アドレス = Workbooks(Myworkbook).Worksheets(Myworksheet).Cells(m, n).Address
                '絶対アドレスを相対アドレスに編集
                Call 相対アドレス取得(絶対アドレス)
                'セル情報を編集
                ActiveSheet.Cells(行, 1) = 絶対アドレス
                ActiveSheet.Cells(行, 2) = m
                ActiveSheet.Cells(行, 3) = n
                ActiveSheet.Cells(行, 4) = Workbooks(Myworkbook).Worksheets(Myworksheet).Cells(m, n)
                ActiveSheet.Cells(行, 5) = Workbooks(Myworkbook).Worksheets(Myworksheet).Cells(m, n).Comment.Text
                '一覧行インクリメント
                行 = 行 + 1
            End If
        Next
    Next
    '一覧行の行と列の幅を整える
    Columns("A:E").EntireColumn.AutoFit
    Rows.EntireRow.AutoFit
    'ステータスバーの復帰
    Application.StatusBar = False
    ORG_BAR = Application.DisplayStatusBar
End Sub

Sub シート削除()
    '１枚のシートのみにする内部モジュール
    Do While Sheets.Count > 1
        Application.DisplayAlerts = False
        Sheets(1).Delete
        Application.DisplayAlerts = True
    Loop
End Sub

Sub 相対アドレス取得(絶対アドレス)
    '絶対アドレスから＄を削除する内部モジュール
    相対アドレス = ""
    For i = 1 To Len(絶対アドレス) Step 1
        If Mid(絶対アドレス, i, 1) = "$" Then
        Else
            相対アドレス = 相対アドレス & Mid(絶対アドレス, i, 1)
        End If
    Next
    絶対アドレス = 相対アドレス
End Sub

Sub シート一覧作成()
    'アクティブシートのシート一覧を作成する
    '2004.12.01
    
    'アクティブシートのブック名保管
    ブック名 = ActiveWorkbook.Name
    '新規ブック追加
    Workbooks.Add
    '一枚のシートのみにする
    Call シート削除
    'タイトル行の編集
    ActiveWorkbook.Sheets(1).Cells(1, 1) = "ブック名"
    ActiveWorkbook.Sheets(1).Cells(1, 2) = "シート名"
    ActiveWorkbook.Sheets(1).Cells(1, 3) = "備考"
    ActiveWorkbook.Sheets(1).Cells(2, 1) = ブック名
    '存在するシートの数だけループ
    For i = 1 To Workbooks(ブック名).Sheets.Count
        'シート名
        ActiveWorkbook.Sheets(1).Cells(i + 1, 2) = Workbooks(ブック名).Sheets(i).Name
        '非表示シートかどうか検査
        If Workbooks(ブック名).Sheets(i).Visible = xlSheetVisible Then
        Else
            ActiveWorkbook.Sheets(1).Cells(i + 1, 3) = "非表示"
        End If
        '空きシートかどうか検査
        If Workbooks(ブック名).Sheets(i).UsedRange.Address = "$A$1" And _
           Workbooks(ブック名).Sheets(i).Range("A1") = "" Then
            ActiveWorkbook.Sheets(1).Cells(i + 1, 3) = _
            ActiveWorkbook.Sheets(1).Cells(i + 1, 3) & "（空き）"
        End If
    Next
    'カラムあわせ
    ActiveWorkbook.Sheets(1).Columns("A:C").EntireColumn.AutoFit
End Sub

Sub シートリンク一覧作成()
    'アクティブシートのシート一覧を作成する（ハイパーリンク付）
    '2005.01.06
    'シート追加するか否か尋ねる
    Res = MsgBox("一覧用シート追加", vbYesNo)
    If Res = vbYes Then
        Sheets.Add Before:=Sheets(1)
    End If
    For i = 2 To ActiveWorkbook.Sheets.Count
        Sheets(1).Select
        Cells(i, 2).Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", _
        SubAddress:=Sheets(i).Name & "!A1", TextToDisplay:=Sheets(i).Name
        '各シートに「Return」ハイパーリンクを付ける
        'Sheets(i).Select
        'Cells(1, Sheets(i).UsedRange.Column + 1).Select
        'ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", _
        'SubAddress:=Sheets(1).Name & "!A1", TextToDisplay:="Return"
    Next
    Sheets(1).Select
End Sub

Sub シート追加()
    '既存シート名格納配列
    Dim 既存シート名()
    '基準文字列の問合せ
    基準文字列 = InputBox("基準文字列", , ActiveSheet.Name)
    '基準文字列が無いとき、処理中断
    If 基準文字列 = "" Then
        Exit Sub
    End If
    '繰り返し数の問合せ
    繰り返し数 = InputBox("繰り返し数", , "1")
    '繰り返し数が無いとき、処理中断
    If 繰り返し数 = "" Then
        Exit Sub
    End If
    '既存シート名の取得
    For j = 1 To ActiveWorkbook.Sheets.Count
        ReDim Preserve 既存シート名(j)
        既存シート名(j) = ActiveWorkbook.Sheets(j).Name
    Next
    既存シート数 = ActiveWorkbook.Sheets.Count
    '繰り返し数分、ループ
    For i = 1 To 繰り返し数 Step 1
        'アクティブシートの後にシート追加
        Sheets.Add After:=ActiveSheet
        'シート名を仮設定
        シート名 = 基準文字列 & i
        '仮設定したシート名が既存シート名と重複していないかチェック
        Call 既存シート比較(シート名, 既存シート名(), 既存シート数)
        '重複しないシート名で更新
        ActiveSheet.Name = シート名
    Next
End Sub

Sub 既存シート比較(シート名, 既存シート名(), 既存シート数)
'ラベル定義
level1:
    'シート名が既存シート名と重複しないまでループ
    For j = 1 To 既存シート数
        If シート名 = 既存シート名(j) Then
            シート名 = シート名 & "@"
            GoTo level1
        End If
    Next
End Sub

Sub シート名変更()
    'シート名を基準文字列＋連番で変更する
    '乱数発生ルーチンを初期化します。
    Randomize
    '基準文字列を問合せ
    基準文字列 = InputBox("基準文字列", , "Sheet1")
    '基準文字列がNULLでも実行する可能性があるため、実行有無を問合せ
    If MsgBox("実行する", vbYesNo) = vbNo Then
        Exit Sub
    End If
    '乱数発生
    乱数 = Int((9999 * Rnd) + 1)
    '一度シート名をランダムに変更
    For i = 1 To ActiveWorkbook.Sheets.Count
        Sheets(i).Name = 乱数 & i
    Next
    'その後基準文字列＋連番でシート名変更
    For i = 1 To ActiveWorkbook.Sheets.Count
        Sheets(i).Name = 基準文字列 & i
    Next
End Sub

Sub シート指定削除()
    'アクティブシート以外を削除
    '2004.12.06
    
    'アクティブブックのシート数が１のとき、処理終了
    If ActiveWorkbook.Sheets.Count = 1 Then
        Exit Sub
    End If
    'インデックス１のシート名とアクティブシート名が同一のとき
    'インデックス２以降を削除
    If Sheets(1).Name = ActiveSheet.Name Then
        Do While ActiveWorkbook.Sheets.Count > 1
            Application.DisplayAlerts = False
            Sheets(2).Delete
            Application.DisplayAlerts = True
        Loop
    'アクティブシートのインデックスが１でないとき
    Else
        'アクティブシートのインデックスを１にするためにシート移動
        ActiveSheet.Move Before:=Sheets(1)
        Do While ActiveWorkbook.Sheets.Count > 1
            Application.DisplayAlerts = False
            Sheets(2).Delete
            Application.DisplayAlerts = True
        Loop
    End If
End Sub
