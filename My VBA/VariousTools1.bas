Attribute VB_Name = "Module1"
Option Compare Text
Sub 取消線除去()
    '選択セル内の取消線を除くマクロ
    '2004.08.30
    
    'ステータスバーの保存
    ORG_BAR = Application.DisplayStatusBar
    Application.StatusBar = True
    '選択セルのアドレスを取得
    開始列 = Selection.Cells.Column
    開始行 = Selection.Cells.Row
    Set myRange = Selection.Cells
    終了列 = myRange.Columns(myRange.Columns.Count).Column
    終了行 = myRange.Rows(myRange.Rows.Count).Row
    'ステータスバーを表示
    Application.StatusBar = Cells(開始行, 開始列).Address & "から" & Cells(終了行, 終了列).Address & "までを選択しています"
    '一括置換か否か問い合わせ
    resp = MsgBox("一括置換？", vbYesNo)
    '選択セルを順番に参照
    For m = 開始行 To 終了行
        For n = 開始列 To 終了列
            'ステータスバーを表示
            Application.StatusBar = Cells(m, n).Address & "を処理中"
            j = ""
            k = 0
            'セル内に取消線があるとき、取消線を除いた値を抽出
            For i = 1 To Len(Cells(m, n)) Step 1
                If Cells(m, n).Characters(i, 1).Font.Strikethrough = False Then
                    j = j + Mid(Cells(m, n), i, 1)
                Else
                    k = k + 1
                End If
            Next
            '取消線が合った場合のみ置換する値を問い合わせ
            If k > 0 Then
                '一括置換でないのときのみ、変換前後を確認
                If resp = vbNo Then
                    j = InputBox(m & "行" & n & "列の変更前：" & Cells(m, n), , j)
                End If
                Cells(m, n) = j
                Cells(m, n).Font.Strikethrough = False
            End If
        Next
    Next
    'ステータスバーの復帰
    Application.StatusBar = False
    ORG_BAR = Application.DisplayStatusBar
    '終了メッセージの出力
    MsgBox "Finished"
End Sub
Sub 検索色付()
    '選択セル内の文字列検索をして色を付けるマクロ
    '2004.11.26
    On Error Resume Next
    'ステータスバーの保存
    ORG_BAR = Application.DisplayStatusBar
    Application.StatusBar = True
    '選択セルのアドレスを取得
    開始列 = Selection.Cells.Column
    開始行 = Selection.Cells.Row
    Set myRange = Selection.Cells
    終了列 = myRange.Columns(myRange.Columns.Count).Column
    終了行 = myRange.Rows(myRange.Rows.Count).Row
    'ステータスバーを表示
    Application.StatusBar = Cells(開始行, 開始列).Address & "から" & Cells(終了行, 終了列).Address & "までを選択しています"
    '検索文字列を問い合わせ
    response1 = InputBox("検索文字列を指定してください" & vbCrLf & _
                         "大文字小文字全角半角は区別されます" & vbCrLf & _
                         "「*,?」はそのまま指定してください", , "(検索文字列)")
    '検索文字列が（NULL）の場合処理中止
    If response1 = "" Then
        MsgBox "処理を中止します"
        Exit Sub
    End If
    response2 = InputBox("文字色を指定してください" & vbCrLf & _
                         "例：赤[3]青[5]桃[7]緑[4]等", response1 & " を検索します", "(文字色1～56)")
    '文字色が指定範囲外の場合処理中止
    If response2 < 1 Or _
       response2 > 56 Then
        MsgBox "処理を中止します"
        Exit Sub
    End If
    j = 0
    '選択セルを順番に参照
    For m = 開始行 To 終了行
        For n = 開始列 To 終了列
            'ステータスバーを表示
            Application.StatusBar = Cells(m, n).Address & "を処理中"
            '検索開始位置の初期値設定
            i = 1
            'セルの値をワークへ格納
            SearchString = Cells(m, n).Value
            'セルの値の長さ分だけループ
            Do While i <= Len(SearchString)
                'セルの中に検索文字列があるか検査
                If InStr(i, SearchString, response1, vbBinaryCompare) > 0 Then
                    'セルの中に検索文字列がある位置を記憶
                    response3 = InStr(i, SearchString, response1, vbBinaryCompare)
                    '検索文字列がある位置から検索文字列の長さ分だけ書式設定
                    '色変更
                    Cells(m, n).Characters(response3, Len(response1)).Font.ColorIndex = response2
                    'ボールド
                    Cells(m, n).Characters(response3, Len(response1)).Font.Bold = True
                    '文字列検索の開始位置をシフトする
                    i = response3 + Len(response1)
                    j = j + 1
                Else
                    '検索文字列が見つからなかったので、文字列検索の開始位置を最大にシフトする
                    i = i + Len(SearchString)
                End If
            Loop
        Next
    Next
    'ステータスバーの復帰
    Application.StatusBar = False
    ORG_BAR = Application.DisplayStatusBar
    '終了メッセージの出力
    MsgBox "Finished" & vbCrLf & _
           "Changed × " & j
End Sub
Sub 改行削除()
    'カウンタ初期化
    l = 0
    '使用セル選択
    ActiveSheet.UsedRange.Select
    '使用セルの範囲アドレス取得
    開始列 = Selection.Cells.Column
    開始行 = Selection.Cells.Row
    Set myRange = Selection.Cells
    終了列 = myRange.Columns(myRange.Columns.Count).Column
    終了行 = myRange.Rows(myRange.Rows.Count).Row
    '使用セルだけループ
     For i = 開始行 To 終了行
        For j = 開始列 To 終了列
            'ワークエリア初期化
            aaa = ""
            '１文字ずつ検査して改行コードがある場合無視し、
            'それ以外の場合移送する
            For k = 1 To Len(Cells(i, j))
                If Mid(Cells(i, j), k, 1) = Chr(10) Or _
                   Mid(Cells(i, j), k, 1) = Chr(13) Then
                    l = l + 1
                Else
                    aaa = aaa & Mid(Cells(i, j), k, 1)
                End If
            Next
            Cells(i, j) = aaa
        Next
    Next
    '終了メッセージを出力する
    MsgBox "改行削除× " & l
End Sub
Sub CSV作成()
    '選択セルをCSV化するマクロ
    '2004.12.10
    
    Dim 出力レコード() As String
    '選択セルのアドレスを取得
    開始列 = Selection.Cells.Column
    開始行 = Selection.Cells.Row
    Set myRange = Selection.Cells
    終了列 = myRange.Columns(myRange.Columns.Count).Column
    終了行 = myRange.Rows(myRange.Rows.Count).Row
    '選択セルを順番に参照
    p = 0
    For m = 開始行 To 終了行
        p = p + 1
        ReDim Preserve 出力レコード(p)
        For n = 開始列 To 終了列
            出力レコード(p) = 出力レコード(p) & """" & Cells(m, n) & ""","
        Next
    Next
    レコード数 = p
    出力ファイル = Application.GetSaveAsFilename(fileFilter:="テキスト ファイル (*.csv), *.csv")
    ' 出力ファイルをオープンする
    OutputFile = FreeFile
    Open 出力ファイル For Output As #OutputFile
    For p = 1 To レコード数
        Print #OutputFile, 出力レコード(p)
    Next
    ' 出力ファイルをクローズする
    Close #OutputFile
    '終了メッセージの出力
    MsgBox "Finished"
End Sub
Sub テキスト罫線作成()
    Dim 列幅() As Variant
    '選択セルをテキスト罫線表形式に変換するマクロ
    '2004.12.10
    
    '選択セルのアドレスを取得
    開始列 = Selection.Cells.Column
    開始行 = Selection.Cells.Row
    Set myRange = Selection.Cells
    終了列 = myRange.Columns(myRange.Columns.Count).Column
    終了行 = myRange.Rows(myRange.Rows.Count).Row
    '列幅のカウント
    総列数 = 終了列 - 開始列 + 1
    ReDim Preserve 列幅(総列数)
    '列幅初期値設定
    For p = 1 To 総列数
        列幅(p) = 2
    Next
    '選択セルを順番に参照し列幅を記憶する
    For m = 開始行 To 終了行
        p = 0
        For n = 開始列 To 終了列
            p = p + 1
            If 列幅(p) < Len(Cells(m, n).Value) Then
                列幅(p) = Len(Cells(m, n).Value)
                If 列幅(p) Mod 2 > 0 Then
                    列幅(p) = 列幅(p) + 1
                End If
            End If
        Next
    Next
    出力ファイル = Application.GetSaveAsFilename(fileFilter:="テキスト ファイル (*.txt), *.txt")
    ' 出力ファイルをオープンする
    OutputFile = FreeFile
    Open 出力ファイル For Output As #OutputFile
    '１行目
    出力レコード = "┏"
    For p = 1 To 総列数 Step 1
        For q = 1 To (列幅(p) / 2) Step 1
            出力レコード = 出力レコード & "━"
        Next
        If p = 総列数 Then
            出力レコード = 出力レコード & "┓"
        Else
            出力レコード = 出力レコード & "┳"
        End If
    Next
    Print #OutputFile, 出力レコード
    '繋ぎ
    出力レコード = "┣"
    For p = 1 To 総列数 Step 1
        For q = 1 To (列幅(p) / 2) Step 1
            出力レコード = 出力レコード & "━"
        Next
        If p = 総列数 Then
            出力レコード = 出力レコード & "┫"
        Else
            出力レコード = 出力レコード & "╋"
        End If
    Next
    '繋ぎレコードをセーブ
    繋ぎレコード = 出力レコード
    '選択セルを順番に参照
    For m = 開始行 To 終了行
        出力レコード = "┃"
        p = 0
        For n = 開始列 To 終了列
            p = p + 1
            セル内テキスト = Cells(m, n)
            Do While Len(セル内テキスト) < 列幅(p)
                セル内テキスト = セル内テキスト & " "
            Loop
            出力レコード = 出力レコード & セル内テキスト & "┃"
        Next
        Print #OutputFile, 出力レコード
        If m = 終了行 Then
        Else
            出力レコード = 繋ぎレコード
            Print #OutputFile, 出力レコード
        End If
    Next
    '最終行
    出力レコード = "┗"
    For p = 1 To 総列数 Step 1
        For q = 1 To (列幅(p) / 2) Step 1
            出力レコード = 出力レコード & "━"
        Next
        If p = 総列数 Then
            出力レコード = 出力レコード & "┛"
        Else
            出力レコード = 出力レコード & "┻"
        End If
    Next
    Print #OutputFile, 出力レコード
    ' 出力ファイルをクローズする
    Close #OutputFile
    '終了メッセージの出力
    MsgBox "Finished"
End Sub
