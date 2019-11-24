Attribute VB_Name = "Module2"
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'        ＥｘｃｅｌＶＢＡ  Ｔｉｐｓ
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'*******************************************************************************
'        データ最下行に SUM関数を挿入
'*******************************************************************************
Sub InputFormula()

    Dim i         As Integer
    Dim myLastClm As Integer
    Dim r         As Long
   '最終列を取得
    myLastClm = Range("A1").End(xlToRight).Column
    
   '最終行を取得
    r = Range("A1").End(xlDown).row
    
   'SUM関数を最終行の下に入力
    For i = 1 To myLastClm
        Cells(r + 1, i).FormulaR1C1 = "=SUM(R[" & -r & "]C:R[-1]C)"
    Next

End Sub

'*******************************************************************************
'        新規入力行を取得する
'*******************************************************************************
Sub NewInputRowGet()

    Dim lastRow  As Long
    
    lastRow = ActiveSheet.Rows.Count
    Cells(lastRow, 1).End(xlUp).Offset(1).Select
    
End Sub

'*******************************************************************************
'        選択セルのサイズ変更
'*******************************************************************************
Sub CellResize()

    Workbooks("Book1").Worksheets("Sheet1").Activate
    Range("b2:c5").Select
    MsgBox "選択セル範囲のサイズを変更し、位置を２行下にずらします"
    Selection.Offset(2).Resize(Selection.Rows.Count + 2, Selection.Columns.Count + 3).Select
    
End Sub

'*******************************************************************************
'        アクティブセル領域全体をSelect
'*******************************************************************************
'        説　明　：アクティブセル領域＝空白行と空白列で囲まれたセル範囲＝データベース
'                  アクティブセル領域全体をSelect ⇒ CurrentRegionによるSelect
'*******************************************************************************
Sub CurrentRegionSelect()

    Range("A1").CurrentRegion.Select
    
'   データベースのデータ件数を算出
    Dim myData件数 As Long
    myData件数 = Range("A1").CurrentRegion.Rows.Count - 1   ' 見出し行を減算している
    
End Sub

'*******************************************************************************
'        データベースを右隣のワークシートにコピー
'*******************************************************************************
Sub Database()
    
'   データベース全体を選択する
    Range("A1").CurrentRegion.Select
    
'   データベースから行1（見出し行）を除いたデータ範囲をコピーする
    Selection.Offset(1).Resize(Selection.Rows.Count - 1).Copy
    
'   右隣のワークシートをアクティブにする
    ActiveSheet.Next.Activate
    
'   コピーしたセル範囲を貼り付ける
    Range("A1").PasteSpecial

'   コピーモードを解除する
    Application.CutCopyMode = False

End Sub

'*******************************************************************************
'        データベース内の特定行を選択する
'*******************************************************************************
Sub databaseRowSelect()

    Range("A6", Range("A6").End(xlToRight)).Select
    
End Sub

'*******************************************************************************
'        データベースに外枠の罫線を引く
'*******************************************************************************
Sub DrawLineOfDatabese()

    With Range("B3").CurrentRegion
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
    End With
    
End Sub

'*******************************************************************************
'        使用されているセル範囲全体をSelect
'*******************************************************************************
'        説　明　：使用されているセル範囲全体をSelect ⇒ UsedRangeによるSelect
'                  UsedRangeプロパティは「Range("A1").CurrentRegion」のように
'                  選択の基準となるセルを意識する必要はない
'*******************************************************************************
Sub UsedRangeSelect()

    ActiveSheet.UsedRange.Select

End Sub

'*******************************************************************************
'        重複行の削除（上方向へのループ）
'*******************************************************************************
'        処理概要：セルAn(A列の最終行)からスタートし 順次 セルAn～A2 の内容を比較し
'                  上下同一のとき下の行を削除する
'*******************************************************************************
Sub DuplicateRowsDelete1()
    
    Dim lastRow   As Long
    Dim myLastRow As Long
    Dim i         As Long
    
'   画面のちらつきを抑止して実行速度を向上させる
    Application.ScreenUpdating = False
'   アクティブシートの最終行を取得
    lastRow = ActiveSheet.Rows.Count
'   データ行の最終行を取得
    myLastRow = Cells(lastRow, 1).End(xlUp).row
    
    For i = myLastRow To 3 Step -1
        If Cells(i, 1).Value = Cells(i - 1, 1).Value Then
            Cells(i, 1).EntireRow.Delete
        End If
    Next i
    
End Sub

'*******************************************************************************
'        重複行の削除（下方向へのループ）
'*******************************************************************************
'        処理概要：セルA2からスタートし 順次 セルA2～An の内容を比較し
'                  上下同一のとき上の行を削除する
'
'                  myLastRowには削除前のデータ格納最終行数 が格納されるが
'                  これは 処理を行うべき行数 である
'                  同一行のとき 行は削除されるが 削除後の現在行数とは関係ない
'*******************************************************************************
Sub DuplicateRowsDelete2()

    Dim lastRow   As Long
    Dim myLastRow As Long
    Dim i         As Long
    
'   画面のちらつきを抑止して実行速度を向上させる
    Application.ScreenUpdating = False
'   アクティブセルの最終行を取得
    lastRow = ActiveSheet.Rows.Count
'   データ行の最終行を取得
    myLastRow = Cells(lastRow, 1).End(xlUp).row
    'Debug.Print "myLastRow = (" & myLastRow & ")"
    Range("A2").Select
    
    For i = 2 To myLastRow
        If Selection.Value = Selection.Offset(1).Value Then
            Selection.EntireRow.Delete
        Else
            Selection.Offset(1).Select
        End If
    Next i
    'Debug.Print "myLastRow = (" & myLastRow & ")"
    
End Sub

'*******************************************************************************
'        アクティブセル領域の中の可視セルだけをコピーする
'*******************************************************************************
Sub CopyVisibleRange()

'   アクティブセル領域を選択する
    Range("A1").CurrentRegion.Select

'   可視セルだけをコピーする
    Selection.SpecialCells(xlCellTypeVisible).Copy

'   別のシートに貼り付ける
    Worksheets("Sheet2").Select
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
End Sub

'*******************************************************************************
'        オートフィルタされた該当データの件数を取得
'*******************************************************************************
'        説　明　：Rowsプロパティには、複数の領域の中で最初の領域の行数しか参照できない
'                  VBAは「領域」という隣接しているセル範囲は Areasコレクション として
'                  扱われる
'*******************************************************************************
Sub CountSelectedData()

    Dim myArea As Range
    Dim myRow  As Integer
    '全ての可視状態の領域を選択する
    Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Select

   '各領域ごとに行数を取得して加算する
    For Each myArea In Selection.Areas
        myRow = myRow + myArea.Rows.Count
    Next

   '見だし行分を除くため 減算する
    MsgBox "抽出件数は " & (myRow - 1) & "件です"
    
End Sub

'*******************************************************************************
'        データが入力された最終セル(UsedRangeの最終セル)をSelect
'*******************************************************************************
Sub SelectLastCell()

    Dim myLastCell As Range
    Dim lastRow    As Long
    Dim lastCol    As Integer
    
    Set myLastCell = Range("A1").SpecialCells(xlCellTypeLastCell)
    
    With myLastCell
        If .row = 1 And .Column = 1 Then
            myLastCell.Select
            Exit Sub
        End If
    End With
    
    'SpecialCells(xlCellTypeLastCell)メソッドは一度データを入力したセルをクリアしても
    'そのまま最後のセルと認識してしまうため 補正する
    If myLastCell.Value = "" Then
    
'       Find(◎What:=検索文字列,
'            After:=指定したセルの次から検索,
'            LookIn:=検索内容[  数式[xlFormulas],
'                             ○値[xlValue],
'                               コメント[xlComments]],
'            LookAt:=検索対象部分[○一部分一致で検索[xlPart],
'                                   全て一致で検索[xlWhole]],
'            SearchOrder:=検索方向[  列[xlByColumns],
'                                  ○行[xlByRows]],
'            SearchDirection:=○(検索方向)行のとき左→右,(検索方向)列のとき上→下[xlNext],
'                               (検索方向)行のとき右→左,(検索方向)列のとき下→上[xlPrevious]
'            MatchCase:=  大文字･小文字を区別[True],
'                       ○大文字・小文字任意 [False]
'            MatchByte:=  全角･半角を区別[True],
'                       ○全角・半角任意 [False]
        lastRow = Cells.Find(What:="*", after:=myLastCell, _
                             SearchOrder:=xlByRows, _
                             SearchDirection:=xlPrevious).row
        lastCol = Cells.Find(What:="*", after:=myLastCell, _
                             SearchOrder:=xlByColumns, _
                             SearchDirection:=xlPrevious).Column
        Cells(lastRow, lastCol).Select
    Else
        myLastCell.Select
    End If
    
End Sub

'*******************************************************************************
'        列の表示／非表示を切り替える
'*******************************************************************************
Sub ToggleColumn()
    
    With Worksheets("Sheet1").Columns("C:D")
        .Hidden = Not .Hidden
    End With
    
End Sub

'*******************************************************************************
'        選択されているセル範囲の上にテキストボックス作成
'*******************************************************************************
Sub DrawTextbox()

'   選択セル範囲の開始位置
    Dim myLeft   As Variant
    Dim myTop    As Variant
'   選択セル範囲の大きさ
    Dim myWidth  As Variant
    Dim myHeight As Variant
    
    myLeft = Selection.Left
    myTop = Selection.Top
    myWidth = Selection.Width
    myHeight = Selection.Height
        
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        myLeft, myTop, myWidth, myHeight).Select
        
End Sub

'*******************************************************************************
'        データコピーによる新規ブック作成
'*******************************************************************************
'        処理概要：新規ブックを作成して既存ブックからデータをコピーし
'                  名前を付けて保存する
'*******************************************************************************
Sub LookNewBook()
    
    Dim myNewWK As Workbook
    
'   新規ブックを作成し その参照をオブジェクト型変数に代入する
    Set myNewWK = Workbooks.Add
    
'   Addメソッド実行後 追加された新規ブックがアクティブとなるので以下の通り 取得も可能
'   myWSName = ActiveWorkbook.Name
    
'   マクロを実行しているブック(Sample.xls)をアクティブにする
    Workbooks("Sample.xls").Activate
    Worksheets(1).Range("A1:E10").Copy
    
'   オブジェクト型変数を利用して新規ブックをアクティブにする
    myNewWK.Activate
    Worksheets(1).Activate
    ActiveSheet.Paste
    
'   新規ブックを名前を付けて保存して閉じる
    myNewWK.SaveAs "NewBook.xls"
    Workbooks("NewBook.xls").Close

    Application.CutCopyMode = False
    
End Sub

'*******************************************************************************
'        ブック変更時 上書き保存
'*******************************************************************************
Sub CheckSaved()

'   ブックが変更されたとき SavedプロパティにFalseがセット
    If ActiveWorkbook.Saved = False Then
       ActiveWorkbook.Save
    End If
    
End Sub

'*******************************************************************************
'        保存確認メッセージを表示せずにブック保存・クローズ
'*******************************************************************************
Sub SaveClose()

'   SaveChanges に Falseを代入するとブックは保存されずに閉じる
    ActiveWorkbook.Close SaveChanges:=True
    
End Sub

'*******************************************************************************
'        開いたブックと同じフォルダにある別のブック(DataBook.xls)を開く
'*******************************************************************************
Private Sub Workbook_Open()

    Application.ScreenUpdating = False

'   カレントドライブを変更
    ChDrive ActiveWorkbook.Path
'   カレントフォルダを変更
    ChDir ActiveWorkbook.Path
    
    Workbooks.Open Filename:="DataBook.xls"

'   カレントドライブ，フォルダを変更したくないとき
'   Pathプロパティが返す文字列を Openメソッドの引数に指定する
    Dim myPath  As String
    
    myPath = ActiveWorkbook.Path
    Workbooks.Open Filename:=myPath & "\DataBook.xls"
    
'   一度も保存されていない新規ブックの場合には
'   Pathプロパティは空の文字列(Null値)を返す
    
End Sub

'*******************************************************************************
'        他のユーザ使用中ブック  強制的クローズ
'*******************************************************************************
Sub CloseReadOnlyBook()

'   Excelは 他のユーザが使用中のブックを開くと そのブックは自動的に「読取専用」となる
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Workbooks.Open "日報.xls"

    '読み取り専用だったらそのブックを閉じる
    If ActiveWorkbook.ReadOnly = True Then
        ActiveWorkbook.Close
        MsgBox "日報.xlsは他のユーザーが使用中です"
    End If
    
'   ブックが始めから読み取り専用に指定されていたら 他のユーザーが使用していなくても
'   開くことはできない
'   しかし VBAの仕様を考えると この方法が最も効率が良く また十分実用に耐え得るマクロと言える
    
End Sub

'*******************************************************************************
'        隠しシート
'*******************************************************************************
'        処理概要：シートをユーザーが表示できないように隠す
'*******************************************************************************
Sub HiddenSheet()

'   xlVeryHiddenにより
'   Excelの［書式(O)］→［シート(H)］→［再表示(U)...］コマンドが無効となる
    Worksheets("Sheet3").Visible = xlVeryHidden
    
'   xlVeryHiddenによって非表示になったシートは
'   VisibleプロパティにTrueを代入すれば再表示される
    
End Sub

'*******************************************************************************
'        ブック非表示
'*******************************************************************************
Sub HiddenBook()

'   ブックを非表示にする
'   WorkbookオブジェクトにはVisibleプロパティがないため
'   ブックのすべてのWindowオブジェクトのVisibleプロパティにFalseを代入しなければならない
    Dim myWindow  As Window

    For Each myWindow In ActiveWorkbook.Windows
        myWindow.Visible = False
    Next myWindow

'   ユーザーは［ウィンドウ(W)］－［再表示(U)...］コマンドで
'   ブックを再表示することができる

End Sub

'*******************************************************************************
'        マクロによる変更は受け入れるようにシートを保護する
'*******************************************************************************
Sub ProtectWSheet()

'   パスワード"pswd1961"でシートを保護する
    Worksheets("Sheet1").Protect _
        Password:="pswd1961", _
        UserInterfaceOnly:=True

'   セルの内容をマクロから変更する
    Range("A1:C10").Value = Array("ABC", "DEF", "HIJ")
    
'   UserInterfaceOnly に False を指定してProtectメソッドを実行したワークシートの場合
'  （UserInterfaceOnlyを省略した場合も同様である）
'   ユーザーはロックされたセルの内容を手動で変更することはできない
'   同時に マクロでセルの内容を変更することもできなくなる
    
'   シートの保護を解除する
    Worksheets("Sheet1").Unprotect Password:="pswd1961"
    
End Sub

'*******************************************************************************
'        入力範囲制限
'*******************************************************************************
Sub LimitArea()

    With Worksheets("納品書")
        'スクロール範囲に名前を定義
        .Range("A1:I14").Name = "入力範囲"
        
        'スクロール範囲を制限
        'セル範囲A1:I14以外のセルを選択したり
        'セル範囲A1:I14が隠れてしまうような画面スクロールは不可能になる
        .ScrollArea = "入力範囲"
        
'   全セルを選択できるように設定を戻すときには ScrollAreaプロパティに空の文字列を代入
        
        'ロックが解除されたセルのみ入力可能とする
        .EnableSelection = xlUnlockedCells
        
        'シートを保護する
        .Protect Contents:=True, UserInterfaceOnly:=True
    End With
    
End Sub

'*******************************************************************************
'        アクティブウィンドウ以外のウィンドウを最小化する
'*******************************************************************************
Sub WindowIcon()

    Dim myWindow  As Window
    Dim myWndName As String
    
    'アクティブウィンドウの名前を取得 (Nameプロパティではないことに注意)
    myWndName = ActiveWindow.Caption
    
    For Each myWindow In Windows
        'アクティブウィンドウでなかったら最小化する
        If myWindow.Caption <> myWndName Then
            myWindow.WindowState = xlMinimized
        End If
        
    Next myWindow
    
End Sub

'*******************************************************************************
'        Excel ウィンドウ最小化・表示
'*******************************************************************************
Sub WindowSize()

'   Excelのウィンドウを最小化する
    Application.WindowState = xlMinimized
    
'   Excelのウィンドウを表示する
    Application.WindowState = xlNormal

End Sub

'*******************************************************************************
'        Sort処理Test
'*******************************************************************************
Sub SortTest()

    Dim rngSort   As Range
    Dim strColumn As String
    
    Workbooks("Book1").Worksheets("Sheet1").Activate
    
    strColumn = "A1"
    Set rngSort = Range(strColumn).CurrentRegion
    
    'rngSort.Select
    'Worksheets(strSheetName).Range(strColumnRange).sort
    rngSort.sort _
                             Key1:=Range(strColumn), _
                             Order1:=xlAscending, _
                             Header:=xlGuess, _
                             OrderCustom:=1, _
                             MatchCase:=False, _
                             Orientation:=xlTopToBottom, _
                             SortMethod:=xlPinYin, _
                             DataOption1:=xlSortNormal

End Sub

