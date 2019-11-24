Option Explicit
'*******************************************************************************
'        契約内容一覧  制度内容一覧 連結処理
'*******************************************************************************
'      < 処理概要 >
'        基となる内容一覧に別な内容一覧をキー指定で連結する。
'*******************************************************************************
Public Sub 契約内容_制度内容_連結()
  
  Const s制度番号 As String = "制度番号"
  
  Dim oTrgRect  As tRect
  Dim oDataTop  As tRect
  Dim oFoundRng As Range
  Dim nProc     As Long
  Dim o制度番号Rngs As Range
  Dim o制度番号Rng  As Range

  Application.ScreenUpdating = False
  
  Set oFoundRng = Cells.Find(What:=s制度番号, LookAt:=xlPart)
  If oFoundRng Is Nothing Then
     Exit Sub
  End If
  
  nProc = 0
  Set oTrgRect.oBgn = Cells(oFoundRng.Row, ActiveSheet.Columns.Count).End(xlToLeft).Offset(, 1)
  Call sb制度データ取得(nProc, oFoundRng, oTrgRect)
  Set oTrgRect.oEnd = Cells(oTrgRect.oBgn.Row, ActiveSheet.Columns.Count).End(xlToLeft).Offset(1)
  
  nProc = 1
  Set o制度番号Rngs = Range(oFoundRng.End(xlDown), Cells(ActiveSheet.Rows.Count, oFoundRng.Column).End(xlUp))
  Set oDataTop.oBgn = Cells(oFoundRng.End(xlDown).Row, oTrgRect.oBgn.Column)
  Set oDataTop.oEnd = Cells(oFoundRng.End(xlDown).Row, oTrgRect.oEnd.Column)
  
  For Each o制度番号Rng In o制度番号Rngs
  
    Set oTrgRect.oBgn = Cells(o制度番号Rng.Row, oTrgRect.oBgn.Column)
    Set oTrgRect.oEnd = Cells(o制度番号Rng.Row, oTrgRect.oEnd.Column)
    
    If o制度番号Rng.Value <> o制度番号Rng.Offset(-1).Value Then
      Call sb制度データ取得(nProc, o制度番号Rng, oTrgRect)
      
      If o制度番号Rng.Offset(-1).Value <> "" Then
        Range(oDataTop.oBgn, oTrgRect.oEnd.Offset(-1)).Select
        Call sb罫線描画(Selection)
      End If
    End If
    
  Next

  Range(oDataTop.oBgn, oTrgRect.oEnd).Select
  Call sb罫線描画(Selection)
  With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  
  nProc = 2
  Call sb制度データ取得(nProc, oFoundRng, oTrgRect)
  
  ActiveSheet.AutoFilterMode = False
  Range(Cells(oDataTop.oBgn.Offset(-1).Row, 2), oDataTop.oEnd.Offset(-1)).AutoFilter
  Application.ScreenUpdating = True
  Cells(1, 1).Activate
     
End Sub

'*******************************************************************************
'        制度データ取得処理
'*******************************************************************************
'      < 処理概要 >
'        処理区分に応じて、制度内容一覧のヘッダー部やデータ部を取得する。
'        処理区分：0  ヘッダー部を取得
'        処理区分：1  データ部を取得
'*******************************************************************************
Private Sub sb制度データ取得(nProc As Long, o制度番号Rng As Range, oTrgRect As tRect)

  Const s制度BOOK_PASS As String = _
        "C:\Users\roshi_000\MyImp\MyOwn\Develop\Excel VBA\My VBA\データ\"
  ' 上記の " がエスケープシーケンスになり gVimで開くと以降がコメントになるため当行追加
  Const s制度BOOK  As String = "制度内容一覧.xlsx"
  Const s制度SHEET As String = "制度"

  Static stsTrgBook As String
  Static stn制度Col As Long
  Static stoSocRect As tRect
  Dim oFoundRng As Range
  Dim nRow As Long
  Dim nCol As Long
  
  Select Case True
  Case nProc = 0               '初期処理
    stsTrgBook = ActiveWorkbook.Name
    Workbooks.Open Filename:=s制度BOOK_PASS & s制度BOOK, ReadOnly:=True
    Worksheets(s制度SHEET).Activate
    Set oFoundRng = Cells.Find(What:=o制度番号Rng.Value, LookAt:=xlPart)
    If oFoundRng Is Nothing Then
      Exit Sub
    End If
    
    stn制度Col = oFoundRng.Column
    Set stoSocRect.oBgn = oFoundRng.Offset(, 1)
    nRow = oFoundRng.End(xlDown).Offset(-1).Row
    nCol = Cells(stoSocRect.oBgn.Row, ActiveSheet.Columns.Count).End(xlToLeft).Column
    Set stoSocRect.oEnd = Cells(nRow, nCol)
  
    Range(stoSocRect.oBgn, stoSocRect.oEnd).Copy
    
    Windows(stsTrgBook).Activate
        
    With oTrgRect.oBgn
      .PasteSpecial Paste:=xlPasteColumnWidths
      .PasteSpecial Paste:=xlPasteAll
    End With
    Application.CutCopyMode = False
    
  Case nProc = 1               'データ処理
    Windows(s制度BOOK).Activate
    Set oFoundRng = Columns(stn制度Col).Find(What:=o制度番号Rng.Value, LookAt:=xlWhole)
    If oFoundRng Is Nothing Then
      Exit Sub
    End If
    
    Set stoSocRect.oBgn = Cells(oFoundRng.Row, stoSocRect.oBgn.Column)
    Set stoSocRect.oEnd = Cells(oFoundRng.Row, stoSocRect.oEnd.Column)
  
    Range(stoSocRect.oBgn, stoSocRect.oEnd).Copy
    
    Windows(stsTrgBook).Activate
    With oTrgRect.oBgn
      .PasteSpecial Paste:=xlPasteColumnWidths
      .PasteSpecial Paste:=xlPasteAllExceptBorders
    End With
    Application.CutCopyMode = False
    
  Case nProc = 2               '終了処理
    Windows(s制度BOOK).Activate
    Application.DisplayAlerts = False
    Workbooks(s制度BOOK).Close
    Application.DisplayAlerts = True
    Windows(stsTrgBook).Activate
  End Select
  
End Sub

'*******************************************************************************
'        罫線描画
'*******************************************************************************
'      < 処理概要 >
'        指定されたエリアを以下の通り罫線を引く。
'        エリアの囲み線，縦線：実線，
'                横線        ：破線 で描画する｡
'*******************************************************************************
Private Sub sb罫線描画(oDrawRng As Range)

  With oDrawRng
    With .Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .Weight = xlThin
    End With
    With .Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .Weight = xlThin
    End With
    With .Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .Weight = xlThin
    End With
    With .Borders(xlEdgeBottom)
      .LineStyle = xlDot
      .Weight = xlThin
    End With
  End With

End Sub

