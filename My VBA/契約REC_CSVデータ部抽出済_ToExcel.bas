Option Explicit
'*******************************************************************************
'        CLM0277 ファイル内容  Excel 変換
'*******************************************************************************
'      < 処理概要 >
'        EASYPLUS で出力したファイル内容(CSVデータ部抽出済)を Excel に変換する。
'*******************************************************************************
Public Sub 契約REC_CSVデータ部抽出済_ToExcel()

  Const sBOOK_PASS As String = _
        "C:\Users\roshi_000\MyImp\MyOwn\Develop\Excel VBA\My VBA\データ\"
  Const sHDR_BOOK  As String = "CLM0277_契約REC_HDR.xlsx"
  Const sHDR_SHEET As String = "契約"
  Dim oFoundRng    As Range

  Application.ScreenUpdating = False
  Columns(2).TextToColumns Destination:=Cells(1, 2), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, Comma:=True, _
    FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), _
                     Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 2), Array(10, 2), _
                     Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15, 2), _
                     Array(16, 2), Array(17, 2), Array(18, 2), Array(19, 2), Array(20, 2), _
                     Array(21, 2), Array(22, 2), Array(23, 2), Array(24, 2), Array(25, 2), _
                     Array(26, 2)), _
    TrailingMinusNumbers:=True
    
'                              数字項目のカンマ編集
' With Range(Cells(2, 6).Offset(, 1), Cells(2, 7).Offset(, 1).End(xlDown))
'   .NumberFormatLocal = "#,##0_ "
' End With

  Call sb共通_ToExcel(sBOOK_PASS, sHDR_BOOK, sHDR_SHEET)

  Set oFoundRng = Cells.Find(What:=sHDR_SHEET, LookAt:=xlPart)
  If oFoundRng Is Nothing Then
    Cells(1, 1).Activate
  Else
    oFoundRng.End(xlDown).Offset(, 1).Select
    ActiveWindow.FreezePanes = True
  End If

End Sub
    
'*******************************************************************************
'        Excel 変換共通処理
'*******************************************************************************
'      < 処理概要 >
'        Excel 変換において、共通処理をまとめる
'*******************************************************************************
Private Sub sb共通_ToExcel(sBOOK_PASS As String, sHDR_BOOK As String, sHDR_SHEET As String)

  Dim sTrgBook As String
  Dim nRow     As Long
  Dim nCol     As Long
  Dim oTrgRect As tRect
  Dim oHDRRect As tRect
  Dim oHDRStr  As tHDRStr
  
  Set oTrgRect.oBgn = Cells(2, 2)
  nRow = Cells(ActiveSheet.Rows.Count, 2).End(xlUp).Row
  nCol = Cells(nRow, ActiveSheet.Columns.Count).End(xlToLeft).Column
  Set oTrgRect.oEnd = Cells(nRow, nCol)
  
  With Range(oTrgRect.oBgn, oTrgRect.oEnd)
    .RowHeight = 12
    .HorizontalAlignment = xlCenter
  End With
'                              ３列挿入 (｢表示｣, ｢モジュール｣, ｢テストケース番号｣列)
  Range(Columns(oTrgRect.oBgn.Column), Columns(oTrgRect.oBgn.Column).Offset(, 2)).Insert Shift:=xlToRight, _
    CopyOrigin:=xlFormatFromRightOrBelow
    
  Set oTrgRect.oBgn = oTrgRect.oBgn.Offset(, -3)
  Set oTrgRect.oEnd = Cells(oTrgRect.oEnd.Row, _
                            Cells(oTrgRect.oEnd.Row, ActiveSheet.Columns.Count).End(xlToLeft).Column)
  Range(oTrgRect.oBgn, oTrgRect.oEnd).Select
  Call sb罫線描画(Selection)

  sTrgBook = ActiveWorkbook.Name
  Workbooks.Open Filename:=sBOOK_PASS & sHDR_BOOK, ReadOnly:=True
  Worksheets(sHDR_SHEET).Activate

  Set oHDRRect.oBgn = Cells(2, 2)
  nRow = Cells(ActiveSheet.Rows.Count, 2).End(xlUp).Offset(1).Row
  nCol = Cells(nRow - 1, ActiveSheet.Columns.Count).End(xlToLeft).Column
  Set oHDRRect.oEnd = Cells(nRow, nCol)
  Range(Rows(oHDRRect.oBgn.Row), Rows(oHDRRect.oEnd.Row)).Copy
  With ActiveSheet.PageSetup
    oHDRStr.sLeftHDR = .LeftHeader
    oHDRStr.sCentHDR = .CenterHeader
    oHDRStr.sRigtHdr = .RightHeader
  End With

  Windows(sTrgBook).Activate
  With Rows(2)
    .Insert Shift:=xlDown
    .PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone
  End With
  
  With ActiveSheet.PageSetup                       '印刷セットアップ
    .LeftHeader = oHDRStr.sLeftHDR
    .CenterHeader = oHDRStr.sCentHDR
    .RightHeader = oHDRStr.sRigtHdr
    .Orientation = xlLandscape                       '横向き
    .Zoom = 75                                       '拡大縮小率
    .CenterHorizontally = False                      'ページ中央 水平
    .CenterVertically = False                        'ページ中央 垂直
  End With
  
  Application.DisplayAlerts = False
  Workbooks(sHDR_BOOK).Close
  Application.DisplayAlerts = True
  
  Range(oTrgRect.oBgn.Offset(-1), Cells(oTrgRect.oBgn.Offset(-1).Row, oTrgRect.oEnd.Column)).AutoFilter

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

  With oDrawRng.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  With oDrawRng.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  With oDrawRng.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  With oDrawRng.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  With oDrawRng.Borders(xlInsideHorizontal)
    .LineStyle = xlDot
    .Weight = xlThin
  End With
  With oDrawRng.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With

End Sub

