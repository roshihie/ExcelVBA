Option Explicit
'*******************************************************************************
'        エクセルシート書式初期化
'*******************************************************************************
'      < 処理概要 >
'        シートの書式を初期化する。
'*******************************************************************************
Public Sub シート書式初期化()

  Dim nSize  As Long
  
  nSize = Application.InputBox(Prompt:="フォントサイズを入力する", _
                               Title:="ＭＳ 明朝 フォントサイズ指定", _
                               Default:=10, _
                               Type:=1)
  If nSize = 0 Then
    Exit Sub
  End If
  
  Application.ScreenUpdating = False

  With Cells
    .Font.Name = "ＭＳ 明朝"
    .Font.Size = nSize
    .ColumnWidth = 2
    .VerticalAlignment = xlCenter
  End With
  
  Columns(1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
  Columns(1).ColumnWidth = 0.3
  Rows(1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
  Rows(1).RowHeight = 3
  
  With ActiveSheet.PageSetup
    .LeftHeader = "&""ＭＳ 明朝,太字""&12&U &A "
    .CenterHeader = "&""ＭＳ 明朝,太字""&12 "
    .CenterFooter = "&""ＭＳ 明朝,標準""― &P／&N ―"
    .HeaderMargin = Application.CentimetersToPoints(0.9)
    .TopMargin = Application.CentimetersToPoints(1.9)
    .LeftMargin = Application.CentimetersToPoints(1.2)
    .RightMargin = Application.CentimetersToPoints(0.5)
    .BottomMargin = Application.CentimetersToPoints(1.2)
    .FooterMargin = Application.CentimetersToPoints(0.7)
  End With
  
  Application.ScreenUpdating = True
  Cells(1, 1).Activate

End Sub
