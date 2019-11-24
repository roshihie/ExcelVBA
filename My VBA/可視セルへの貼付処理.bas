Option Explicit
'*******************************************************************************
'        可視セル範囲への貼付処理
'*******************************************************************************
'      < 処理概要 >
'        クリップボードへコピーしたデータを可視セル範囲へ貼り付ける。
'        （フィルタリングした範囲への貼り付け）
'*******************************************************************************
Public Sub 可視セルへの貼付処理()

  Dim oClipb   As DataObject
  Dim oTrgRect As tRect
  Dim vClipbStr   As Variant
  Dim asCellStr() As String
  Dim bGetTopRng  As Boolean
  
  Dim nGap     As Long
  Dim oSocRng  As Range
  Dim oDstRng  As Range
  Dim oTopRng  As Range
  Dim oDst     As Range
  Dim i        As Long
  Dim j        As Long
  
  On Error GoTo InputCancel
  Set oSocRng = Application.InputBox(Prompt:="コピーするエリアを指定して下さい。", _
                                     Title:="コピーエリア指定", _
                                     Type:=8)
  Set oDstRng = ActiveCell.Worksheet.AutoFilter.Range
  Set oDstRng = oDstRng.Resize(, oDstRng.Columns.Count + 1)
  Set oDstRng = Intersect(oDstRng, oDstRng.Offset(1))
  
  bGetTopRng = True
  Do While (bGetTopRng = True)
    Set oTopRng = Application.InputBox(Prompt:="貼り付ける可視エリアの先頭セルを指定して下さい。", _
                                       Title:="貼り付けエリア先頭指定", _
                                       Type:=8)
    If Intersect(oDstRng, oTopRng) Is Nothing Then
      MsgBox ("貼り付ける可視エリアの先頭セルを正しく指定して下さい。")
    Else
      bGetTopRng = False
    End If
  Loop
                                      
  nGap = oDstRng.Columns(1).Column - 1           'オートフィルタ範囲の先頭列の差異
  Set oTrgRect.oBgn = oTopRng                    ' 貼付先頭セルの取得
                                                 
  With oDstRng.Columns(oTopRng.Column - nGap)    ' オートフィルタ範囲内での貼付先頭セルと同列の最終セルの取得
    Set oTrgRect.oEnd = .Cells(.Cells.Count)
  End With
                                                 ' オートフィルタ範囲を貼付先頭セル～最終列(同列)の範囲に限定
  Set oDstRng = Intersect(oDstRng, Range(oTrgRect.oBgn, oTrgRect.oEnd))
                                                 ' 可視セルを取得
  Set oDstRng = oDstRng.SpecialCells(xlCellTypeVisible)
  
  oSocRng.Copy
  
  Set oClipb = New DataObject
  With oClipb
    .GetFromClipboard
    On Error Resume Next
    vClipbStr = .GetText
    On Error GoTo 0
  End With

  If Not IsEmpty(vClipbStr) Then
    vClipbStr = Split(CStr(vClipbStr), vbCrLf)
    i = 0
    For Each oDst In oDstRng.Cells
      If i > UBound(vClipbStr) Then Exit For
      asCellStr = Split(vClipbStr(i), vbTab)
      For j = 0 To UBound(asCellStr)
        oDst.Offset(, j).Value = asCellStr(j)
      Next
      
      i = i + 1
    Next
  End If
  
  Set oClipb = Nothing

InputCancel:
End Sub
