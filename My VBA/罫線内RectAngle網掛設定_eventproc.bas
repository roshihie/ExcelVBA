Option Explicit

Dim isERR As Integer                                       ' ERROR フラグ

Public Type typCellPos                                     ' セルポジション型
    lngRow As Long
    lngCol As Long
End Type

Dim conItm          As Variant                             ' ヘッダー行の列(明細行文字 網掛色変更列)配列
                                                           ' ヘッダー行の列(明細行文字 網掛色変更列)名称
Const conHeadItm    As String = "№,○,大　分　類,中　分　類,概　　　 要,発 生 日 付,対 処 日 付,"

Const conArrayMax   As Integer = 10                        ' ヘッダー行の列 MAX数
Const conDiffItm    As Integer = 2                         ' ヘッダー行と明細行の差(行数)
Const conRGBRed     As Integer = 191                       '網掛色(RGB Red)
Const conRGBGreen   As Integer = 191                       '網掛色(RGB Green)
Const conRGBBlue    As Integer = 191                       '網掛色(RGB Blue)

'*******************************************************************************
'        罫線内ＲｅｃｔＡｎｇｌｅ 網掛設定
'*******************************************************************************
'        処理概要：一覧表の ActiveCell行 を含む 直近の上罫線，下罫線で囲まれた
'                  RectAngle を網掛する
'                  このとき、ヘッダー行の指定列の明細行(明細行文字 網掛色設定列)に対して
'                  上罫線が存在しない行の文字は 網掛色に設定する
'
'*******************************************************************************
Public Sub 罫線内RectAngle網掛設定()

    Dim posStart            As typCellPos                  ' 処理対象のSTARTセル位置 (行はSTART行，列は"is"行固定)
    Dim posEnd              As typCellPos                  ' 処理対象のENDセル位置   (行はEND行，列は"is"行固定)
    Dim posItm(conArrayMax) As typCellPos                  ' 明細行文字 網掛色設定列 配列
                                                           ' (行はSTART行，列は文字色設定列)
    
    Dim i                   As Integer
    
    Application.ScreenUpdating = False                     ' 画面更新 停止
    Call ProcInit(posStart, posEnd, posItm)                ' 初期処理
    
    Call ProcInteriorSet(posStart, posEnd)                 ' 網掛設定
    
    Do While (posItm(0).lngRow <= posEnd.lngRow)           ' 明細行文字 網掛色設定行 ≦ 処理対象END行 の間 処理を行う
    
       Call ProcFontSet(posEnd, posItm)                    ' 明細行文字 網掛色設定
    Loop

End Sub

'*******************************************************************************
'        初　期　処　理
'*******************************************************************************
'        処理概要：処理対象のSTART行，END行を確定し、明細行文字 網掛色設定列を配列に格納する
'
'            戻り値　：なし
'            引数１　：開始ポジション   typCellPos
'            引数２　：終了ポジション   typCellPos
'            引数３  ：項目ポジション   typCellPos
'*******************************************************************************
Private Sub ProcInit(posStart As typCellPos, _
                     posEnd As typCellPos, _
                     posItm() As typCellPos)
                     
    Dim rngFind   As Range                                 ' FIND関数のリターン値
    Dim lngCol    As Long                                  ' ワーク列NO
    Dim intRtn    As Integer                               ' MSGBOX関数のリターン値
    
    Dim i, j      As Integer
    
    conItm = Split(conHeadItm, ",")                        ' ヘッダー行の列名称を配列にする
    
    Set rngFind = Cells.Find("is")                         ' is列(全行 "1" 埋め込み列) 検索
    If rngFind Is Nothing Then                             ' is列 がないとき
       intRtn = MsgBox(prompt:="is列(全行 ""1"" 埋め込み列) が見つかりません", Buttons:=vbOKOnly + vbCritical)
       If intRtn = vbOK Then
          isERR = True
          Exit Sub
       End If
    End If
    
    posStart.lngCol = Cells(rngFind.Row, 1).End(xlToRight).Column   ' is列の最左列NO を取得
    posStart.lngRow = ActiveCell.Row                                ' ActiveCell行の直近の上罫線行 取得
    Do While (Cells(posStart.lngRow, posStart.lngCol).Borders(xlEdgeTop).LineStyle = xlLineStyleNone)
    
       posStart.lngRow = posStart.lngRow - 1
    Loop
                                                           ' is列の最右列NO を取得
    posEnd.lngCol = Cells(rngFind.Row, ActiveSheet.Columns.Count).End(xlToLeft).Column
    posEnd.lngRow = ActiveCell.Row                                  ' ActiveCell行の直近の下罫線行 取得
    Do While (Cells(posEnd.lngRow, posEnd.lngCol).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone)
    
       posEnd.lngRow = posEnd.lngRow + 1
    Loop
    
    i = 0
    Do While conItm(i) <> ""                               ' ヘッダー行の列名称配列の全データ分処理を行う
    
       Set rngFind = Cells.Find(conItm(i))                 ' ヘッダー行の列名称配列をFINDする
       If rngFind Is Nothing Then                          ' ヘッダー行の列名称がないとき
          posItm(i).lngRow = 0
          posItm(i).lngCol = 0
       Else                                                ' ヘッダー行の列名称があるとき
          posItm(i).lngCol = rngFind.Column                ' 列NO にヘッダー行の列名称の列 設定
          posItm(i).lngRow = posStart.lngRow               ' 行NO にヘッダー行の列名称の明細行 設定
       End If
       i = i + 1
    Loop
    
End Sub

'*******************************************************************************
'        罫線内ＲｅｃｔＡｎｇｌｅ 網掛設定
'*******************************************************************************
'        処理概要：ActiveCell行 を含む 直近の罫線で囲まれた RectAngle の網掛を行う
'                  網掛カラー：RGB(191, 191, 191)
'
'            戻り値　：なし
'            引数１　：開始ポジション   typCellPos
'            引数２　：終了ポジション   typCellPos
'*******************************************************************************
Private Sub ProcInteriorSet(posStart As typCellPos, _
                            posEnd As typCellPos)
    
    Range(Cells(posStart.lngRow, posStart.lngCol), _
          Cells(posEnd.lngRow, posEnd.lngCol)).Interior.Color = RGB(conRGBRed, conRGBGreen, conRGBBlue)

End Sub

'*******************************************************************************
'        明細行文字 網掛色設定処理
'*******************************************************************************
'        処理概要：ヘッダー行の指定列の明細行(明細行文字 網掛色設定列)に対して
'                  上罫線が存在しない行の文字は 網掛色に設定する
'
'            戻り値　：なし
'            引数１　：終了ポジション   typCellPos
'            引数２　：項目ポジション   typCellPos
'*******************************************************************************
Private Sub ProcFontSet(posEnd As typCellPos, _
                        posItm() As typCellPos)

    Dim i    As Integer
    
    i = 0                                                  ' (明細行文字 網掛色設定列の全データ分処理を行う
    Do While (i <= conArrayMax And _
              posItm(i).lngCol <> 0)
              
       posItm(i).lngRow = posItm(i).lngRow + 1             ' 上罫線が存在する明細行は処理なし
                                                           ' 文字色設定列の明細行≦終了ポジションの行 の間、処理を行う
       Do While (posItm(i).lngRow <= posEnd.lngRow)
                                                           ' 文字色を 網掛色に設定する
          Cells(posItm(i).lngRow, posItm(i).lngCol).Font.Color = RGB(conRGBRed, conRGBGreen, conRGBBlue)
                                                           ' 日付列のとき、曜日列も文字色も 網掛色に設定する
          If (Cells(2, posItm(i).lngCol).Value = "発 生 日 付" Or _
              Cells(2, posItm(i).lngCol).Value = "対 処 日 付") Then
              Cells(posItm(i).lngRow, posItm(i).lngCol).Offset(, 1).Font.Color = RGB(conRGBRed, conRGBGreen, conRGBBlue)
          End If
           
          posItm(i).lngRow = posItm(i).lngRow + 1
       Loop
          
       i = i + 1
    Loop

End Sub
