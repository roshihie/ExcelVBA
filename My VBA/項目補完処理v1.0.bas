Option Explicit

Dim isERR As Integer                                                ' ERROR フラグ

Public Type typCellPos                                              ' セルポジション型
    lngRow As Long
    lngCol As Long
End Type

Dim conItm          As Variant                                      ' ヘッダー行の列配列
                                                                    ' ヘッダー行の列名称
Const conHeadItm    As String = "№,○,大　分　類,中　分　類,概　　　 要,発 生 日 付,対 処 日 付,"

Const conArrayMax   As Integer = 10                                 ' ヘッダー行のコピー補完対象列 MAX数
Const conDiffItm    As Integer = 2                                  ' ヘッダー行と明細行の差(行数)
Const conRGBMax     As Integer = 255                                '網掛け色(RGB White用)

'*******************************************************************************
'        項 目 補 完 処 理
'*******************************************************************************
'        処理概要：★簿記２級学習ポイント★.xlsm において、コピー補完対象列の上罫線が存在する
'                  明細行をコピーし、下罫線が存在する行までペーストして補完する
'
'                  コピー補完対象列："№", "○", "大分類", "中分類", "概要", "発生日付", "対処日付"
'
'                  なお、№ はシリアルに自動符番を行う
'*******************************************************************************
Public Sub 項目補完処理()

    Dim posStart            As typCellPos                           ' 処理対象のSTARTセル位置 (行はSTART行，列は"is"行固定)
    Dim posEnd              As typCellPos                           ' 処理対象のENDセル位置   (行はEND行，列は"is"行固定)
    Dim posItm(conArrayMax) As typCellPos                           ' コピー補完対象列のセル位置 配列
                                                                    ' (行は明細行固定，列はコピー補完対象列)
    Dim intNo               As Integer                              ' "№"列のシリアルNO
    
    Dim i                   As Integer
    
    Application.ScreenUpdating = False                              ' 画面更新 停止
    Call ProcInit(posStart, posEnd, posItm)                         ' 初期処理
    
    intNo = 0
    Do While (posItm(0).lngRow <= posEnd.lngRow)                    ' コピー補完対象列の行 ≦ 処理対象END行 の間 処理を行う
    
       intNo = intNo + 1                                            ' №シリアル値 インクリメント
       Call ProcItmCopy(intNo, posItm)                              ' 項目補完処理
    Loop

End Sub

'*******************************************************************************
'        初　期　処　理
'*******************************************************************************
'        処理概要：処理対象のSTART行，END行を確定し、コピー補完対象列を配列に格納する
'
'            戻り値　：なし
'            引数１　：開始ポジション   typCellPos
'            引数２　：終了ポジション   typCellPos
'            引数３  ：項目ポジション   typCellPos
'*******************************************************************************
Private Sub ProcInit(posStart As typCellPos, _
                     posEnd As typCellPos, _
                     posItm() As typCellPos)
                     
    Dim rngFind   As Range                                          ' FIND関数のリターン値
    Dim lngCol    As Long                                           ' ワーク列NO
    Dim intRtn    As Integer                                        ' MSGBOX関数のリターン値
    
    Dim i, j      As Integer
    
    conItm = Split(conHeadItm, ",")                                 ' ヘッダー行の列名称を配列にする
    
    Set rngFind = Cells.Find("is")                                  ' is列(全行 "1" 埋め込み列) 検索
    If rngFind Is Nothing Then                                      ' is列 がないとき
       intRtn = MsgBox(prompt:="is列(全行 ""1"" 埋め込み列) が見つかりません", Buttons:=vbOKOnly + vbCritical)
       If intRtn = vbOK Then
          isERR = True
          Exit Sub
       End If
    Else                                                            ' is列 があるとき
       lngCol = rngFind.Column                                          ' is列 の列NOを設定
    End If
                                                                    ' 処理対象のENDセル位置
    posEnd.lngRow = Cells(ActiveSheet.Rows.Count, lngCol).End(xlUp).Row ' 行NO に一覧表明細部の最終行列 設定
    posEnd.lngCol = lngCol                                              ' 列NO に is列 設定
                                                                    ' 処理対象のSTARTセル位置
    posStart.lngRow = Cells(posEnd.lngRow, lngCol).End(xlUp).Row      ' 行NO に一覧表明細部の先頭行列 設定
    posStart.lngCol = lngCol                                          ' 列NO に is列 設定
    
    i = 0
    Do While conItm(i) <> ""                                        ' ヘッダー行の列名称配列の全データ分処理を行う
    
       Set rngFind = Cells.Find(conItm(i))                          ' ヘッダー行の列名称配列をFINDする
       If rngFind Is Nothing Then                                   ' ヘッダー行の列名称がないとき
          posItm(i).lngRow = 0
          posItm(i).lngCol = 0
       Else                                                         ' ヘッダー行の列名称があるとき
          posItm(i).lngRow = rngFind.Offset(conDiffItm).Row           ' 行NO にヘッダー行の列名称の明細行 設定
          posItm(i).lngCol = rngFind.Column                           ' 列NO にヘッダー行の列名称の列 設定
       End If
       i = i + 1
    Loop
    
End Sub

'*******************************************************************************
'        項 目 補 完 処 理
'*******************************************************************************
'        処理概要：コピー補完対象列それぞれに対して、上罫線のある行の値をコピーして
'                  下罫線のある行までペーストする
'
'            戻り値　：なし
'            引数１　：シリアルNO       Integer
'            引数２　：項目ポジション   typCellPos
'*******************************************************************************
Private Sub ProcItmCopy(intNo As Integer, _
                        posItm() As typCellPos)

    Dim i    As Integer

    If Cells(posItm(2).lngRow, posItm(2).lngCol) <> "" Then         ' 大分類列の明細行の値≠ブランクのとき
       Cells(posItm(0).lngRow, posItm(0).lngCol) = intNo              ' №列の明細行の値にシリアルNO 設定
    End If
    
    i = 0                                                           ' コピー補完対象列の全データ分処理を行う
    Do While (i <= conArrayMax And _
              posItm(i).lngCol <> 0)
              
       Cells(posItm(i).lngRow, posItm(i).lngCol).Copy               ' コピー補完対象列の上罫線が存在する明細行をコピー
       posItm(i).lngRow = posItm(i).lngRow + 1
                                                                    ' コピー補完対象列の次の上罫線が出てくるまで、処理を行う
       Do While (Cells(posItm(i).lngRow, posItm(i).lngCol).Borders(xlEdgeTop).LineStyle = xlLineStyleNone)

          Cells(posItm(i).lngRow, posItm(i).lngCol).PasteSpecial (xlPasteAllExceptBorders)            ' ペースト(罫線を除く全て)
          Cells(posItm(i).lngRow, posItm(i).lngCol).Font.Color = RGB(conRGBMax, conRGBMax, conRGBMax) ' 文字色を WHITE にする
                                                                    ' 日付列のとき、曜日列も文字色も WHITE にする
          If (Cells(2, posItm(i).lngCol).Value = "発 生 日 付" Or _
              Cells(2, posItm(i).lngCol).Value = "対 処 日 付") Then
              Cells(posItm(i).lngRow, posItm(i).lngCol).Offset(, 1).Font.Color = RGB(conRGBMax, conRGBMax, conRGBMax)
          End If
           
          posItm(i).lngRow = posItm(i).lngRow + 1
       Loop
          
       i = i + 1
    Loop

End Sub
