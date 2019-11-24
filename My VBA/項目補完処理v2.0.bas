Option Explicit

Dim isERR As Integer                                       ' ERROR フラグ

Public Type typCellPos                                     ' セルポジション型
    lngRow As Long
    lngCol As Long
End Type

Const conStartStr   As String = "№"                       ' 一覧表の開始文字列
Dim conItm          As Variant                             ' ヘッダー行の列配列
                                                           ' ヘッダー行の列名称
Const conHeadItm    As String = "№,○,大　分　類,中　分　類,概　　　　 要,概要（検索用）,出 力 日 付,"

Const conArrayMax   As Integer = 10                        ' ヘッダー行のコピー補完対象列 MAX数
Const conDiffItm    As Integer = 2                         ' ヘッダー行と明細行の差(行数)

'*******************************************************************************
'        項 目 補 完 処 理
'*******************************************************************************
'        処理概要：★ＥｘｃｅｌＶＢＡまとめ★.xlsm において、コピー補完対象列の上罫線が存在する
'                  明細行をコピーし、下罫線が存在する行までペーストして補完する
'
'                  コピー補完対象列："№", "○", "大分類", "中分類", "概要", "出力日付"
'
'                  なお、№ はシリアルに自動符番を行う
'*******************************************************************************
Public Sub 項目補完処理()

    Dim posStart            As typCellPos                  ' 処理対象のSTARTセル位置 (行はSTART行，列は"is"列固定)
    Dim posEnd              As typCellPos                  ' 処理対象のENDセル位置   (行はEND行，列は"is"列固定)
    Dim posItm(conArrayMax) As typCellPos                  ' コピー補完対象列のセル位置 配列
                                                           ' (行は明細行固定，列はコピー補完対象列)
    Dim intNo               As Integer                     ' "№"列のシリアルNO
    
    Dim i                   As Integer
    
    Application.ScreenUpdating = False                     ' 画面更新 停止
    Call ProcInit(posStart, posEnd, posItm)                ' 初期処理
    
    intNo = 0
    Do While (posItm(0).lngRow <= posEnd.lngRow)           ' コピー補完対象列の行 ≦ 処理対象END行 の間 処理を行う
    
       Call ProcItmPrep(posItm)                            ' 概要(検索用) １行目を作成
       intNo = intNo + 1                                   ' №シリアル値 インクリメント
       Call ProcItmCopy(intNo, posStart, posItm)           ' 項目補完処理
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
                     
    Dim rngFind   As Range                                 ' FIND関数のリターン値
    Dim intRtn    As Integer                               ' MSGBOX関数のリターン値
    
    Dim i         As Integer
    Dim j         As Integer
    
    conItm = Split(conHeadItm, ",")                        ' ヘッダー行の列名称を配列にする
    
    Set rngFind = Cells.Find(conStartStr)                  ' 開始文字列 検索
    If rngFind Is Nothing Then                             ' 開始文字列 がないとき
       intRtn = MsgBox(prompt:="開始文字列 が見つかりません", Buttons:=vbOKOnly + vbCritical)
       If intRtn = vbOK Then
          isERR = True
          Exit Sub
       End If
    End If
                                                           ' 処理対象のSTARTセル
    posStart.lngRow = rngFind.Row                          ' 行NO に開始文字列の行 設定
    posStart.lngCol = rngFind.Column                       ' 列NO に開始文字列の列 設定
                                                           ' 処理対象のEndセル位置
    posEnd.lngRow = Cells.SpecialCells(xlCellTypeLastCell).Row    ' 行NO に使用範囲内最終セルの行 設定
    posEnd.lngCol = Cells.SpecialCells(xlCellTypeLastCell).Column ' 行NO に使用範囲内最終セルの行 設定
    
    i = 0
    j = 0
    Do While conItm(i) <> ""                               ' ヘッダー行の列名称配列の全データ分処理を行う
    
       Set rngFind = Cells.Find(conItm(i))                 ' ヘッダー行の列名称配列をFINDする
       If rngFind Is Nothing Then                          ' ヘッダー行の列名称がないとき
       Else                                                ' ヘッダー行の列名称があるとき
          posItm(j).lngRow = rngFind.Offset(conDiffItm).Row       ' 行NO にヘッダー行の列名称の明細行 設定
          posItm(j).lngCol = rngFind.Column                ' 列NO にヘッダー行の列名称の列 設定
          
          j = j + 1
       End If
       
       i = i + 1
    Loop
    
End Sub

'*******************************************************************************
'        項 目 補 完 準 備 処 理
'*******************************************************************************
'        処理概要："概要"列に複数行設定されているとき、それらをすべて結合し
'                  概要(検索用)の１行目に設定する
'
'            戻り値　：なし
'            引数１　：項目ポジション   typCellPos
'*******************************************************************************
Private Sub ProcItmPrep(posItm() As typCellPos)

    Dim intRow  As Integer
    Dim rngCell As Range
    Dim strComb As String
    Dim lngRGB  As Long

    strComb = ""
    intRow = 1                                             ' "概要"列の上罫線設定行の網掛色を保存
    lngRGB = Cells(posItm(4).lngRow, posItm(4).lngCol).Interior.Color
                                                           ' "概要"列の次の下罫線が出てくる差分行数を算出
    Do While (Cells(posItm(4).lngRow, posItm(4).lngCol).Offset(intRow).Borders(xlEdgeTop).LineStyle = xlLineStyleNone)
                                                           ' "概要"列のコピー補完行を初期化
       If Cells(posItm(4).lngRow, posItm(4).lngCol).Offset(intRow).Font.Color = lngRGB Then
          Cells(posItm(4).lngRow, posItm(4).lngCol).Offset(intRow).Value = ""
       End If
       
       intRow = intRow + 1
    Loop
                                                           ' "概要"列上罫線～下罫線設定行のすべての文字列を結合
    For Each rngCell In Range(Cells(posItm(4).lngRow, posItm(4).lngCol), _
                              Cells(posItm(4).lngRow, posItm(4).lngCol).Offset(intRow - 1))
       
       strComb = strComb & rngCell.Value
    Next
                                                           '"概要(検索用)"列 の上罫線設定行に "概要"列の結合文字を設定
    Cells(posItm(5).lngRow, posItm(5).lngCol).Value = strComb  

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
                        posStart As typCellPos, _
                        posItm() As typCellPos)

    Dim i       As Integer
    Dim intRow  As Integer
    Dim lngRGB  As Long
    Dim rngCell As Range

    If Cells(posItm(2).lngRow, posItm(2).lngCol) <> "" Then    ' 大分類列の明細行の値≠ブランクのとき
       Cells(posItm(0).lngRow, posItm(0).lngCol) = intNo       ' "№"列の明細行の値にシリアルNO 設定
    End If
    
    i = 0                                                  ' コピー補完対象列の全データ分処理を行う
    Do While (i <= conArrayMax And _
              posItm(i).lngCol <> 0)
              
       Cells(posItm(i).lngRow, posItm(i).lngCol).Copy      ' コピー補完対象列の上罫線が存在する明細行をコピー
       lngRGB = Cells(posItm(i).lngRow, posItm(i).lngCol).Interior.Color
       posItm(i).lngRow = posItm(i).lngRow + 1
                                                           ' コピー補完対象列の次の上罫線が出てくるまで、処理を行う
       Do While (Cells(posItm(i).lngRow, posItm(i).lngCol).Borders(xlEdgeTop).LineStyle = xlLineStyleNone)

          If Cells(posItm(i).lngRow, posItm(i).lngCol).Value = "" Then
             Cells(posItm(i).lngRow, posItm(i).lngCol).PasteSpecial (xlPasteAllExceptBorders) ' ペースト(罫線を除く全て)
             Cells(posItm(i).lngRow, posItm(i).lngCol).Font.Color = lngRGB                    ' 文字色を WHITE にする
          Else
             Cells(posItm(i).lngRow, posItm(i).lngCol).Copy
          End If
                                                           ' 日付列のとき、曜日列も文字色も WHITE にする
          If Cells(posStart.lngCol, posItm(i).lngCol).Value = "出 力 日 付" Then
             Cells(posItm(i).lngRow, posItm(i).lngCol).Offset(, 1).Font.Color = lngRGB
          End If
           
          posItm(i).lngRow = posItm(i).lngRow + 1
       Loop
          
       i = i + 1
    Loop

End Sub
