Option Explicit

Dim isERR As Integer                                                ' ERROR フラグ

Public Type typCellPos                                              ' セルポジション型
    lngRow As Long
    lngCol As Long
End Type

Dim conItm          As Variant                                      ' ヘッダー行の初期設定必要列 配列
Const conHeadItm    As String = "is,発 生 日 付,対 処 日 付,"       ' ヘッダー行の初期設定必要列 名称

Const conArrayMax   As Integer = 10                                 ' ヘッダー行の初期設定必要列 MAX数
Const conDiffItm    As Integer = 2                                  ' ヘッダー行と明細行の差(行数)
Const conRGBMax     As Integer = 255                                '網掛け色(RGB White用)

'*******************************************************************************
'        新 規 行 挿 入 処 理
'*******************************************************************************
'        処理概要：★簿記２級学習ポイント★.xlsm において、挿入した新規行に
'                  固定値 または Excel関数が設定されている列に 初期設定を行う
'
'                  固定値 設定列  ："is"
'                  Excel関数設定列："発生日付", "対処日付"
'
'*******************************************************************************
Public Sub 新規行挿入処理()

    Dim posItm(conArrayMax) As typCellPos                           ' 初期設定必要列のセル位置 配列
    Dim posHead             As typCellPos                           ' ヘッダー行のセル位置
                                                                    ' (行は ActiveCell行固定，列は初期設定必要列)
    Dim i                   As Integer
    
    Application.ScreenUpdating = False                              ' 画面更新 停止
    Call ProcInit(posHead, posItm)                                  ' 初期処理
    Call ProcNewLine(posHead, posItm)                               ' 項目補完処理

End Sub

'*******************************************************************************
'        初　期　処　理
'*******************************************************************************
'        処理概要：Excel関数設定列を配列に格納する
'
'            戻り値　：なし
'            引数１  ：ヘッダーポジション  typCellPos
'            引数２  ：項目ポジション      typCellPos
'*******************************************************************************
Private Sub ProcInit(posHead As typCellPos, _
                     posItm() As typCellPos)
                     
    Dim rngFind   As Range                                          ' FIND関数のリターン値
    Dim lngCol    As Long                                           ' ワーク列NO
    Dim intRtn    As Integer                                        ' MSGBOX関数のリターン値
    
    Dim i, j      As Integer
    
    conItm = Split(conHeadItm, ",")                                 ' ヘッダー行の列名称を配列にする
    
    i = 0
    Do While conItm(i) <> ""                                        ' ヘッダー行の列名称配列の全データ分処理を行う
    
       Set rngFind = Cells.Find(conItm(i))                          ' ヘッダー行の列名称配列をFINDする
       If rngFind Is Nothing Then                                   ' ヘッダー行の列名称がないとき
          posItm(i).lngRow = 0
          posItm(i).lngCol = 0
       Else                                                         ' ヘッダー行の列名称があるとき
          posHead.lngRow = rngFind.Row                                ' ヘッダーポジション.行NO にヘッダー行 設定
          posHead.lngCol = rngFind.Column                             ' ヘッダーポジション.列NO にヘッダー行の列名称の列 設定
          posItm(i).lngRow = ActiveCell.Row                           ' 項目ポジション.行NO にActiveCell行 設定
          posItm(i).lngCol = rngFind.Column                           ' 項目ポジション.列NO にヘッダー行の列名称の列 設定
       End If
       i = i + 1
    Loop
    
End Sub

'*******************************************************************************
'        新規行挿入処理
'*******************************************************************************
'        処理概要：ActiveCell行に対して、新規行を追加し、固定値 または Excel関数設定列には
'                  同様の設定を行う
'
'            戻り値　：なし
'            引数１  ：ヘッダーポジション  typCellPos
'            引数２  ：項目ポジション      typCellPos
'*******************************************************************************
Private Sub ProcNewLine(posHead As typCellPos, _
                        posItm() As typCellPos)

    Dim i    As Integer

    ActiveCell.EntireRow.Insert

    i = 0                                                           ' 初期設定必要列の全データ分処理を行う
    Do While (i <= conArrayMax And _
              posItm(i).lngCol <> 0)
       
       Select Case Cells(posHead.lngRow, posItm(i).lngCol).Value
          Case "is"                                                 ' "is"列
             Cells(posItm(i).lngRow, posItm(i).lngCol).Value = "1"    ' 固定値 "1" 設定
          Case "発 生 日 付"                                        ' "発生日付"
                                                                      ' 曜日算出関数 設定
             Cells(posItm(i).lngRow, posItm(i).lngCol).Offset(, 1).Formula = _
                "=IF(AH" & ActiveCell.Row & "<>"""",""("" & CHOOSE(WEEKDAY(DATE(YEAR(AH" & ActiveCell.Row & _
                "),MONTH(AH" & ActiveCell.Row & "),DAY(AH" & ActiveCell.Row & _
                ")),1),""日"",""月"",""火"",""水"",""木"",""金"",""土"") & "")"","""")"
          Case "対 処 日 付"                                        ' "対処日付"
                                                                      ' 曜日算出関数 設定
             Cells(posItm(i).lngRow, posItm(i).lngCol).Offset(, 1).Formula = _
                "=IF(BJ" & ActiveCell.Row & "<>"""",""("" & CHOOSE(WEEKDAY(DATE(YEAR(BJ" & ActiveCell.Row & _
                "),MONTH(BJ" & ActiveCell.Row & "),DAY(BJ" & ActiveCell.Row & _
                ")),1),""日"",""月"",""火"",""水"",""木"",""金"",""土"") & "")"","""")"
          
       End Select
       
       i = i + 1
    Loop

End Sub
