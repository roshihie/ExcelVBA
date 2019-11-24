Attribute VB_Name = "Module4"
Option Explicit

Dim isERR As Integer                                                ' ERROR フラグ

Public Type typDatNam                                               ' 祝日格納型
    datDate As Date
    strName As String
End Type

Public Type typCellPos                                              ' セルポジション型
    lngRow As Long
    lngCol As Long
End Type

Dim lngSheetLastRow As Long                                         ' セル最終行

'*******************************************************************************
'        ＣＳＫ納品書セットアップ処理
'*******************************************************************************
Public Sub 納品書セットアップ()

    Dim posStart    As typCellPos
    Dim posEnd      As typCellPos
    Dim intEigyoCnt As Integer
    
    lngSheetLastRow = ActiveSheet.Rows.Count
    
    Call ProcInit(posStart, posEnd)
    If isERR Then
        Exit Sub
    End If
    
    Call ProcDetailInitSet(posStart, posEnd)
                           
    Call ProcHolidaySet(posStart, posEnd, intEigyoCnt)
    
    Call ProcDetailTimeSet(posStart, posEnd, intEigyoCnt)

End Sub

'*******************************************************************************
'        初　期　処　理
'*******************************************************************************
'        処理概要：日付エリアの開始行,列、および終了行,列を取得する
'
'            戻り値　：なし
'            引数１　：開始ポジション   typCellPos
'            引数２　：終了ポジション   typCellPos
'*******************************************************************************
Private Sub ProcInit(posStart As typCellPos, posEnd As typCellPos)
    
    Dim rngFind As Range
    Dim intRtn As Integer
    
    isERR = False
    Set rngFind = Columns(2).Find(What:="21", _
                                  After:=Range("B1"), _
                                  Lookat:=xlWhole, _
                                  Matchbyte:=True)
    If rngFind Is Nothing Then
        intRtn = MsgBox(prompt:="日付：21日 が 見つかりません", Buttons:=vbOKOnly + vbCritical)
        If intRtn = vbOK Then
            isERR = True
            Exit Sub
        End If
    Else
        posStart.lngRow = rngFind.Row
        posStart.lngCol = rngFind.Column
    End If
        
    Set rngFind = Columns(2).Find(What:="20", _
                                  After:=Range("B1"), _
                                  Lookat:=xlWhole, _
                                  Matchbyte:=True)
    If rngFind Is Nothing Then
        intRtn = MsgBox(prompt:="日付：20日 が 見つかりません", Buttons:=vbOKOnly + vbCritical)
        If intRtn = vbOK Then
            isERR = True
            Exit Sub
        End If
    Else
        posEnd.lngRow = rngFind.Row
        posEnd.lngCol = rngFind.Column
    End If
        
End Sub

'*******************************************************************************
'        納品書　日付行　初期設定
'*******************************************************************************
'        処理概要：月末日，曜日の設定  および
'                  開始時間，終了時間，休憩時間，コメント欄，作業内容
'                  クリア＆フォント黒に設定する
'
'            戻り値　：なし
'            引数１　：開始ポジション   typCellPos
'            引数２　：終了ポジション   typCellPos
'*******************************************************************************
Private Sub ProcDetailInitSet(posStart As typCellPos, posEnd As typCellPos)

    Const conWeekday As String = "日月火水木金土"
    Const conBlack   As Variant = 1
    
    Dim idxRow As Long
    Dim datRowdat As Date, datEndMonth As Date
    Dim lngEndMonth As Long
                                        ' 自年月日 の 月末日算出
    datEndMonth = DateSerial(Year(Range("L3").Value), Month(Range("L3").Value), 0)
    lngEndMonth = Day(datEndMonth)
    
    For idxRow = posStart.lngRow To posEnd.lngRow
                                        ' 日付，曜日設定
        If Cells(idxRow, 2).Value = "" Then                         ' 当行の日付がブランクのとき
            If Cells(idxRow - 1, 2).Value = "" Then                     ' 前行の日付もブランクのとき
                                                                            ' 処理なし
            ElseIf Cells(idxRow - 1, 2).Value < lngEndMonth Then        ' 前行の日付ありのとき
                Cells(idxRow, 2).Value = Cells(idxRow - 1, 2).Value + 1     '当行の日付 ← 前行の日付＋1
            End If
        End If
        Select Case Cells(idxRow, 2).Value
            Case Is > 20                                            ' 当行の日付＞20 のとき
                If Cells(idxRow, 2).Value > lngEndMonth Then               ' 当行の日付＞自年月日の月末日 のとき
                    Range(Cells(idxRow, 2), Cells(idxRow, 3)).ClearContents       ' 当行の日付，曜日クリア
                Else                                                    ' 当行の日付≦自年月日の月末日 のとき
                    datRowdat = DateSerial(Year(Range("F3").Value), _
                                          Month(Range("F3").Value), _
                                          Cells(idxRow, 2).Value)              ' 当行の曜日セット
                    Cells(idxRow, 3).Value = Mid(conWeekday, Weekday(datRowdat, vbSunday), 1)
                End If
                                                                    ' 当行の日付＝1日～20日まで
            Case Is >= 1
                datRowdat = DateSerial(Year(Range("L3").Value), _
                                      Month(Range("L3").Value), _
                                      Cells(idxRow, 2).Value)                  ' 当行の曜日セット
                Cells(idxRow, 3).Value = Mid(conWeekday, Weekday(datRowdat, vbSunday), 1)
            Case Else                                               ' 当行の日付がブランクのとき 処理
        End Select
                                        ' 日付，曜日　フォント黒
        With Range(Cells(idxRow, 2), Cells(idxRow, 3))
            .Font.ColorIndex = conBlack
        End With
                                        ' 開始時間，終了時間 クリア＆フォント黒
        With Range(Cells(idxRow, 4), Cells(idxRow, 7))
            .ClearContents
            .Font.ColorIndex = conBlack
        End With
                                        ' 休憩時間 クリア＆フォント黒
        With Range(Cells(idxRow, 10), Cells(idxRow, 11))
            .ClearContents
            .Font.ColorIndex = conBlack
        End With
                                        ' コメント欄，作業内容 クリア＆フォント黒
        With Range(Cells(idxRow, 14), Cells(idxRow, 17))
            .ClearContents
            .Font.ColorIndex = conBlack
        End With
    Next

End Sub

'*******************************************************************************
'        納品書　日付行　土日，休日設定
'*******************************************************************************
'*******************************************************************************
'        処理概要：祝日の配列を取得し、日付行の該当日付が祝日のとき
'                  作業内容欄に祝日名称を表示しフォント赤に設定し
'                  土曜日はフォント青，日曜日はフォント赤に設定する
'                  また、営業日をカウントして返す
'
'            戻り値　：なし
'            引数１　：開始ポジション   typCellPos
'            引数２　：終了ポジション   typCellPos
'            引数３　：営業日カウント   Integer
'*******************************************************************************
Private Sub ProcHolidaySet(posStart As typCellPos, posEnd As typCellPos, _
                           intEigyoCnt As Integer)

    Const conHoliday    As Integer = 1
    Const conNotHoliday As Integer = 0
    Const conBlue       As Variant = 5
    Const conRed        As Variant = 3
    
    Dim aryHoliday() As typDatNam
    Dim intHolidayCnt As Integer
    
    Dim intYYYY As Integer, intMM As Integer, intMMCnt As Integer
    Dim isHoliday As Boolean
    Dim stsHoliday As Integer
    Dim stsWeekday As Integer

    Dim datRowDate As Date
    Dim idxRow As Long, idxArray As Long
    
    intYYYY = Year(Range("F3").Value)
    intMM = Month(Range("F3").Value)
    intMMCnt = 2
    intEigyoCnt = 0
    intHolidayCnt = -1
    isHoliday = FuncHolidayGet(intYYYY, intMM, intMMCnt, aryHoliday, intHolidayCnt)
    
    For idxRow = posStart.lngRow To posEnd.lngRow
        If Cells(idxRow, 2).Value <> "" Then
            If Cells(idxRow, 2).Value > 20 Then
                datRowDate = DateSerial(Year(Range("F3").Value), Month(Range("F3").Value), Cells(idxRow, 2).Value)
            Else
                datRowDate = DateSerial(Year(Range("L3").Value), Month(Range("L3").Value), Cells(idxRow, 2).Value)
            End If
            
            stsHoliday = conNotHoliday
            idxArray = 0
            
            Do While (idxArray <= intHolidayCnt)                    ' VBA の判定は 完全評価(Complete)型であり、すべての条件文を
                                                                    ' 判定指定して True/False を決定する（⇔短絡評価(Short-circuit)型）
                If aryHoliday(idxArray).datDate > datRowDate Then
                    Exit Do
                End If
            
                If datRowDate = aryHoliday(idxArray).datDate Then
                
                    stsHoliday = conHoliday
                    Range(Cells(idxRow, 2), Cells(idxRow, 3)).Font.ColorIndex = conRed
                    Cells(idxRow, 15).Value = aryHoliday(idxArray).strName
                    Cells(idxRow, 15).Font.ColorIndex = conRed
                    
                End If
                idxArray = idxArray + 1
            
            Loop
            
            stsWeekday = Weekday(datRowDate, vbSunday)
            
            If stsHoliday = conHoliday Then
            ElseIf stsWeekday = vbSaturday Then
                Range(Cells(idxRow, 2), Cells(idxRow, 3)).Font.ColorIndex = conBlue
            ElseIf stsWeekday = vbSunday Then
                Range(Cells(idxRow, 2), Cells(idxRow, 3)).Font.ColorIndex = conRed
            Else
                intEigyoCnt = intEigyoCnt + 1
            End If
        End If
    Next

End Sub

'*******************************************************************************
'        祝 日 取 得  制 御 処 理
'*******************************************************************************
'        処理概要：開始年月から指定された月数分の祝日を取得するための
'                  制御を行い、祝日の配列を返す
'
'            戻り値　：祝日ありなし     Boolean    (True:あり，False:なし)
'            引数１　：開始年           Integer
'            引数２　：開始月           Integer
'            引数３　：取得月数         Integer
'            引数４　：祝日の配列       typDatNam
'            引数５　：配列格納件数     Integer
'*******************************************************************************
Private Function FuncHolidayGet(intYYYY As Integer, _
                                intMM As Integer, _
                                intMMCnt As Integer, _
                                aryHoliday() As typDatNam, _
                                intHolidayCnt As Integer) As Boolean
    Dim idxRow        As Integer
    Dim intCurYYYY As Integer, intCurMM As Integer

    If IsMissing(intMMCnt) Then
        intMMCnt = 1
    End If
    
    intCurYYYY = intYYYY
    intCurMM = intMM
    ReDim aryHoliday(0)
    
    For idxRow = 1 To intMMCnt
    
        Call ProcHolidayGet(intCurYYYY, intCurMM, aryHoliday, intHolidayCnt)
        If intCurMM >= 12 Then
            intCurYYYY = intCurYYYY + 1
            intCurMM = 1
        Else
            intCurMM = intCurMM + 1
        End If
        
    Next
    
    If intHolidayCnt >= 0 Then
        FuncHolidayGet = True
    Else
        FuncHolidayGet = False
    End If
    
End Function

'*******************************************************************************
'        祝 日 取 得 処 理
'*******************************************************************************
'        処理概要：開始年月から指定された月数分の祝日を取得するための
'                  制御を行い、祝日の配列を返す
'
'            戻り値　：なし
'            引数１　：年               Integer
'            引数２　：月               Integer
'            引数３　：祝日配列         Date       array
'            引数４　：配列格納件数     Integer
'*******************************************************************************
Private Sub ProcHolidayGet(intCurYYYY As Integer, _
                   intCurMM As Integer, _
                   aryHoliday() As typDatNam, _
                   intHolidayCnt As Integer)
                                        ' 月判定により処理分岐
    Select Case intCurMM
        Case 1                              ' 1月
            intHolidayCnt = intHolidayCnt + 4
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' 元旦(1/1) (振替休日なし)
            aryHoliday(intHolidayCnt - 3).strName = "元旦"
            aryHoliday(intHolidayCnt - 3).datDate = DateSerial(intCurYYYY, intCurMM, 1)
                                                ' 会社休日(1/2) (振替休日なし)
            aryHoliday(intHolidayCnt - 2).strName = "会社休日(年末年始)"
            aryHoliday(intHolidayCnt - 2).datDate = DateSerial(intCurYYYY, intCurMM, 2)
                                                ' 会社休日(1/3) (振替休日なし)
            aryHoliday(intHolidayCnt - 1).strName = "会社休日(年末年始)"
            aryHoliday(intHolidayCnt - 1).datDate = DateSerial(intCurYYYY, intCurMM, 3)
                                                ' 成人の日
            aryHoliday(intHolidayCnt).strName = "成人の日"
            If intCurYYYY < 2000 Then               ' 1999年まで 15日固定(1/15)
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 15))
            Else                                    ' 2000年以降 第2月曜日(ハッピーマンデー)
                aryHoliday(intHolidayCnt).datDate = FuncHappyMonday(intCurYYYY, intCurMM, 2, vbMonday)
            End If
                
        Case 2                              ' 2月
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' 建国記念日(2/11)(振替休日あり)
            aryHoliday(intHolidayCnt).strName = "建国記念の日"
            Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 11))
            
        Case 3                              ' 3月
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' 春分の日 取得処理
            aryHoliday(intHolidayCnt).strName = "春分の日"
            Call ProcSyunbunDay(aryHoliday, intHolidayCnt, intCurYYYY)
            
        Case 4                              ' 4月
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' みどりの日(4/29)⇒(2007年以降)昭和の日(4/29) (振替休日あり)
            If intCurYYYY < 2007 Then
                aryHoliday(intHolidayCnt).strName = "みどりの日"
            Else
                aryHoliday(intHolidayCnt).strName = "昭和の日"
            End If
            Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 29))
            
        Case 5                              ' 5月
            If intCurYYYY >= 1985 Then          ' 1985年以上
                
                If intCurYYYY < 2007 Then       ' 2007年未満
                
                    intHolidayCnt = intHolidayCnt + 3
                    ReDim Preserve aryHoliday(intHolidayCnt)
                                                    ' 憲法記念日(5/3) (振替休日なし)
                    aryHoliday(intHolidayCnt - 2).strName = "憲法記念日"
                    aryHoliday(intHolidayCnt - 2).datDate = DateSerial(intCurYYYY, intCurMM, 3)
                                                    ' 国民の休日(5/4)⇒(2007年以降)みどりの日(5/4) (振替休日なし)
                    aryHoliday(intHolidayCnt - 1).strName = "国民の休日"
                    aryHoliday(intHolidayCnt - 1).datDate = DateSerial(intCurYYYY, intCurMM, 4)
                                                    ' こどもの日(5/5) (振替休日あり)
                    aryHoliday(intHolidayCnt).strName = "こどもの日"
                    Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 5))
                Else                            ' 2007年以降
                                                    '5/3,5/4,5/5 がいずれかが日曜日のとき 5/6 振替休日
                    If (Weekday(DateSerial(intCurYYYY, intCurMM, 3), vbSunday) = vbSunday Or _
                        Weekday(DateSerial(intCurYYYY, intCurMM, 4), vbSunday) = vbSunday) Then
                        
                        intHolidayCnt = intHolidayCnt + 4
                        ReDim Preserve aryHoliday(intHolidayCnt)
                        
                        aryHoliday(intHolidayCnt - 3).strName = "憲法記念日"
                        aryHoliday(intHolidayCnt - 3).datDate = DateSerial(intCurYYYY, intCurMM, 3)
                        aryHoliday(intHolidayCnt - 2).strName = "みどりの日"
                        aryHoliday(intHolidayCnt - 2).datDate = DateSerial(intCurYYYY, intCurMM, 4)
                        aryHoliday(intHolidayCnt - 1).strName = "こどもの日"
                        aryHoliday(intHolidayCnt - 1).datDate = DateSerial(intCurYYYY, intCurMM, 5)
                        aryHoliday(intHolidayCnt).strName = "振替休日"
                        aryHoliday(intHolidayCnt).datDate = DateSerial(intCurYYYY, intCurMM, 6)
                    Else
                        intHolidayCnt = intHolidayCnt + 3
                        ReDim Preserve aryHoliday(intHolidayCnt)
                        
                        aryHoliday(intHolidayCnt - 2).strName = "憲法記念日"
                        aryHoliday(intHolidayCnt - 2).datDate = DateSerial(intCurYYYY, intCurMM, 3)
                        aryHoliday(intHolidayCnt - 1).strName = "みどりの日"
                        aryHoliday(intHolidayCnt - 1).datDate = DateSerial(intCurYYYY, intCurMM, 4)
                        aryHoliday(intHolidayCnt).strName = "こどもの日"
                        Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 5))
                    End If
                End If
            Else                                ' 1985年未満
                                                    ' 憲法記念日(5/3) (振替休日あり)
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "憲法記念日"
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 3))
                                                    ' こどもの日(5/5) (振替休日あり)
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "こどもの日"
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 5))
            End If
                
        Case 6                              ' 6月
                                                ' 祝日なし
        Case 7                              ' 7月
            If intCurYYYY >= 1996 Then          ' 1996年以降
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                                                    ' 海の日
                aryHoliday(intHolidayCnt).strName = "海の日"
                If intCurYYYY >= 2003 Then              ' 2003年以降 第3月曜日(ハッピーマンデー)
                    aryHoliday(intHolidayCnt).datDate = FuncHappyMonday(intCurYYYY, intCurMM, 3, vbMonday)
                Else                                    ' 2002年まで 20日固定(7/20) (振替休日あり)
                    Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 20))
                End If
            End If
            
        Case 8                              ' 8月
                                                ' 祝日なし
        Case 9                              ' 9月
            
            If intCurYYYY >= 2003 Then          ' 2003年以降
                                                    ' 敬老の日 第3月曜日(ハッピーマンデー)
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "敬老の日"
                aryHoliday(intHolidayCnt).datDate = FuncHappyMonday(intCurYYYY, intCurMM, 3, vbMonday)
                                                    ' 秋分の日 取得処理
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "秋分の日"
                Call ProcSyuubunDay(aryHoliday, intHolidayCnt, intCurYYYY)
                                                    ' 国民の休日 判定(敬老の日と秋分の日の間に1日空いたとき 国民の休日)
                If (aryHoliday(intHolidayCnt).datDate - aryHoliday(intHolidayCnt - 1).datDate) = 2 Then
                
                    intHolidayCnt = intHolidayCnt + 1
                    ReDim Preserve aryHoliday(intHolidayCnt)
                    
                    aryHoliday(intHolidayCnt) = aryHoliday(intHolidayCnt - 1)
                                                        ' 国民の休日
                    aryHoliday(intHolidayCnt - 1).strName = "国民の休日"
                    aryHoliday(intHolidayCnt - 1).datDate = aryHoliday(intHolidayCnt - 2).datDate + 1
                End If
            Else                                ' 2003年未満
                                                    ' 敬老の日 15日固定(9/15) (振替休日あり)
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "敬老の日"
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 15))
                                                    ' 秋分の日 取得処理
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "秋分の日"
                Call ProcSyuubunDay(aryHoliday, intHolidayCnt, intCurYYYY)
            End If
            
        Case 10                             ' 10月
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' 体育の日
            aryHoliday(intHolidayCnt).strName = "体育の日"
            If intCurYYYY >= 2000 Then              ' 2000年以降 第2月曜日(ハッピーマンデー)
                aryHoliday(intHolidayCnt).datDate = FuncHappyMonday(intCurYYYY, intCurMM, 2, vbMonday)
            Else                                    ' 1999年未満 10日固定(10/10) (振替休日あり)
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 10))
            End If
            
        Case 11                             ' 11月
                                                ' 文化の日(11/3) (振替休日あり)
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
            aryHoliday(intHolidayCnt).strName = "文化の日"
            Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 3))
                                                ' 勤労感謝の日(11/23) (振替休日あり)
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
            aryHoliday(intHolidayCnt).strName = "勤労感謝の日"
            Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 23))
            
        Case 12                             ' 12月
            If intCurYYYY >= 1989 Then          ' 1989年以降
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                                                    ' 天皇誕生日(12/23) (振替休日あり)
                aryHoliday(intHolidayCnt).strName = "天皇誕生日"
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 23))
            End If
            
            intHolidayCnt = intHolidayCnt + 3
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' 会社休日(12/29) (振替休日なし)
            aryHoliday(intHolidayCnt - 2).strName = "会社休日(年末年始)"
            aryHoliday(intHolidayCnt - 2).datDate = DateSerial(intCurYYYY, intCurMM, 29)
                                                ' 会社休日(12/30) (振替休日なし)
            aryHoliday(intHolidayCnt - 1).strName = "会社休日(年末年始)"
            aryHoliday(intHolidayCnt - 1).datDate = DateSerial(intCurYYYY, intCurMM, 30)
                                                ' 会社休日(12/31) (振替休日なし)
            aryHoliday(intHolidayCnt).strName = "会社休日(年末年始)"
            aryHoliday(intHolidayCnt).datDate = DateSerial(intCurYYYY, intCurMM, 31)
            
    End Select
        
End Sub

'*******************************************************************************
'        振替休日判定・設定処理
'*******************************************************************************
'        処理概要：祝日が日曜日ならば 振替休日として 月曜日を休日とする
'
'            戻り値　：なし
'            引数１　：祝日配列         Date       array
'            引数２　：配列格納件数     Integer
'            引数３　：振替前休日       Date
'*******************************************************************************
Private Sub ProcFuriHoliday(aryHoliday() As typDatNam, _
                            intHolidayCnt As Integer, _
                            datHoliday As Date)

    aryHoliday(intHolidayCnt).datDate = datHoliday                  ' 振替前休日 設定
    
    If Weekday(datHoliday, vbSunday) = vbSunday Then                ' 振替前休日＝日曜日のとき
    
        intHolidayCnt = intHolidayCnt + 1
        ReDim Preserve aryHoliday(intHolidayCnt)
                
        aryHoliday(intHolidayCnt).datDate = datHoliday + 1              ' 振替後休日 設定
        aryHoliday(intHolidayCnt).strName = "振替休日"
    End If

End Sub

'*******************************************************************************
'        ハッピーマンデー取得処理
'*******************************************************************************
'        処理概要：祝日の指定された週の指定曜日を算出する
'                  (ハッピーマンデーの指定曜日はすべて月曜日である)
'
'            戻り値　：ハッピーマンデー Date
'            引数１　：指定年           Integer
'            引数２　：指定月           Integer
'            引数３　：指定週           Integer
'            引数４　：指定曜日         Integer
'*******************************************************************************
Private Function FuncHappyMonday(intCurYYYY As Integer, _
                         intCurMM As Integer, _
                         intWeekNo As Integer, _
                         intWeekday As Integer) As Date

    Dim dat1stMonth  As Date
    Dim int1stWeekday As Integer
    
    dat1stMonth = DateSerial(intCurYYYY, intCurMM, 1)               ' 月初日      算出
    int1stWeekday = Weekday(dat1stMonth, vbSunday)                  ' 月初日 曜日 算出

    If intWeekday < int1stWeekday Then                              ' 指定曜日＜月初日の曜日 のとき
        intWeekNo = intWeekNo + 1                                       ' 指定週＋1
    End If
                                        ' 前月の最終土曜日 → 指定週の前週の土曜日 → 指定週の指定曜日 算出
    FuncHappyMonday = dat1stMonth - int1stWeekday + (intWeekNo - 1) * 7 + intWeekday

End Function

'*******************************************************************************
'        春分の日　取得処理
'*******************************************************************************
'        処理概要：指定された年の春分の日を取得して、振替休日判定・設定を行う
'
'                  20.8357：20日20時3分，20.8341：20日20時1分，21.851：21日20時25分
'                  0.242194：１年間で365日を超える時間＝5時間48分
'
'            戻り値　：なし
'            引数１　：祝日配列         Date       array
'            引数２　：配列格納件数     Integer
'            引数３　：指定年           Integer
'*******************************************************************************
Private Sub ProcSyunbunDay(aryHoliday() As typDatNam, _
                           intHolidayCnt As Integer, _
                           intCurYYYY As Integer)

    Const conMM  As Integer = 3
    Dim intDD As Integer, intYYDiff1980 As Integer
                                        ' 祝日法施行(1947年)以前，2151年以降(簡易計算不可)は無視
    intYYDiff1980 = intCurYYYY - 1980                               ' 指定年と1980年の差分を算出
    
    Select Case intCurYYYY
        Case Is <= 1979                                             ' 指定年≦1979年 のとき
            intDD = Int(20.8357 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
        Case Is <= 2099                                             ' 指定年≦2099年 のとき
            intDD = Int(20.8341 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
        Case Else                                                   ' 指定年＞2100年 のとき
            intDD = Int(21.851 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
    End Select
    
    aryHoliday(intHolidayCnt).strName = "春分の日"
    Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, conMM, intDD))

End Sub

'*******************************************************************************
'        秋分の日　取得処理
'*******************************************************************************
'        処理概要：指定された年の秋分の日を取得して、振替休日判定・設定を行う
'
'                  23.2588：23日6時12分，23.2488：23日5時58分，24.2488：24日5時58分
'                  0.242194：１年間で365日を超える時間＝5時間48分
'
'            戻り値　：なし
'            引数１　：祝日配列         Date       array
'            引数２　：配列格納件数     Integer
'            引数３　：指定年           Integer
'*******************************************************************************
Private Sub ProcSyuubunDay(aryHoliday() As typDatNam, _
                           intHolidayCnt As Integer, _
                           intCurYYYY As Integer)

    Const conMM  As Integer = 9
    Dim intYYDiff1980 As Integer, intDD As Integer
                                        ' 祝日法施行(1947年)以前，2151年以降(簡易計算不可)は無視
    intYYDiff1980 = intCurYYYY - 1980                                  ' 指定年と1980年の差分を算出
    
    Select Case intCurYYYY
        Case Is <= 1979                                             ' 指定年≦1979年 のとき
            intDD = Int(23.2588 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
        Case Is <= 2099                                             ' 指定年≦2099年 のとき
            intDD = Int(23.2488 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
        Case Else                                                   ' 指定年＞2100年 のとき
            intDD = Int(24.2488 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
    End Select
    
    aryHoliday(intHolidayCnt).strName = "秋分の日"
    Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, conMM, intDD))

End Sub

'*******************************************************************************
'        納品書　開始時間，終了時間，休憩時間　設定
'*******************************************************************************
'        処理概要：納品書の日付行の開始時間，終了時間，休憩時間の設定を行う
'
'            戻り値　：なし
'            引数１　：開始ポジション   typCellPos
'            引数２　：終了ポジション   typCellPos
'            引数３　：営業日数カウント Integer
'*******************************************************************************
Private Sub ProcDetailTimeSet(posStart As typCellPos, posEnd As typCellPos, _
                              intEigyoCnt As Integer)
                              
    Const conMonthlyTime As Integer = 175
    Const varFromTime    As Variant = "08:50:00"
    Const con1dayToTime  As Date = "17:20:00"
    Const conBlack       As Variant = 1

    Dim dblCalcTime   As Double
    Dim dblTimeSho    As Double
    Dim int1dayHour   As Integer
    Dim dbl1dayTime   As Double
    Dim timToTime     As Date
    Dim is1dayMini    As Boolean
    Dim idxRow        As Long
    
    dblCalcTime = conMonthlyTime / intEigyoCnt
    int1dayHour = Int(dblCalcTime)
    
    dblTimeSho = dblCalcTime - int1dayHour
    
    Select Case dblTimeSho
        Case Is > 0.75
            dbl1dayTime = int1dayHour + 1
        Case Is > 0.5
            dbl1dayTime = int1dayHour + 0.75
        Case Is > 0.25
            dbl1dayTime = int1dayHour + 0.5
        Case Is > 0
            dbl1dayTime = int1dayHour + 0.25
        Case Else
            dbl1dayTime = int1dayHour
    End Select
    
    is1dayMini = True
    If dbl1dayTime = 7.5 Then
        dbl1dayTime = dbl1dayTime + 1                               ' 昼休み　　　　：1時間
    ElseIf dbl1dayTime > 7.5 Then
        is1dayMini = False
        dbl1dayTime = dbl1dayTime + 1.25                            ' 昼休み＋夕休み：1.25時間
    Else
    End If
    timToTime = DateAdd("n", dbl1dayTime * 60, varFromTime)
    
    For idxRow = posStart.lngRow To posEnd.lngRow
    
        If Cells(idxRow, 2).Value <> "" Then
            If Cells(idxRow, 2).Font.ColorIndex = conBlack Then
                Cells(idxRow, 4).Value = Hour(varFromTime)
                Cells(idxRow, 5).Value = Minute(varFromTime)
                Cells(idxRow, 6).Value = Hour(timToTime)
                Cells(idxRow, 7).Value = Minute(timToTime)
                
                Cells(idxRow, 10).Value = "1"
                If is1dayMini Then
                    Cells(idxRow, 11).Value = "00"
                Else
                    Cells(idxRow, 11).Value = "15"
                End If
            End If
        End If
    Next
    
End Sub

