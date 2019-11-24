Attribute VB_Name = "modGetSyukujitsu2"
'*******************************************************************************
'   祝日判定処理        ※年月指定により祝日(振休補正後)を配列で返す②
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'*******************************************************************************
Option Explicit
Private Const g_cnsFURI = "(振替休日)"
Private Const g_cnsKYU2 = "国民の休日"
' 祝日テーブル(ユーザー定義)
Public Type typSyuku
    dteDate As Date                 ' 日付
    intFuri As Integer              ' 振替休日SW(1=振替休日, 0=通常)
    strName As String               ' 祝日名称
End Type
' 下記処理で作成される祝日テーブル
Public g_tblSyuku() As typSyuku     ' 祝日テーブル(呼び元で利用する)

'*******************************************************************************
' 当該年月の祝日情報のテーブルを作成する(当月1ヶ月用)
'
' 戻り値：祝日テーブルの要素数(マイナス時は祝日なし)
' 引数　：Arg1=年(Integer)
' 　　　　Arg2=月(Integer)
'*******************************************************************************
Public Function FP_GetHoliday1(intY As Integer, _
                               intM As Integer) As Long
    Dim IX As Long              ' 配列のIndex

    ' 配列の初期化(要素数)
    IX = -1
    ReDim g_tblSyuku(0)         ' 一旦､初期化
    ' 祝日情報のテーブルを作成(1ヶ月分共通処理)
    Call GP_GetHolidaySub(intY, intM, IX)
    ' 戻り値のセット
    FP_GetHoliday1 = IX
End Function

'*******************************************************************************
' 前当翌3ヶ月の祝日情報のテーブルを作成する(当月+前後の3ヶ月用)
'
' 戻り値：祝日テーブルの要素数
' 引数　：Arg1=年(Integer)
' 　　　　Arg2=月(Integer)
'*******************************************************************************
Public Function FP_GetHoliday3(intYear As Integer, _
                               intMonth As Integer) As Long
    Dim intY As Integer, intM As Integer
    Dim IX As Long, IX2 As Long
    
    ' 配列の初期化(要素数)
    IX = -1
    ReDim g_tblSyuku(0)         ' 一旦､初期化
    ' 前月の年月を算出
    If intMonth = 1 Then
        intY = intYear - 1
        intM = 12
    Else
        intY = intYear
        intM = intMonth - 1
    End If
    ' 前・当・翌の3ヶ月を繰り返す
    For IX2 = 1 To 3
        ' 祝日情報のテーブルを作成(1ヶ月分共通処理)
        Call GP_GetHolidaySub(intY, intM, IX)
        ' 翌月の年月を算出
        If intM = 12 Then
            intY = intY + 1
            intM = 1
        Else
            intM = intM + 1
        End If
    Next IX2
    ' 戻り値をセット
    FP_GetHoliday3 = IX
End Function

'*******************************************************************************
' ※以下はサブ処理
'*******************************************************************************
' 祝日情報のテーブルを作成(1ヶ月分共通処理)
'
' 戻り値：(なし)
' 引数　：Arg1=年(Integer)
' 　　　　Arg2=月(Integer)
' 　　　　Arg3=テーブル最終位置(Long)  ※直前項目の登録位置
'*******************************************************************************
Private Sub GP_GetHolidaySub(intY As Integer, _
                             intM As Integer, _
                             IX As Long)
    Dim strName As String, strName2 As String
    
    ' 月による分岐
    Select Case intM
        '-----------------------------------------------------------------------
        ' 1月
        Case 1
            ' 元旦(1/1)
            Call GP_GetHolidaySub2(DateSerial(intY, intM, 1), IX, "元旦")
            ' 成人の日
            strName = "成人の日"
            If intY < 2000 Then
                ' 1999年までは15日固定
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 15), IX, strName)
            Else
                ' 2000年以降は第2月曜日
                Call GP_GetHolidaySub3(intY, intM, 2, 2, IX, strName)
            End If
        '-----------------------------------------------------------------------
        ' 2月
        Case 2
            ' 建国記念の日(2/11)
            Call GP_GetHolidaySub2(DateSerial(intY, intM, 11), IX, "建国記念の日")
        '-----------------------------------------------------------------------
        ' 3月
        Case 3
            ' 春分の日(※専用処理)
            Call GP_GetSyunbun(intY, IX)
        '-----------------------------------------------------------------------
        ' 4月
        Case 4
            ' みどりの日(4/29) ⇒ 昭和の日(2007年～)
            If intY >= 2007 Then
                strName = "昭和の日"
            Else
                strName = "みどりの日"
            End If
            Call GP_GetHolidaySub2(DateSerial(intY, intM, 29), IX, strName)
        '-----------------------------------------------------------------------
        ' 5月
        Case 5
            strName = "憲法記念日"
            strName2 = "子供の日"
            If intY >= 1985 Then
                IX = IX + 3
                ReDim Preserve g_tblSyuku(IX)
                ' 憲法記念日(5/3)
                g_tblSyuku(IX - 2).dteDate = DateSerial(intY, intM, 3)
                g_tblSyuku(IX - 2).strName = strName
                ' 国民の休日(5/4) ⇒ みどりの日(2007年～)
                g_tblSyuku(IX - 1).dteDate = DateSerial(intY, intM, 4)
                If intY >= 2007 Then
                    g_tblSyuku(IX - 1).strName = "みどりの日"
                Else
                    g_tblSyuku(IX - 1).strName = g_cnsKYU2
                End If
                ' 子供の日(5/5)
                If intY < 2007 Then
                    IX = IX - 1     ' 一旦減算(下位Procで加算されるため)
                    Call GP_GetHolidaySub2(DateSerial(intY, intM, 5), IX, strName2)
                Else
                    g_tblSyuku(IX).dteDate = DateSerial(intY, intM, 5)
                    g_tblSyuku(IX).strName = strName2
                    ' 2007年以降は5/3,5/4が日曜の場合も、5/6が振り返られる
                    If ((Weekday(g_tblSyuku(IX - 2).dteDate, vbSunday) = vbSunday) Or _
                        (Weekday(g_tblSyuku(IX - 1).dteDate, vbSunday) = vbSunday) Or _
                        (Weekday(g_tblSyuku(IX).dteDate, vbSunday) = vbSunday)) Then
                        IX = IX + 1
                        ReDim Preserve g_tblSyuku(IX)
                        g_tblSyuku(IX).dteDate = DateSerial(intY, intM, 6)
                        g_tblSyuku(IX).intFuri = 1
                        g_tblSyuku(IX).strName = g_cnsFURI
                    End If
                End If
            Else
                ' 憲法記念日(5/3)
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 3), IX, strName)
                ' 子供の日(5/5)
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 5), IX, strName2)
            End If
        '-----------------------------------------------------------------------
        ' 6月
        Case 6
            ' 祝日なし
        '-----------------------------------------------------------------------
        ' 7月
        Case 7
            If intY >= 1996 Then
                strName = "海の日"
                If intY >= 2003 Then
                    ' 海の日(第3月曜日)
                    Call GP_GetHolidaySub3(intY, intM, 3, 2, IX, strName)
                Else
                    ' 海の日(7/20)
                    Call GP_GetHolidaySub2(DateSerial(intY, intM, 20), IX, strName)
                End If
            End If
        '-----------------------------------------------------------------------
        ' 8月
        Case 8
            ' 祝日なし
        '-----------------------------------------------------------------------
        ' 9月
        Case 9
            strName = "敬老の日"
            If intY >= 2003 Then
                ' 敬老の日(第3月曜日)
                Call GP_GetHolidaySub3(intY, intM, 3, 2, IX, strName)
            Else
                ' 敬老の日(9/15)
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 15), IX, strName)
            End If
            ' 秋分の日(※専用処理)
            Call GP_GetSyuubun(intY, IX)
        '-----------------------------------------------------------------------
        ' 10月
        Case 10
            strName = "体育の日"
            If intY >= 2000 Then
                ' 体育の日(第2月曜日)
                Call GP_GetHolidaySub3(intY, intM, 2, 2, IX, strName)
            Else
                ' 体育の日(10/10)
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 10), IX, strName)
            End If
        '-----------------------------------------------------------------------
        ' 11月
        Case 11
            ' 文化の日(11/3)
            Call GP_GetHolidaySub2(DateSerial(intY, intM, 3), IX, "文化の日")
            ' 勤労感謝の日(11/23)
            Call GP_GetHolidaySub2(DateSerial(intY, intM, 23), IX, "勤労感謝の日")
        '-----------------------------------------------------------------------
        ' 12月
        Case 12
            If intY >= 1989 Then
                ' 天皇誕生日(12/23)
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 23), IX, "天皇誕生日")
            End If
    End Select
End Sub

'*******************************************************************************
' 当該祝日が日曜なら翌日を振替休日にしてテーブルセット(共通Sub処理)
'
' 戻り値：(なし)
' 引数　：Arg1=祝日日付(Date)
' 　　　　Arg2=テーブル最終位置(Long)  ※直前項目の登録位置
' 　　　　Arg3=祝日の名称(String)
'*******************************************************************************
Private Sub GP_GetHolidaySub2(dteHoliday As Date, _
                              IX As Long, _
                              strName As String)
    ' 当該祝日
    IX = IX + 1
    ReDim Preserve g_tblSyuku(IX)
    g_tblSyuku(IX).dteDate = dteHoliday
    g_tblSyuku(IX).strName = strName
    If Weekday(dteHoliday, vbSunday) = vbSunday Then
        ' 日曜と重なった場合の翌日を振替休日とする
        IX = IX + 1
        ReDim Preserve g_tblSyuku(IX)
        g_tblSyuku(IX).dteDate = dteHoliday + 1
        g_tblSyuku(IX).intFuri = 1          ' 振替休日
        g_tblSyuku(IX).strName = g_cnsFURI
    End If
End Sub

'*******************************************************************************
' 年月第n週のm曜日を算出してテーブルセット(共通Sub処理)
'
' 戻り値：(なし)
' 引数　：Arg1=年(Integer)
' 　　　　Arg2=月(Integer)
' 　　　　Arg3=週(Integer)
' 　　　　Arg4=曜日コード(Integer)     ※1=日曜, 2=月曜．．．7=土曜(2のみ利用)
' 　　　　Arg5=テーブル最終位置(Long)  ※直前項目の登録位置
' 　　　　Arg6=祝日の名称(String)
'*******************************************************************************
Private Sub GP_GetHolidaySub3(intY As Integer, _
                              intM As Integer, _
                              intW As Integer, _
                              intG As Integer, _
                              IX As Long, _
                              strName As String)
    Dim dteDate As Date
    Dim intG2 As Integer
    
    IX = IX + 1
    ReDim Preserve g_tblSyuku(IX)
    dteDate = DateSerial(intY, intM, 1)     ' 月初日
    intG2 = Weekday(dteDate, vbSunday)      ' 月初日の曜日
    If intG2 > intG Then intW = intW + 1    ' 初週調整
    g_tblSyuku(IX).dteDate = dteDate - intG2 + (intW - 1) * 7 + intG
    g_tblSyuku(IX).strName = strName
End Sub

'*******************************************************************************
' 春分の日の算出(簡易計算方式)
'
' 戻り値：(なし)
' 引数　：Arg1=年(Integer)
' 　　　　Arg2=テーブル最終位置(Long)  ※直前項目の登録位置
'*******************************************************************************
Private Sub GP_GetSyunbun(intY As Integer, _
                          IX As Long)
    Dim intD As Integer, intY2 As Integer, dteDate As Date
    
    ' 祝日法施行(1947年)以前,2151年以降(簡易計算不可)は無視
    IX = IX + 1
    ReDim Preserve g_tblSyuku(IX)
    intY2 = intY - 1980
    Select Case intY
        Case Is <= 1979
            intD = Int(20.8357 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Is <= 2099
            intD = Int(20.8431 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Else
            intD = Int(21.851 + (0.242194 * intY2) - Int(intY2 / 4))
    End Select
    dteDate = DateSerial(intY, 3, intD)
    ' 当該日付が日曜の場合は翌日とする(振替休日とはしない)
    If Weekday(dteDate, vbSunday) = vbSunday Then
        g_tblSyuku(IX).dteDate = dteDate + 1
    Else
        g_tblSyuku(IX).dteDate = dteDate
    End If
    g_tblSyuku(IX).strName = "春分の日"
End Sub

'*******************************************************************************
' 秋分の日の算出(簡易計算方式)
'
' 戻り値：(なし)
' 引数　：Arg1=年(Integer)
' 　　　　Arg2=テーブル最終位置(Long)  ※直前項目の登録位置
'*******************************************************************************
Private Sub GP_GetSyuubun(intY As Integer, _
                          IX As Long)
    Dim intD As Integer, intY2 As Integer, dteDate As Date
    
    ' 祝日法施行(1947年)以前,2151年以降(簡易計算不可)は無視
    intY2 = intY - 1980
    Select Case intY
        Case Is <= 1979
            intD = Int(23.2588 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Is <= 2099
            intD = Int(23.2488 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Else
            intD = Int(24.2488 + (0.242194 * intY2) - Int(intY2 / 4))
    End Select
    dteDate = DateSerial(intY, 9, intD)
    ' 当該日付が日曜の場合は翌日とする(振替休日とはしない)
    If Weekday(dteDate, vbSunday) = vbSunday Then
        dteDate = dteDate + 1
    End If
    ' 2003年以降は敬老の日の翌々日が秋分の日の場合､間の日は｢国民の休日｣になる
    If ((intY >= 2003) And ((dteDate - g_tblSyuku(IX).dteDate) = 2)) Then
        IX = IX + 2
        ReDim Preserve g_tblSyuku(IX)
        g_tblSyuku(IX - 1).dteDate = dteDate - 1
        g_tblSyuku(IX - 1).strName = g_cnsKYU2
    Else
        IX = IX + 1
        ReDim Preserve g_tblSyuku(IX)
    End If
    g_tblSyuku(IX).dteDate = dteDate
    g_tblSyuku(IX).strName = "秋分の日"
End Sub

'--------------------------------<< End of Source >>----------------------------

