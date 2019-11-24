VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_CALENDAR3 
   Caption         =   "日付選択"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2385
   OleObjectBlob   =   "FRM_CALENDAR3.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FRM_CALENDAR3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'   カレンダーフォーム3(日付入力部品)    ※ユーザーフォーム
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'*******************************************************************************
Option Explicit
'-------------------------------------------------------------------------------
' [起算曜日] ※カレンダーを月曜開始(曜日左端)にする場合は｢2｣に変更して下さい。
Private Const g_cnsStartYobi = 1                ' 1=日曜日,2=月曜日(他は不可)
'-------------------------------------------------------------------------------
' [年の表示限度(From/To)]
Private Const g_cnsYearFrom = 1947              ' 祝日法施行
Private Const g_cnsYearToAdd = 3                ' システム日の年+n年までの指定
'-------------------------------------------------------------------------------
' フォーム上の色指定等の定数
Private Const cnsBC_Select = &HFFCC33           ' 選択日付の背景色
Private Const cnsBC_Other = &HE0E0E0            ' 当月以外の背景色
Private Const cnsBC_Sunday = &HFFDDFF           ' 日曜の背景色
Private Const cnsBC_Saturday = &HDDFFDD         ' 土曜の背景色
Private Const cnsBC_Month = &HFFFFFF            ' 当月土日以外の背景色
Private Const cnsFC_Hori = &HFF                 ' 祝日の文字色
Private Const cnsFC_Normal = &HC00000           ' 祝日以外の文字色
Private Const cnsDefaultGuide = "矢印キーで操作できます。"
'-------------------------------------------------------------------------------
' フォーム表示中に保持するモジュール変数
Private tblDate(1 To 45) As MSForms.Label       ' 日付ラベル
Private tblDate2(1 To 45) As Date               ' 日付
Private tblYobi(1 To 45) As Integer             ' 曜日
Private tblGuide(1 To 45) As String             ' ガイド
Private g_intCurYear As Integer                 ' 現在表示年
Private g_intCurMonth As Integer                ' 現在表示月
Private g_FormDate1 As Date                     ' 現在日付
Private g_CurPos As Integer                     ' 現在日付位置
Private g_POS_F As Integer                      ' 月初日位置
Private g_POS_T As Integer                      ' 月末日位置
Private g_swBatch As Boolean                    ' イベント抑制SW
Private g_VisibleYear As Boolean                ' Conboの年表示スイッチ
Private g_VisibleMonth As Boolean               ' Comboの月表示スイッチ
Private g_intSunday As Integer                  ' 日曜日の曜日コード
Private g_intSaturday As Integer                ' 土曜日の曜日コード

'*******************************************************************************
' ■フォーム上のイベント
'*******************************************************************************
' ｢月｣コンボの操作イベント
'*******************************************************************************
Private Sub CBO_MONTH_Change()
    Dim intMonth As Integer

    If g_swBatch Then Exit Sub
    intMonth = CInt(CBO_MONTH.Text)
    g_FormDate1 = DateSerial(g_intCurYear, intMonth, 1)
    Call ERASE_YEAR_MONTH                       ' 年月コンボの非表示化
    ' カレンダー作成
    Call GP_MakeCalendar
End Sub

'*******************************************************************************
' ｢年｣コンボの操作イベント
'*******************************************************************************
Private Sub CBO_YEAR_Change()
    Dim intYear As Integer
    
    If g_swBatch Then Exit Sub
    intYear = CInt(CBO_YEAR.Text)
    g_FormDate1 = DateSerial(intYear, g_intCurMonth, 1)
    Call ERASE_YEAR_MONTH                       ' 年月コンボの非表示化
    ' カレンダー作成
    Call GP_MakeCalendar
End Sub

'*******************************************************************************
' 各日付ラベルのイベント(クラス処理はしないでそれぞれClickイベント等で受ける)
'*******************************************************************************
' 各日付ラベル(7曜×6週=42件、対応日付は表示時点で配列化されている)
Private Sub LBL_01_Click(): Call GP_ClickCalendar(tblDate2(1)):  End Sub
Private Sub LBL_02_Click(): Call GP_ClickCalendar(tblDate2(2)):  End Sub
Private Sub LBL_03_Click(): Call GP_ClickCalendar(tblDate2(3)):  End Sub
Private Sub LBL_04_Click(): Call GP_ClickCalendar(tblDate2(4)):  End Sub
Private Sub LBL_05_Click(): Call GP_ClickCalendar(tblDate2(5)):  End Sub
Private Sub LBL_06_Click(): Call GP_ClickCalendar(tblDate2(6)):  End Sub
Private Sub LBL_07_Click(): Call GP_ClickCalendar(tblDate2(7)):  End Sub
Private Sub LBL_08_Click(): Call GP_ClickCalendar(tblDate2(8)):  End Sub
Private Sub LBL_09_Click(): Call GP_ClickCalendar(tblDate2(9)):  End Sub
Private Sub LBL_10_Click(): Call GP_ClickCalendar(tblDate2(10)): End Sub
Private Sub LBL_11_Click(): Call GP_ClickCalendar(tblDate2(11)): End Sub
Private Sub LBL_12_Click(): Call GP_ClickCalendar(tblDate2(12)): End Sub
Private Sub LBL_13_Click(): Call GP_ClickCalendar(tblDate2(13)): End Sub
Private Sub LBL_14_Click(): Call GP_ClickCalendar(tblDate2(14)): End Sub
Private Sub LBL_15_Click(): Call GP_ClickCalendar(tblDate2(15)): End Sub
Private Sub LBL_16_Click(): Call GP_ClickCalendar(tblDate2(16)): End Sub
Private Sub LBL_17_Click(): Call GP_ClickCalendar(tblDate2(17)): End Sub
Private Sub LBL_18_Click(): Call GP_ClickCalendar(tblDate2(18)): End Sub
Private Sub LBL_19_Click(): Call GP_ClickCalendar(tblDate2(19)): End Sub
Private Sub LBL_20_Click(): Call GP_ClickCalendar(tblDate2(20)): End Sub
Private Sub LBL_21_Click(): Call GP_ClickCalendar(tblDate2(21)): End Sub
Private Sub LBL_22_Click(): Call GP_ClickCalendar(tblDate2(22)): End Sub
Private Sub LBL_23_Click(): Call GP_ClickCalendar(tblDate2(23)): End Sub
Private Sub LBL_24_Click(): Call GP_ClickCalendar(tblDate2(24)): End Sub
Private Sub LBL_25_Click(): Call GP_ClickCalendar(tblDate2(25)): End Sub
Private Sub LBL_26_Click(): Call GP_ClickCalendar(tblDate2(26)): End Sub
Private Sub LBL_27_Click(): Call GP_ClickCalendar(tblDate2(27)): End Sub
Private Sub LBL_28_Click(): Call GP_ClickCalendar(tblDate2(28)): End Sub
Private Sub LBL_29_Click(): Call GP_ClickCalendar(tblDate2(29)): End Sub
Private Sub LBL_30_Click(): Call GP_ClickCalendar(tblDate2(30)): End Sub
Private Sub LBL_31_Click(): Call GP_ClickCalendar(tblDate2(31)): End Sub
Private Sub LBL_32_Click(): Call GP_ClickCalendar(tblDate2(32)): End Sub
Private Sub LBL_33_Click(): Call GP_ClickCalendar(tblDate2(33)): End Sub
Private Sub LBL_34_Click(): Call GP_ClickCalendar(tblDate2(34)): End Sub
Private Sub LBL_35_Click(): Call GP_ClickCalendar(tblDate2(35)): End Sub
Private Sub LBL_36_Click(): Call GP_ClickCalendar(tblDate2(36)): End Sub
Private Sub LBL_37_Click(): Call GP_ClickCalendar(tblDate2(37)): End Sub
Private Sub LBL_38_Click(): Call GP_ClickCalendar(tblDate2(38)): End Sub
Private Sub LBL_39_Click(): Call GP_ClickCalendar(tblDate2(39)): End Sub
Private Sub LBL_40_Click(): Call GP_ClickCalendar(tblDate2(40)): End Sub
Private Sub LBL_41_Click(): Call GP_ClickCalendar(tblDate2(41)): End Sub
Private Sub LBL_42_Click(): Call GP_ClickCalendar(tblDate2(42)): End Sub
'-------------------------------------------------------------------------------
' 昨日、今日、明日ラベル
Private Sub LBL_43_Click(): Call GP_ClickCalendar(tblDate2(43)): End Sub
Private Sub LBL_44_Click(): Call GP_ClickCalendar(tblDate2(44)): End Sub
Private Sub LBL_45_Click(): Call GP_ClickCalendar(tblDate2(45)): End Sub
'-------------------------------------------------------------------------------
' 可変ガイドメッセージ
Private Sub LBL_01_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(1): End Sub
Private Sub LBL_02_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(2): End Sub
Private Sub LBL_03_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(3): End Sub
Private Sub LBL_04_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(4): End Sub
Private Sub LBL_05_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(5): End Sub
Private Sub LBL_06_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(6): End Sub
Private Sub LBL_07_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(7): End Sub
Private Sub LBL_08_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(8): End Sub
Private Sub LBL_09_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(9): End Sub
Private Sub LBL_10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(10): End Sub
Private Sub LBL_11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(11): End Sub
Private Sub LBL_12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(12): End Sub
Private Sub LBL_13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(13): End Sub
Private Sub LBL_14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(14): End Sub
Private Sub LBL_15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(15): End Sub
Private Sub LBL_16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(16): End Sub
Private Sub LBL_17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(17): End Sub
Private Sub LBL_18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(18): End Sub
Private Sub LBL_19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(19): End Sub
Private Sub LBL_20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(20): End Sub
Private Sub LBL_21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(21): End Sub
Private Sub LBL_22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(22): End Sub
Private Sub LBL_23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(23): End Sub
Private Sub LBL_24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(24): End Sub
Private Sub LBL_25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(25): End Sub
Private Sub LBL_26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(26): End Sub
Private Sub LBL_27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(27): End Sub
Private Sub LBL_28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(28): End Sub
Private Sub LBL_29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(29): End Sub
Private Sub LBL_30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(30): End Sub
Private Sub LBL_31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(31): End Sub
Private Sub LBL_32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(32): End Sub
Private Sub LBL_33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(33): End Sub
Private Sub LBL_34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(34): End Sub
Private Sub LBL_35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(35): End Sub
Private Sub LBL_36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(36): End Sub
Private Sub LBL_37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(37): End Sub
Private Sub LBL_38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(38): End Sub
Private Sub LBL_39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(39): End Sub
Private Sub LBL_40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(40): End Sub
Private Sub LBL_41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(41): End Sub
Private Sub LBL_42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(42): End Sub
Private Sub LBL_43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(43): End Sub
Private Sub LBL_44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(44): End Sub
Private Sub LBL_45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(45): End Sub
'-------------------------------------------------------------------------------
' 固定ガイド(曜日ラベル等)
Private Sub LBL_SUN_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_MON_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_TUE_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_WED_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_THU_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_FRI_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_SAT_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_PREV_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = "前月に戻ります(PageUp)": End Sub
Private Sub LBL_NEXT_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = "翌月に進みます(PageDown)": End Sub
Private Sub LBL_YM_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = "年か月を選択します。": End Sub
Private Sub LBL_YEAR_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = "年が選択できます。": End Sub
Private Sub LBL_MONTH_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = "月が選択できます。": End Sub

'*******************************************************************************
' 「←(前月)」Clickイベント
'*******************************************************************************
Private Sub LBL_PREV_Click()
    Call ERASE_YEAR_MONTH                       ' 年月コンボの非表示化
    g_FormDate1 = DateSerial(g_intCurYear, g_intCurMonth - 1, 1)
    ' カレンダー作成
    Call GP_MakeCalendar
End Sub

'*******************************************************************************
' 「→(翌月)」Clickイベント
'*******************************************************************************
Private Sub LBL_NEXT_Click()
    Call ERASE_YEAR_MONTH                       ' 年月コンボの非表示化
    g_FormDate1 = DateSerial(g_intCurYear, g_intCurMonth + 1, 1)
    ' カレンダー作成
    Call GP_MakeCalendar
End Sub

'*******************************************************************************
' 「月」Clickイベント
'*******************************************************************************
Private Sub LBL_MONTH_Click()
    Dim intMonth As Integer
    Dim IX As Long, CUR As Long
    
    Call ERASE_YEAR            ' 年コンボが表示されていたら消去
    ' 年コンボの表示
    g_swBatch = True
    With CBO_MONTH
        .Clear
        For intMonth = 1 To 12
            .AddItem Format(intMonth, "00")
            If intMonth = g_intCurMonth Then CUR = IX
            IX = IX + 1
        Next intMonth
        .ListIndex = CUR
        .Visible = True
        g_VisibleMonth = True
    End With
    g_swBatch = False
End Sub

'*******************************************************************************
' 「年」Clickイベント
'*******************************************************************************
Private Sub LBL_YEAR_Click()
    Dim intYear As Integer, intYearSTR As Integer, intYearEND As Integer
    Dim IX As Long, CUR As Long
    
    Call ERASE_MONTH            ' 月コンボが表示されていたら消去
    ' 年コンボの表示
    g_swBatch = True
    With CBO_YEAR
        .Clear
        intYearSTR = g_intCurYear - 10
        If intYearSTR < g_cnsYearFrom Then intYearSTR = g_cnsYearFrom
        intYearEND = g_intCurYear + 10
        intYear = Year(Date) + g_cnsYearToAdd
        If intYearEND > intYear Then intYearEND = intYear
        For intYear = intYearSTR To intYearEND
            .AddItem CStr(intYear)
            If intYear = g_intCurYear Then CUR = IX
            IX = IX + 1
        Next intYear
        .ListIndex = CUR
        .Visible = True
        g_VisibleYear = True
    End With
    g_swBatch = False
End Sub

'*******************************************************************************
' フォーム表示(繰り返し表示の場合はHideのみのためInitializeは起きない)
'*******************************************************************************
Private Sub UserForm_Activate()
    ' Tagから日付を取り出す
    g_FormDate1 = CDate(Me.Tag)
    ' Tagは非数値状態にしておく
    Me.Tag = False
    ' コンボは非表示
    CBO_YEAR.Visible = False
    CBO_MONTH.Visible = False
    g_VisibleYear = False
    g_VisibleMonth = False
    ' 初期の年月をセット
    If g_FormDate1 = 0 Then g_FormDate1 = Date
    ' カレンダー作成
    Call GP_MakeCalendar
    LBL_GUIDE.Caption = cnsDefaultGuide             ' ガイド表示
End Sub

'*******************************************************************************
' フォーム初期化(繰り返し表示の場合はHideのみのためInitializeは起きない)
'*******************************************************************************
Private Sub UserForm_Initialize()
    Dim IX As Integer
    Dim strName As String
    
    ' 起算曜日による曜日見出しの位置修正
    If g_cnsStartYobi = 2 Then
        ' 月曜起算
        LBL_MON.Left = 2
        LBL_TUE.Left = 18.5
        LBL_WED.Left = 35
        LBL_THU.Left = 51.5
        LBL_FRI.Left = 68
        LBL_SAT.Left = 84.5
        LBL_SUN.Left = 101
        ' 曜日コードの設定
        g_intSunday = 7
        g_intSaturday = 6
    Else
        ' 日曜起算
        LBL_SUN.Left = 2
        LBL_MON.Left = 18.5
        LBL_TUE.Left = 35
        LBL_WED.Left = 51.5
        LBL_THU.Left = 68
        LBL_FRI.Left = 84.5
        LBL_SAT.Left = 101
        ' 曜日コードの設定
        g_intSunday = 1
        g_intSaturday = 7
    End If
    ' 日付ラベルをObject型配列変数にセット(処理内ではこの変数で値を登録)
    For IX = 1 To 45
        strName = "LBL_" & Format(IX, "00")
        Set tblDate(IX) = Me.Controls(strName)
    Next IX
    g_swCalendar1Loaded = True              ' Load判定スイッチ(⇒Load)
End Sub

'*******************************************************************************
' フォーム上キーボード処理
'*******************************************************************************
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                             ByVal Shift As Integer)

    ' KeyCode(Shift併用)による制御
    Select Case KeyCode
        Case vbKeyReturn, vbKeyExecute, vbKeySeparator  ' Enter(決定)
            Call GP_ClickCalendar(g_FormDate1)
        Case vbKeyCancel, vbKeyEscape                   ' Cancel, Esc(終了)
            Me.Hide
        Case vbKeyPageDown                              ' PageDown(次月)
            Call LBL_NEXT_Click
        Case vbKeyPageUp                                ' KeyPageUp(前月)
            Call LBL_PREV_Click
        Case vbKeyRight, vbKeyNumpad6, vbKeyAdd         ' →(翌日)
            Call GP_MOVE_DAY(1)
        Case vbKeyLeft, vbKeyNumpad4, vbKeySubtract     ' ←(前日)
            Call GP_MOVE_DAY(-1)
        Case vbKeyUp, vbKeyNumpad8                      ' ↑(7日後)
            Call GP_MOVE_DAY(-7)
        Case vbKeyDown, vbKeyNumpad2                    ' ↓(7日前)
            Call GP_MOVE_DAY(7)
        Case vbKeyHome                                  ' Home(月初)
            Call GP_MOVE_DAY(g_POS_F - g_CurPos)
        Case vbKeyEnd                                   ' End(月末)
            Call GP_MOVE_DAY(g_POS_T - g_CurPos)
        Case vbKeyTab                                   ' Tab(Shiftによる)
            If Shift = 1 Then
                Call GP_MOVE_DAY(-1)            ' 前日
            Else
                Call GP_MOVE_DAY(1)             ' 翌日
            End If
        Case vbKeyF11                                   ' F11(前年)
            g_FormDate1 = DateSerial(g_intCurYear - 1, g_intCurMonth, 1)
            Call ERASE_YEAR_MONTH                       ' 年月コンボの非表示化
            ' カレンダー作成
            Call GP_MakeCalendar
        Case vbKeyF12                                   ' F12(翌年)
            g_FormDate1 = DateSerial(g_intCurYear + 1, g_intCurMonth, 1)
            Call ERASE_YEAR_MONTH                       ' 年月コンボの非表示化
            ' カレンダー作成
            Call GP_MakeCalendar
    End Select
End Sub

'*******************************************************************************
' フォーム上マウス移動
'*******************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    Me.LBL_GUIDE.Caption = cnsDefaultGuide
End Sub

'*******************************************************************************
' フォーム終了
'*******************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call ERASE_YEAR_MONTH                       ' 年月コンボの非表示化
    ' 閉じる[×]ボタンが押された時、Unloadされないようにする
    Cancel = True
    Me.Hide
End Sub

Private Sub UserForm_Terminate()
    g_swCalendar1Loaded = False             ' Load判定スイッチ(⇒Unload)
End Sub

'*******************************************************************************
' ■共通サブ処理
'*******************************************************************************
' カレンダー表示処理
'*******************************************************************************
Private Sub GP_MakeCalendar()
    Dim dteDate As Date, dteDate2 As Date, dteDateF As Date, dteDateT As Date
    Dim intYOBI As Integer, intYear As Integer
    Dim IX As Long, IX2 As Long, IXH As Long, IXH_MAX As Long
    Dim tblTODAY As Variant
    
    ' 指定年月が利用可能かチェック
    intYear = Year(g_FormDate1)                      ' 指定年
    If ((intYear < g_cnsYearFrom) Or _
        (intYear > (Year(Date) + g_cnsYearToAdd))) Then
        MsgBox "祝日計算範囲を超えています。", vbExclamation, Me.Caption
        g_FormDate1 = tblDate2(g_CurPos)
    End If
    g_intCurYear = Year(g_FormDate1)                 ' 指定年
    g_intCurMonth = Month(g_FormDate1)               ' 指定月
    dteDateF = DateSerial(g_intCurYear, g_intCurMonth, 1)       ' 月初日
    dteDateT = DateSerial(g_intCurYear, g_intCurMonth + 1, 0)   ' 月末日
    LBL_YM.Caption = g_intCurYear & "年" & Format(g_intCurMonth, "00") & "月"
    ' 前後3ヶ月の祝日テーブル取得(共通関数より)
    IXH_MAX = modGetSyukujitsu2.FP_GetHoliday3(g_intCurYear, g_intCurMonth)
    ' テーブル要素を1件追加しておく(ループ中の要素数判断を不要にする)
    IX = IXH_MAX + 1
    ReDim Preserve g_tblSyuku(IX)
    g_tblSyuku(IX).dteDate = DateSerial(g_intCurYear + 2, 1, 1)
    ' 指定日付から一旦、前週の最終日(土曜日)に戻す
    dteDate = DateSerial(g_intCurYear, g_intCurMonth, 1)        ' 月初日
    If g_cnsStartYobi = 2 Then
        intYOBI = Weekday(dteDate, vbMonday)                    ' 曜日の取得
    Else
        intYOBI = Weekday(dteDate, vbSunday)                    ' 曜日の取得
    End If
    dteDate = dteDate - intYOBI
    intYOBI = 0
    ' 先頭の祝日テーブル位置判定(マッチング利用のため)
    IXH = 0
    dteDate2 = dteDate + 1      ' カレンダー内の初日
    Do While IXH <= IXH_MAX
        If g_tblSyuku(IXH).dteDate >= dteDate2 Then Exit Do
        IXH = IXH + 1
    Loop
    '---------------------------------------------------------------------------
    ' フォーム上の日付セット(7曜×6週=42件固定)
    For IX = 1 To 42
        ' 当位置の日付、曜日を算出
        intYOBI = intYOBI + 1
        If intYOBI > 7 Then intYOBI = 1
        dteDate = dteDate + 1
        ' 日付は別テーブルにセット
        tblDate2(IX) = dteDate
        tblYobi(IX) = intYOBI
        tblGuide(IX) = Format(dteDate, cnsDateFormat) & _
            "(" & Format(dteDate, "aaa") & ")"
        If dteDate = dteDateF Then
            ' 当月初日
            g_POS_F = IX
        ElseIf dteDate = dteDateT Then
            ' 当月末日
            g_POS_T = IX
        End If
        ' ラベルコントロールを配列化した変数
        With tblDate(IX)
            ' ラベルに日付をセット
            .Caption = Day(dteDate)
            ' 月度、曜日によりラベルの書式をセット
            .Font.Bold = False
            .ForeColor = cnsFC_Normal
            If dteDate = g_FormDate1 Then
                ' 初期選択日付
                .BackColor = cnsBC_Select
                g_CurPos = IX
            ElseIf Month(dteDate) = g_intCurMonth Then
                ' 当月
                Select Case intYOBI
                    Case g_intSunday    ' 日曜日
                        .BackColor = cnsBC_Sunday
                    Case g_intSaturday  ' 土曜日
                        .BackColor = cnsBC_Saturday
                    Case Else
                        .BackColor = cnsBC_Month
                End Select
            Else
                ' 当月以外
                .BackColor = cnsBC_Other
            End If
            ' 祝日(含振替休日)の判定
            If g_tblSyuku(IXH).dteDate = dteDate Then
                ' 文字色を赤とする
                .ForeColor = cnsFC_Hori
                If Month(dteDate) = g_intCurMonth Then .Font.Bold = True
                tblGuide(IX) = tblGuide(IX) & " " & g_tblSyuku(IXH).strName
                ' 祝日テーブルの参照Indexを加算
                IXH = IXH + 1
            End If
        End With
    Next IX
    '---------------------------------------------------------------------------
    ' 昨日､今日､明日の処理
    dteDate = Date              ' 今日
    If ((Year(dteDate) <> g_intCurYear) Or (Month(dteDate) < g_intCurMonth)) Then
        IXH_MAX = modGetSyukujitsu2.FP_GetHoliday3(Year(dteDate), Month(dteDate))
    End If
    IXH = 0
    dteDate = Date - 1          ' 昨日
    Do While IXH <= IXH_MAX
        If g_tblSyuku(IXH).dteDate >= dteDate Then Exit Do
        IXH = IXH + 1
    Loop
    tblTODAY = Array("[昨日]", "[今日]", "[明日]")
    For IX = 43 To 45
        tblDate2(IX) = dteDate
        tblGuide(IX) = tblTODAY(IX2) & Format(dteDate, cnsDateFormat) & _
            "(" & Format(dteDate, "aaa") & ")"
        ' 祝日(含振替休日)の判定
        If IXH <= IXH_MAX Then
            ' 祝日テーブルの日付との一致を判定
            If g_tblSyuku(IXH).dteDate = dteDate Then
                tblGuide(IX) = tblGuide(IX) & " " & g_tblSyuku(IXH).strName
                ' 祝日テーブルの参照Indexを加算
                IXH = IXH + 1
            End If
        End If
        dteDate = dteDate + 1
        IX2 = IX2 + 1
    Next IX
    LBL_GUIDE.Caption = tblGuide(g_CurPos)      ' ガイド表示
End Sub

'*******************************************************************************
' カレンダー上の移動処理
'*******************************************************************************
Private Sub GP_MOVE_DAY(intIDOU As Integer)
    Dim intPOS As Integer
    Dim dteDate As Date
    
    Call ERASE_YEAR_MONTH                       ' 年月コンボの非表示化
    ' 移動後の位置,日付を算出
    intPOS = g_CurPos + intIDOU             ' 移動後位置
    dteDate = g_FormDate1 + intIDOU         ' 移動後日付
    If ((intPOS < 1) Or (intPOS > 42)) Then
        ' 前月又は翌月に移動
        g_FormDate1 = dteDate
        Call GP_MakeCalendar
        Exit Sub
    End If
    '---------------------------------------------------------------------------
    ' 以前の位置の日付ラベルの背景色を元に戻す
    With tblDate(g_CurPos)
        If ((g_CurPos >= g_POS_F) And (g_CurPos <= g_POS_T)) Then
            ' 当月内
            Select Case tblYobi(g_CurPos)
                Case g_intSunday:   .BackColor = cnsBC_Sunday
                Case g_intSaturday: .BackColor = cnsBC_Saturday
                Case Else: .BackColor = cnsBC_Month
            End Select
        Else
            ' 前後月
            .BackColor = cnsBC_Other
        End If
    End With
    '---------------------------------------------------------------------------
    ' 今回の位置の日付ラベルの背景色を選択状態に変更
    With tblDate(intPOS)
        .BackColor = cnsBC_Select
    End With
    ' 現在日付(退避)を更新
    g_FormDate1 = dteDate
    g_CurPos = intPOS
    LBL_GUIDE.Caption = tblGuide(g_CurPos)      ' ガイド表示
End Sub

'*******************************************************************************
' カレンダークリック処理
'*******************************************************************************
Private Sub GP_ClickCalendar(dteDate As Date)
    Call ERASE_YEAR_MONTH                       ' 年月コンボの非表示化
    Me.Tag = CLng(dteDate)  ' 現在の選択日付(シリアル値)
    Me.Hide
End Sub

'*******************************************************************************
' ｢年｣｢月｣コンボの非表示化
'*******************************************************************************
Private Sub ERASE_YEAR_MONTH()
    Call ERASE_YEAR
    Call ERASE_MONTH
End Sub

'*******************************************************************************
' ｢年｣コンボの非表示化
'*******************************************************************************
Private Sub ERASE_YEAR()
    If g_VisibleYear Then
        CBO_YEAR.Visible = False
        g_VisibleYear = False
    End If
End Sub

'*******************************************************************************
' ｢月｣コンボの非表示化
'*******************************************************************************
Private Sub ERASE_MONTH()
    If g_VisibleMonth Then
        CBO_MONTH.Visible = False
        g_VisibleMonth = False
    End If
End Sub

'--------------------------------<< End of Source >>----------------------------
