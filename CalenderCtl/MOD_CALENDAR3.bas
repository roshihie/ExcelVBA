Attribute VB_Name = "MOD_CALENDAR3"
'*******************************************************************************
'   カレンダーフォーム3(日付入力部品)   ※呼び出しプロシージャ
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'*******************************************************************************
Option Explicit
Public Const cnsDateFormat = "YYYY/MM/DD"   ' デフォルトの日付Format
Private Const cnsCaption = "日付選択"       ' デフォルトのCaption
Public g_swCalendar1Loaded As Boolean       ' Load判定スイッチ

'*******************************************************************************
' ユーザーフォームのテキストボックス(MsForms.TextBox)から表示させる
'*******************************************************************************
' [引数]
' 　・テキストボックス(Object、シートからの場合はコントロールツールボックスの物)
' 　・カレンダーフォームのCaption(String) ※Option、デフォルトは"日付選択"
' 　・値を返す時のFormat(String) ※Option、デフォルトは"YYYY/MM/DD"
' 　・カレンダーフォームの表示位置：横(Long) ※Option
' 　・カレンダーフォームの表示位置：縦(Long) ※Option
'*******************************************************************************
Public Sub ShowCalendarFromTextBox2(objTextBox As MSForms.TextBox, _
                                    Optional strCaption As String, _
                                    Optional strFormat As String, _
                                    Optional lngLeft As Long, _
                                    Optional lngTop As Long)
    Dim dteDate As Date
    
    ' 元となる日付をテキストボックスから取得
    If IsDate(Trim(objTextBox.Text)) Then
        dteDate = CDate(Trim(objTextBox.Text))
    End If
    ' Caption(タイトル)指定がない場合はデフォルト("日付選択")を指定
    If strCaption = "" Then strCaption = cnsCaption
    ' 表示フォーマット指定がない場合はデフォルト("YYYY/MM/DD")を指定
    If strFormat = "" Then strFormat = cnsDateFormat
    ' カレンダーフォーム
    With FRM_CALENDAR3
        ' Tagに元日付(シリアル値)をセット
        .Tag = CLng(dteDate)
        ' Captionをセット
        .Caption = strCaption
        ' フォーム表示位置の確認
        If ((lngLeft <> 0) And (lngTop <> 0)) Then
            ' 指定がある場合はマニュアル指定
            .StartUpPosition = 0
            .Left = lngLeft
            .Top = lngTop
        Else
            ' 指定がない場合はオーナーフォームの中央
            .StartUpPosition = 1
        End If
        ' カレンダーフォームを表示
        .Show
        ' フォームがUnloadされた場合は以降の処理を無視する
        On Error Resume Next
        ' Tagの日付を確認
        If IsNumeric(.Tag) <> True Then Exit Sub
        If Err.Number <> 0 Then Exit Sub
        On Error GoTo 0
        ' Tagから選択日付を取り出してテキストボックスにセット
        dteDate = CDate(.Tag)
        objTextBox.Text = Format(dteDate, strFormat)
    End With
End Sub

'*******************************************************************************
' セル(Range)から表示させる
'*******************************************************************************
' [引数]
' 　・セル(Object) ※原則として単一セル
' 　・カレンダーフォームのCaption(String) ※Option、デフォルトは"日付選択"
' 　・カレンダーフォームの表示位置：横(Long) ※Option
' 　・カレンダーフォームの表示位置：縦(Long) ※Option
'*******************************************************************************
Public Sub ShowCalendarFromRange2(objRange As Range, _
                                  Optional strCaption As String, _
                                  Optional lngLeft As Long, _
                                  Optional lngTop As Long)
    Dim dteDate As Date

    ' 元となる日付をセルから取得
    If IsDate(Trim(objRange.Value)) Then
        dteDate = CDate(Trim(objRange.Value))
    End If
    ' Caption(タイトル)指定がない場合はデフォルト("日付選択")を指定
    If strCaption = "" Then strCaption = cnsCaption
    ' カレンダーフォーム
    With FRM_CALENDAR3
        ' Tagに元日付(シリアル値)をセット
        .Tag = CLng(dteDate)
        ' Captionをセット
        .Caption = strCaption
        ' フォーム表示位置の確認
        If ((lngLeft <> 0) And (lngTop <> 0)) Then
            ' 指定がある場合はマニュアル指定
            .StartUpPosition = 0
            .Left = lngLeft
            .Top = lngTop
        Else
            ' 指定がない場合はオーナーフォームの中央
            .StartUpPosition = 1
        End If
        ' カレンダーフォームを表示
        .Show
        ' フォームがUnloadされた場合は以降の処理を無視する
        On Error Resume Next
        ' Tagの日付を確認
        If IsNumeric(.Tag) <> True Then Exit Sub
        If Err.Number <> 0 Then Exit Sub
        On Error GoTo 0
        ' Tagから選択日付を取り出してセルにセット
        dteDate = CDate(.Tag)
        objRange.Value = dteDate
    End With
End Sub

'--------------------------------<< End of Source >>----------------------------



