Option Explicit
'*******************************************************************************
'        ソース解析シート  新旧コンペア＆新シートへ反映
'*******************************************************************************
'      < 処理概要 >
'        ソース解析自由帳のCOBOLソースの新旧シートをコンペアし、新シート上に
'        旧シートのコメントを反映する。
'*******************************************************************************
Private Type tSocStrct                 'ソース構造体
  bMatch     As Boolean                  'マッチング結果
  nOpidx     As Long                     '相手側index
  sSeqno     As String                   'ソースSEQNO
  sSourc     As String                   'ソースコード
  sComnt(46) As String                   'コメント(C列～AV列)
End Type

Private Type tSocPropt                 'ソース属性
  oSheet     As Worksheet                'シート
  nSocMinRow As Long                     'ソース最小行
  nSocMaxRow As Long                     'ソース最大行
  nSocMaxidx As Long                     'ソース最大index
End Type

Private Type tSocCompr                 'コンペアindex
  nLimBegin  As Long                     '制限idx Begin
  nLimEnd    As Long                     '制限idx End
  nCmpBegin  As Long                     'コンペアidx Begin
  nCmpEnd    As Long                     'コンペアidx End
End Type

Private Type tSocBody                  'ソースエリア
  oSoc()     As tSocStrct                'ソースシート
  oPrp       As tSocPropt                'ソース属性
  oidx       As tSocCompr                'コンペアindex
End Type

Const g_cnSocCol       As Long = 2     'ソース列
Const g_cnLeftComntCol As Long = 3     'コメント列(左端)
Const g_cnRigtComntCol As Long = 48    'コメント列(右端)
Const g_cnSocMinRow    As Long = 3     'ソース最小行

Dim g_oFso   As New FileSystemObject   '##### Debug
Dim g_oTStrm As Object                 '##### Debug

'*******************************************************************************
'        メイン処理
'*******************************************************************************
Public Sub COBOL解析シートコンペア()
  
  Const cnOutFile As String = "C:\Users\roshi_000\MyImp\MyOwn\Develop\Excel VBA\My VBA\Debug.txt"  '##### Debug

  Const cnStartBlk As Long = 45        'マッチングブロック初期値
  Const cnStep     As Long = -1        '減算行
  
  Dim oNew    As tSocBody              '新ソース
  Dim oOld    As tSocBody              '旧ソース
  Dim nBlock  As Long                  'コンペアブロック
  
  Call sb初期処理(oNew, oOld)
  
  Set g_oTStrm = g_oFso.CreateTextFile(cnOutFile, True)  '##### Debug
  
  For nBlock = cnStartBlk To 1 Step cnStep
    Call sbマッチング処理(nBlock, oNew, oOld)
  Next
  
  Call sb処理後Debug(oNew, oOld)                         '##### Debug
 'Call sb新シート編集(oNew, oOld)
 'Call sb終了処理

  g_oTStrm.Close                                         '##### Debug

End Sub

'*******************************************************************************
'        初期処理
'*******************************************************************************
Private Sub sb初期処理(oNew As tSocBody, oOld As tSocBody)
  
  Dim sNewSheet  As String
  Dim sOldSheet  As String
  
  sNewSheet = Application.InputBox(Prompt:="新ソート名を入力して下さい。", _
                                   Title:="新ソース名指定", Type:=2)
  sOldSheet = Application.InputBox(Prompt:="旧ソート名を入力して下さい。", _
                                   Title:="旧ソース名指定", Type:=2)
                                   
  Call sbソースエリア設定(oNew, sNewSheet)
  Call sbソースエリア設定(oOld, sOldSheet)
  
  oNew.oSoc(oNew.oPrp.nSocMaxidx).nOpidx = oOld.oPrp.nSocMaxidx
  oOld.oSoc(oOld.oPrp.nSocMaxidx).nOpidx = oNew.oPrp.nSocMaxidx
  
End Sub

'*******************************************************************************
'        ソースエリア設定
'*******************************************************************************
'      < 処理概要 >
'        ソース属性，ソース構造体を設定する。
'*******************************************************************************
Private Sub sbソースエリア設定(oCmn As tSocBody, sCmnSheet As String)
  
  Const cnLenSeq As Long = 6
  Const cnBgnSoc As Long = 7
  Const cnLenSoc As Long = 66
  
  Dim nRow       As Long
  Dim nidx       As Long
  Dim nComntCol  As Long
  Dim nComntidx  As Long

  With oCmn.oPrp
    Set .oSheet = Worksheets(sCmnSheet)
    .nSocMinRow = g_cnSocMinRow
    .nSocMaxRow = .oSheet.Cells(.oSheet.Rows.Count, g_cnSocCol).End(xlUp).Row
    .nSocMaxidx = .nSocMaxRow - g_cnSocMinRow
    ReDim oCmn.oSoc(.nSocMaxidx)
  End With

  For nRow = oCmn.oPrp.nSocMinRow To oCmn.oPrp.nSocMaxRow
    nidx = nRow - g_cnSocMinRow
    With oCmn.oSoc(nidx)
      .bMatch = False
      .nOpidx = 0
      .sSeqno = Left(oCmn.oPrp.oSheet.Cells(nRow, g_cnSocCol).Value, cnLenSeq)
      .sSourc = Trim(Mid(oCmn.oPrp.oSheet.Cells(nRow, g_cnSocCol).Value, cnBgnSoc, cnLenSoc))
      
      For nComntCol = g_cnLeftComntCol To g_cnRigtComntCol
        nComntidx = nComntCol - g_cnLeftComntCol
        .sComnt(nComntidx) = oCmn.oPrp.oSheet.Cells(nRow, nComntCol).Value
      Next
    End With
  Next
  
End Sub

'*******************************************************************************
'        マッチング処理
'*******************************************************************************
Private Sub sbマッチング処理(nBlock As Long, _
                             oNew As tSocBody, oOld As tSocBody)
  Dim nNewRept  As Long
  Dim nOldRept  As Long
  Dim nMchCond  As Long

  oOld.oidx.nLimBegin = 0                        'Old制約idx 設定
  oOld.oidx.nLimEnd = oOld.oPrp.nSocMaxidx

  oOld.oidx.nCmpBegin = oOld.oidx.nLimBegin      'Oldコンペアidx 設定
  nOldRept = fnコンペアidx確定(nBlock, oOld)

  Do While (nOldRept = 0)                        'Oldコンペアidx Begin 取得可能の間 行う

    Call sb新ソース制約idx確定(oOld, oNew)         'New制約idx 設定
    oNew.oidx.nCmpBegin = oNew.oidx.nLimBegin      'Newコンペアidx 設定
    nNewRept = fnコンペアidx確定(nBlock, oNew)
      
    nMchCond = 9                                   'コンペア結果 未実施

    Do While (nNewRept = 0)                        'Newコンペアidx Begin, End 取得済の間 行う
 
      nMchCond = 1                                   'コンペア結果 アンマッチ
      nMchCond = fnコンペア実施(oNew, oOld)          'コンペア実施
        
      If nMchCond = 0 Then                           'コンペア結果 マッチ
        Call sbマッチ時処理(oNew, oOld)
        Call sbマッチDebug(nBlock, oNew, oOld)         '##### Debug #####
        Exit Do
      End If

      oNew.oidx.nCmpBegin = oNew.oidx.nCmpBegin + 1  'Newコンペアidx 設定
      nNewRept = fnコンペアidx確定(nBlock, oNew)
    Loop

    If (nMchCond = 0) Then                       'コンペア結果 マッチ
      oOld.oidx.nCmpBegin = oOld.oidx.nCmpBegin + nBlock
      
    ElseIf (nMchCond = 1) Then                   'コンペア結果 アンマッチ
      oOld.oidx.nCmpBegin = oOld.oidx.nCmpBegin + 1
       
    ElseIf (nMchCond = 9) Then                   'コンペア結果 未実施
      oOld.oidx.nCmpBegin = oOld.oidx.nCmpBegin + nBlock
        
    End If

    nOldRept = fnコンペアidx確定(nBlock, oOld)   'Oldコンペアidx 設定
  Loop
    
End Sub

'*******************************************************************************
'        コンペアidx 確定
'*******************************************************************************
'      < 処理概要 >
'        引数にブロック行，ソース本体を受け取り、コンペアidx Begin，End を
'        取得する。
'        コンペアidx Begin が取得不能の場合      戻り値に 99
'        コンペアidx End   が制約idxを超えた場合 戻り値に 90
'        コンペア範囲にマッチ確定行がある場合    戻り値に 10 (上位に返さない)
'        コンペアidx Begin, End が取得された場合 戻り値に 00 を返す。
'*******************************************************************************
Private Function fnコンペアidx確定(nBlock As Long, oCmn As tSocBody) As Long

  Dim nCmnRept  As Long
  Dim nNextEnd  As Long

  nNextEnd = oCmn.oidx.nCmpBegin
  nCmnRept = 10
  
  Do While (nCmnRept = 10)
                                       'コンペアidx Begin 取得
    oCmn.oidx.nCmpBegin = nNextEnd
    If (fnコンペアidx_Begin確定(oCmn) = False) Then
      fnコンペアidx確定 = 99
      Exit Function
    End If
                                       'コンペアidx End   取得
    oCmn.oidx.nCmpEnd = oCmn.oidx.nCmpBegin + nBlock - 1
    nCmnRept = fnコンペア範囲行チェック(nBlock, oCmn)
    If nCmnRept = 90 Then
      fnコンペアidx確定 = nCmnRept
      Exit Function
    End If
    nNextEnd = oCmn.oidx.nCmpEnd + 1
  Loop
  
  fnコンペアidx確定 = 0

End Function

'*******************************************************************************
'        コンペアidx Begin 確定
'*******************************************************************************
'      < 処理概要 >
'        引数にソース本体を受け取り、ソース本体を検索して、
'        マッチ未確定の最初行を取得可能の場合、
'            結果をコンペアidx Beginに設定し、戻り値に True を返す。
'        マッチ未確定行が取得不可の場合、
'            戻り値に False を返す。
'*******************************************************************************
Private Function fnコンペアidx_Begin確定(oCmn As tSocBody) As Boolean

  Dim nidx  As Long
  Dim nEnd  As Long

  fnコンペアidx_Begin確定 = False
  nidx = oCmn.oidx.nCmpBegin
  
  Do While (nidx <= oCmn.oidx.nLimEnd)
    If oCmn.oSoc(nidx).bMatch = False Then
      oCmn.oidx.nCmpBegin = nidx
      fnコンペアidx_Begin確定 = True
      Exit Function
    End If
    nidx = nidx + 1
  Loop

End Function

'*******************************************************************************
'        コンペア範囲行チェック
'*******************************************************************************
'      < 処理概要 >
'        引数にブロック行，ソース本体を受け取り、コンペアidx Begin＋１行から
'        コンペアidx Begin＋１行からコンペアidx Begin＋ブロック行まで、すべて
'        マッチ未確定行であるチェックを行う。
'        制限idx End＜コンペアidx End の場合、戻り値に 90
'        コンペア範囲にマッチ確定行がある場合 戻り値に 10
'        それ以外の場合                       戻り値に 00 を返す。
'*******************************************************************************
Private Function fnコンペア範囲行チェック(nBlock As Long, oCmn As tSocBody) As Long

  Dim nidx  As Long
  
  If oCmn.oidx.nLimEnd < oCmn.oidx.nCmpEnd Then
    fnコンペア範囲行チェック = 90
    Exit Function
  End If
  
  nidx = oCmn.oidx.nCmpBegin + 1
  Do While (nidx <= oCmn.oidx.nCmpEnd)
    If oCmn.oSoc(nidx).bMatch = True Then
      fnコンペア範囲行チェック = 10
      Exit Function
    End If
    nidx = nidx + 1
  Loop

  fnコンペア範囲行チェック = 0

End Function

'*******************************************************************************
'        新ソース制約 index 確定
'*******************************************************************************
'      < 処理概要 >
'        引数に旧ソース，新ソースを受け取り、
'      ・旧ソースコンペアidx End から下方に旧ソースを検索して、マッチ確定済行の
'        最初行を取得し、それに該当する相手(新ソース)の index から
'        新ソースのマッチ未確定行を新ソース制約idx End とする。
'      ・旧ソースコンペアidx Begin から上方 旧ソースを検索して、マッチ確定済行の
'        最初行を取得し、それに該当する相手(新ソース)の index から
'        新ソースのマッチ未確定行を新ソース制約idx Begin とする。
'*******************************************************************************
Private Sub sb新ソース制約idx確定(oOld As tSocBody, oNew As tSocBody)

  Dim nidx  As Long
  Dim nEnd  As Long
                                       '新ソース制約idx End   取得
  nEnd = oOld.oPrp.nSocMaxidx
  nidx = oOld.oidx.nCmpEnd
  oNew.oidx.nLimEnd = oNew.oPrp.nSocMaxidx
  
  Do While (nidx <= nEnd)
    If oOld.oSoc(nidx).bMatch = True Then
      oNew.oidx.nLimEnd = oOld.oSoc(nidx).nOpidx - 1
      Exit Do
    End If
    nidx = nidx + 1
  Loop
                                       '新ソース制約idx Begin 取得
  nEnd = 0
  nidx = oOld.oidx.nCmpBegin
  oNew.oidx.nLimBegin = 0
  
  Do While (nidx >= nEnd)
    If oOld.oSoc(nidx).bMatch = True Then
      oNew.oidx.nLimBegin = oOld.oSoc(nidx).nOpidx + 1
      Exit Do
    End If
    nidx = nidx - 1
  Loop

End Sub

'*******************************************************************************
'        コンペア実施
'*******************************************************************************
'      < 処理概要 >
'        新旧ソースのコンペア範囲index Begin～End のコンペアを行う。
'        コンペアブロック単位でマッチしている場合 0 を返す。
'        アンマッチの場合                         1 を返す。
'*******************************************************************************
Private Function fnコンペア実施(oNew As tSocBody, oOld As tSocBody) As Long

  Dim nNewidx  As Long
  Dim nOldidx  As Long

  nNewidx = oNew.oidx.nCmpBegin
  nOldidx = oOld.oidx.nCmpBegin
  
  Do While (nNewidx <= oNew.oidx.nCmpEnd)
    If oNew.oSoc(nNewidx).sSourc <> oOld.oSoc(nOldidx).sSourc Then
      fnコンペア実施 = 1
      Exit Function
    End If
    nNewidx = nNewidx + 1
    nOldidx = nOldidx + 1
  Loop

  fnコンペア実施 = 0

End Function

'*******************************************************************************
'        マッチ時処理
'*******************************************************************************
'      < 処理概要 >
'        コンペアブロック単位でマッチした場合、コンペア結果に True, 相手idx に
'        旧ソース，新ソースそれぞれ相手のマッチした index を設定する。
'*******************************************************************************
Private Sub sbマッチ時処理(oNew As tSocBody, oOld As tSocBody)

  Dim nNewidx  As Long
  Dim nOldidx  As Long

  nNewidx = oNew.oidx.nCmpBegin
  nOldidx = oOld.oidx.nCmpBegin

  Do While (nNewidx <= oNew.oidx.nCmpEnd)
    oNew.oSoc(nNewidx).bMatch = True
    oNew.oSoc(nNewidx).nOpidx = nOldidx
    oOld.oSoc(nOldidx).bMatch = True
    oOld.oSoc(nOldidx).nOpidx = nNewidx
    
    nNewidx = nNewidx + 1
    nOldidx = nOldidx + 1
  Loop

End Sub

'*******************************************************************************
'        マッチDebug 処理
'*******************************************************************************
Private Sub sbマッチDebug(nBlock As Long, oNew As tSocBody, oOld As tSocBody)

  g_oTStrm.WriteLine "########### マッチ情報 ###########"
  g_oTStrm.WriteLine "  ●ブロック行数        : " & nBlock
  g_oTStrm.WriteLine "    新ソース   CmpBegin : " & oNew.oidx.nCmpBegin
  g_oTStrm.WriteLine "               CmpEnd   : " & oNew.oidx.nCmpEnd
  g_oTStrm.WriteLine " "
  g_oTStrm.WriteLine "    旧ソース   CmpBegin : " & oOld.oidx.nCmpBegin
  g_oTStrm.WriteLine "               CmpEnd   : " & oOld.oidx.nCmpEnd
  
End Sub
  
'*******************************************************************************
'        処理後Debug 処理
'*******************************************************************************
Private Sub sb処理後Debug(oNew As tSocBody, oOld As tSocBody)

  Dim nidx  As Long
  
  g_oTStrm.WriteLine " "
  g_oTStrm.WriteLine " "
  g_oTStrm.WriteLine "####### New Source ############################################################"
  For nidx = 0 To oNew.oPrp.nSocMaxidx
    With oNew.oSoc(nidx)
      g_oTStrm.WriteLine .sSeqno & " " & .sSourc & " " & Space(75 - Len(.sSourc)) & .bMatch
    End With
  Next
  
  g_oTStrm.WriteLine " "
  g_oTStrm.WriteLine " "
  g_oTStrm.WriteLine "####### Old Source ############################################################"
  For nidx = 0 To oOld.oPrp.nSocMaxidx
    With oOld.oSoc(nidx)
      g_oTStrm.WriteLine .sSeqno & " " & .sSourc & " " & Space(75 - Len(.sSourc)) & .bMatch
    End With
  Next

End Sub

'*******************************************************************************
'        新シート編集処理
'*******************************************************************************
Private Sub sb新シート編集(oNew As tSocBody, oOld As tSocBody)

  Dim nNewidx  As Long
  Dim nOldidx  As Long
  Dim nidx     As Long
  
  oOld.oidx.nLimBegin = 0                        'Old制約idx 設定
  oOld.oidx.nLimEnd = oOld.oPrp.nSocMaxidx
  oNew.oidx.nLimBegin = 0                        'New制約idx 設定
  oNew.oidx.nLimEnd = oNew.oPrp.nSocMaxidx
  
  nNewidx = 0
  nOldidx = 0
  Do While (nNewidx <= oNew.oidx.nLimEnd)

    oNew.oidx.nCmpBegin = fnマッチ行取得(oNew, nNewidx)
    nidx = oNew.oidx.nCmpBegin
    oOld.oidx.nCmpBegin = oOld.oSoc(nidx).nOpidx
    
    For nidx = nOldidx To oOld.oidx.nCmpBegin - 1
      'アンマッチ Oldソース 出力
    Next
    
    For nidx = nNewidx To oNew.oidx.nCmpBegin - 1
      'アンマッチ Newソース 出力
    Next
    
    nNewidx = oNew.oidx.nCmpBegin
    Do While (nidx <= oNew.oidx.nLimEnd And _
              oNew.oSoc(nidx).bMatch = False)
        Exit Do
      End
      'マッチ Old ソース 出力
      
      nNewidx = nNewidx + 1
    Loop
    
  Loop

End Sub
