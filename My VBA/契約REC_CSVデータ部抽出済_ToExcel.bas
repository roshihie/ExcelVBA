Option Explicit
'*******************************************************************************
'        CLM0277 �t�@�C�����e  Excel �ϊ�
'*******************************************************************************
'      < �����T�v >
'        EASYPLUS �ŏo�͂����t�@�C�����e(CSV�f�[�^�����o��)�� Excel �ɕϊ�����B
'*******************************************************************************
Public Sub �_��REC_CSV�f�[�^�����o��_ToExcel()

  Const sBOOK_PASS As String = _
        "C:\Users\roshi_000\MyImp\MyOwn\Develop\Excel VBA\My VBA\�f�[�^\"
  Const sHDR_BOOK  As String = "CLM0277_�_��REC_HDR.xlsx"
  Const sHDR_SHEET As String = "�_��"
  Dim oFoundRng    As Range

  Application.ScreenUpdating = False
  Columns(2).TextToColumns Destination:=Cells(1, 2), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, Comma:=True, _
    FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), _
                     Array(6, 2), Array(7, 2), Array(8, 2), Array(9, 2), Array(10, 2), _
                     Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15, 2), _
                     Array(16, 2), Array(17, 2), Array(18, 2), Array(19, 2), Array(20, 2), _
                     Array(21, 2), Array(22, 2), Array(23, 2), Array(24, 2), Array(25, 2), _
                     Array(26, 2)), _
    TrailingMinusNumbers:=True
    
'                              �������ڂ̃J���}�ҏW
' With Range(Cells(2, 6).Offset(, 1), Cells(2, 7).Offset(, 1).End(xlDown))
'   .NumberFormatLocal = "#,##0_ "
' End With

  Call sb����_ToExcel(sBOOK_PASS, sHDR_BOOK, sHDR_SHEET)

  Set oFoundRng = Cells.Find(What:=sHDR_SHEET, LookAt:=xlPart)
  If oFoundRng Is Nothing Then
    Cells(1, 1).Activate
  Else
    oFoundRng.End(xlDown).Offset(, 1).Select
    ActiveWindow.FreezePanes = True
  End If

End Sub
    
'*******************************************************************************
'        Excel �ϊ����ʏ���
'*******************************************************************************
'      < �����T�v >
'        Excel �ϊ��ɂ����āA���ʏ������܂Ƃ߂�
'*******************************************************************************
Private Sub sb����_ToExcel(sBOOK_PASS As String, sHDR_BOOK As String, sHDR_SHEET As String)

  Dim sTrgBook As String
  Dim nRow     As Long
  Dim nCol     As Long
  Dim oTrgRect As tRect
  Dim oHDRRect As tRect
  Dim oHDRStr  As tHDRStr
  
  Set oTrgRect.oBgn = Cells(2, 2)
  nRow = Cells(ActiveSheet.Rows.Count, 2).End(xlUp).Row
  nCol = Cells(nRow, ActiveSheet.Columns.Count).End(xlToLeft).Column
  Set oTrgRect.oEnd = Cells(nRow, nCol)
  
  With Range(oTrgRect.oBgn, oTrgRect.oEnd)
    .RowHeight = 12
    .HorizontalAlignment = xlCenter
  End With
'                              �R��}�� (��\���, ����W���[���, ��e�X�g�P�[�X�ԍ����)
  Range(Columns(oTrgRect.oBgn.Column), Columns(oTrgRect.oBgn.Column).Offset(, 2)).Insert Shift:=xlToRight, _
    CopyOrigin:=xlFormatFromRightOrBelow
    
  Set oTrgRect.oBgn = oTrgRect.oBgn.Offset(, -3)
  Set oTrgRect.oEnd = Cells(oTrgRect.oEnd.Row, _
                            Cells(oTrgRect.oEnd.Row, ActiveSheet.Columns.Count).End(xlToLeft).Column)
  Range(oTrgRect.oBgn, oTrgRect.oEnd).Select
  Call sb�r���`��(Selection)

  sTrgBook = ActiveWorkbook.Name
  Workbooks.Open Filename:=sBOOK_PASS & sHDR_BOOK, ReadOnly:=True
  Worksheets(sHDR_SHEET).Activate

  Set oHDRRect.oBgn = Cells(2, 2)
  nRow = Cells(ActiveSheet.Rows.Count, 2).End(xlUp).Offset(1).Row
  nCol = Cells(nRow - 1, ActiveSheet.Columns.Count).End(xlToLeft).Column
  Set oHDRRect.oEnd = Cells(nRow, nCol)
  Range(Rows(oHDRRect.oBgn.Row), Rows(oHDRRect.oEnd.Row)).Copy
  With ActiveSheet.PageSetup
    oHDRStr.sLeftHDR = .LeftHeader
    oHDRStr.sCentHDR = .CenterHeader
    oHDRStr.sRigtHdr = .RightHeader
  End With

  Windows(sTrgBook).Activate
  With Rows(2)
    .Insert Shift:=xlDown
    .PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone
  End With
  
  With ActiveSheet.PageSetup                       '����Z�b�g�A�b�v
    .LeftHeader = oHDRStr.sLeftHDR
    .CenterHeader = oHDRStr.sCentHDR
    .RightHeader = oHDRStr.sRigtHdr
    .Orientation = xlLandscape                       '������
    .Zoom = 75                                       '�g��k����
    .CenterHorizontally = False                      '�y�[�W���� ����
    .CenterVertically = False                        '�y�[�W���� ����
  End With
  
  Application.DisplayAlerts = False
  Workbooks(sHDR_BOOK).Close
  Application.DisplayAlerts = True
  
  Range(oTrgRect.oBgn.Offset(-1), Cells(oTrgRect.oBgn.Offset(-1).Row, oTrgRect.oEnd.Column)).AutoFilter

End Sub

'*******************************************************************************
'        �r���`��
'*******************************************************************************
'      < �����T�v >
'        �w�肳�ꂽ�G���A���ȉ��̒ʂ�r���������B
'        �G���A�̈͂ݐ��C�c���F�����C
'                ����        �F�j�� �ŕ`�悷��
'*******************************************************************************
Private Sub sb�r���`��(oDrawRng As Range)

  With oDrawRng.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  With oDrawRng.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  With oDrawRng.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  With oDrawRng.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  With oDrawRng.Borders(xlInsideHorizontal)
    .LineStyle = xlDot
    .Weight = xlThin
  End With
  With oDrawRng.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With

End Sub

