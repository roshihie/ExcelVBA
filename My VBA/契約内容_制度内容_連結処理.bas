Option Explicit
'*******************************************************************************
'        �_����e�ꗗ  ���x���e�ꗗ �A������
'*******************************************************************************
'      < �����T�v >
'        ��ƂȂ���e�ꗗ�ɕʂȓ��e�ꗗ���L�[�w��ŘA������B
'*******************************************************************************
Public Sub �_����e_���x���e_�A��()
  
  Const s���x�ԍ� As String = "���x�ԍ�"
  
  Dim oTrgRect  As tRect
  Dim oDataTop  As tRect
  Dim oFoundRng As Range
  Dim nProc     As Long
  Dim o���x�ԍ�Rngs As Range
  Dim o���x�ԍ�Rng  As Range

  Application.ScreenUpdating = False
  
  Set oFoundRng = Cells.Find(What:=s���x�ԍ�, LookAt:=xlPart)
  If oFoundRng Is Nothing Then
     Exit Sub
  End If
  
  nProc = 0
  Set oTrgRect.oBgn = Cells(oFoundRng.Row, ActiveSheet.Columns.Count).End(xlToLeft).Offset(, 1)
  Call sb���x�f�[�^�擾(nProc, oFoundRng, oTrgRect)
  Set oTrgRect.oEnd = Cells(oTrgRect.oBgn.Row, ActiveSheet.Columns.Count).End(xlToLeft).Offset(1)
  
  nProc = 1
  Set o���x�ԍ�Rngs = Range(oFoundRng.End(xlDown), Cells(ActiveSheet.Rows.Count, oFoundRng.Column).End(xlUp))
  Set oDataTop.oBgn = Cells(oFoundRng.End(xlDown).Row, oTrgRect.oBgn.Column)
  Set oDataTop.oEnd = Cells(oFoundRng.End(xlDown).Row, oTrgRect.oEnd.Column)
  
  For Each o���x�ԍ�Rng In o���x�ԍ�Rngs
  
    Set oTrgRect.oBgn = Cells(o���x�ԍ�Rng.Row, oTrgRect.oBgn.Column)
    Set oTrgRect.oEnd = Cells(o���x�ԍ�Rng.Row, oTrgRect.oEnd.Column)
    
    If o���x�ԍ�Rng.Value <> o���x�ԍ�Rng.Offset(-1).Value Then
      Call sb���x�f�[�^�擾(nProc, o���x�ԍ�Rng, oTrgRect)
      
      If o���x�ԍ�Rng.Offset(-1).Value <> "" Then
        Range(oDataTop.oBgn, oTrgRect.oEnd.Offset(-1)).Select
        Call sb�r���`��(Selection)
      End If
    End If
    
  Next

  Range(oDataTop.oBgn, oTrgRect.oEnd).Select
  Call sb�r���`��(Selection)
  With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  
  nProc = 2
  Call sb���x�f�[�^�擾(nProc, oFoundRng, oTrgRect)
  
  ActiveSheet.AutoFilterMode = False
  Range(Cells(oDataTop.oBgn.Offset(-1).Row, 2), oDataTop.oEnd.Offset(-1)).AutoFilter
  Application.ScreenUpdating = True
  Cells(1, 1).Activate
     
End Sub

'*******************************************************************************
'        ���x�f�[�^�擾����
'*******************************************************************************
'      < �����T�v >
'        �����敪�ɉ����āA���x���e�ꗗ�̃w�b�_�[����f�[�^�����擾����B
'        �����敪�F0  �w�b�_�[�����擾
'        �����敪�F1  �f�[�^�����擾
'*******************************************************************************
Private Sub sb���x�f�[�^�擾(nProc As Long, o���x�ԍ�Rng As Range, oTrgRect As tRect)

  Const s���xBOOK_PASS As String = _
        "C:\Users\roshi_000\MyImp\MyOwn\Develop\Excel VBA\My VBA\�f�[�^\"
  ' ��L�� " ���G�X�P�[�v�V�[�P���X�ɂȂ� gVim�ŊJ���ƈȍ~���R�����g�ɂȂ邽�ߓ��s�ǉ�
  Const s���xBOOK  As String = "���x���e�ꗗ.xlsx"
  Const s���xSHEET As String = "���x"

  Static stsTrgBook As String
  Static stn���xCol As Long
  Static stoSocRect As tRect
  Dim oFoundRng As Range
  Dim nRow As Long
  Dim nCol As Long
  
  Select Case True
  Case nProc = 0               '��������
    stsTrgBook = ActiveWorkbook.Name
    Workbooks.Open Filename:=s���xBOOK_PASS & s���xBOOK, ReadOnly:=True
    Worksheets(s���xSHEET).Activate
    Set oFoundRng = Cells.Find(What:=o���x�ԍ�Rng.Value, LookAt:=xlPart)
    If oFoundRng Is Nothing Then
      Exit Sub
    End If
    
    stn���xCol = oFoundRng.Column
    Set stoSocRect.oBgn = oFoundRng.Offset(, 1)
    nRow = oFoundRng.End(xlDown).Offset(-1).Row
    nCol = Cells(stoSocRect.oBgn.Row, ActiveSheet.Columns.Count).End(xlToLeft).Column
    Set stoSocRect.oEnd = Cells(nRow, nCol)
  
    Range(stoSocRect.oBgn, stoSocRect.oEnd).Copy
    
    Windows(stsTrgBook).Activate
        
    With oTrgRect.oBgn
      .PasteSpecial Paste:=xlPasteColumnWidths
      .PasteSpecial Paste:=xlPasteAll
    End With
    Application.CutCopyMode = False
    
  Case nProc = 1               '�f�[�^����
    Windows(s���xBOOK).Activate
    Set oFoundRng = Columns(stn���xCol).Find(What:=o���x�ԍ�Rng.Value, LookAt:=xlWhole)
    If oFoundRng Is Nothing Then
      Exit Sub
    End If
    
    Set stoSocRect.oBgn = Cells(oFoundRng.Row, stoSocRect.oBgn.Column)
    Set stoSocRect.oEnd = Cells(oFoundRng.Row, stoSocRect.oEnd.Column)
  
    Range(stoSocRect.oBgn, stoSocRect.oEnd).Copy
    
    Windows(stsTrgBook).Activate
    With oTrgRect.oBgn
      .PasteSpecial Paste:=xlPasteColumnWidths
      .PasteSpecial Paste:=xlPasteAllExceptBorders
    End With
    Application.CutCopyMode = False
    
  Case nProc = 2               '�I������
    Windows(s���xBOOK).Activate
    Application.DisplayAlerts = False
    Workbooks(s���xBOOK).Close
    Application.DisplayAlerts = True
    Windows(stsTrgBook).Activate
  End Select
  
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

  With oDrawRng
    With .Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .Weight = xlThin
    End With
    With .Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .Weight = xlThin
    End With
    With .Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .Weight = xlThin
    End With
    With .Borders(xlEdgeBottom)
      .LineStyle = xlDot
      .Weight = xlThin
    End With
  End With

End Sub

