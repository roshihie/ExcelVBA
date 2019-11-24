Option Explicit
'*******************************************************************************
'        �G�N�Z���V�[�g����������
'*******************************************************************************
'      < �����T�v >
'        �V�[�g�̏���������������B
'*******************************************************************************
Public Sub �V�[�g����������()

  Dim nSize  As Long
  
  nSize = Application.InputBox(Prompt:="�t�H���g�T�C�Y����͂���", _
                               Title:="�l�r ���� �t�H���g�T�C�Y�w��", _
                               Default:=10, _
                               Type:=1)
  If nSize = 0 Then
    Exit Sub
  End If
  
  Application.ScreenUpdating = False

  With Cells
    .Font.Name = "�l�r ����"
    .Font.Size = nSize
    .ColumnWidth = 2
    .VerticalAlignment = xlCenter
  End With
  
  Columns(1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
  Columns(1).ColumnWidth = 0.3
  Rows(1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
  Rows(1).RowHeight = 3
  
  With ActiveSheet.PageSetup
    .LeftHeader = "&""�l�r ����,����""&12&U &A "
    .CenterHeader = "&""�l�r ����,����""&12 "
    .CenterFooter = "&""�l�r ����,�W��""�\ &P�^&N �\"
    .HeaderMargin = Application.CentimetersToPoints(0.9)
    .TopMargin = Application.CentimetersToPoints(1.9)
    .LeftMargin = Application.CentimetersToPoints(1.2)
    .RightMargin = Application.CentimetersToPoints(0.5)
    .BottomMargin = Application.CentimetersToPoints(1.2)
    .FooterMargin = Application.CentimetersToPoints(0.7)
  End With
  
  Application.ScreenUpdating = True
  Cells(1, 1).Activate

End Sub
