Option Explicit
'*******************************************************************************
'        ���Z���͈͂ւ̓\�t����
'*******************************************************************************
'      < �����T�v >
'        �N���b�v�{�[�h�փR�s�[�����f�[�^�����Z���͈͂֓\��t����B
'        �i�t�B���^�����O�����͈͂ւ̓\��t���j
'*******************************************************************************
Public Sub ���Z���ւ̓\�t����()

  Dim oClipb   As DataObject
  Dim oTrgRect As tRect
  Dim vClipbStr   As Variant
  Dim asCellStr() As String
  Dim bGetTopRng  As Boolean
  
  Dim nGap     As Long
  Dim oSocRng  As Range
  Dim oDstRng  As Range
  Dim oTopRng  As Range
  Dim oDst     As Range
  Dim i        As Long
  Dim j        As Long
  
  On Error GoTo InputCancel
  Set oSocRng = Application.InputBox(Prompt:="�R�s�[����G���A���w�肵�ĉ������B", _
                                     Title:="�R�s�[�G���A�w��", _
                                     Type:=8)
  Set oDstRng = ActiveCell.Worksheet.AutoFilter.Range
  Set oDstRng = oDstRng.Resize(, oDstRng.Columns.Count + 1)
  Set oDstRng = Intersect(oDstRng, oDstRng.Offset(1))
  
  bGetTopRng = True
  Do While (bGetTopRng = True)
    Set oTopRng = Application.InputBox(Prompt:="�\��t������G���A�̐擪�Z�����w�肵�ĉ������B", _
                                       Title:="�\��t���G���A�擪�w��", _
                                       Type:=8)
    If Intersect(oDstRng, oTopRng) Is Nothing Then
      MsgBox ("�\��t������G���A�̐擪�Z���𐳂����w�肵�ĉ������B")
    Else
      bGetTopRng = False
    End If
  Loop
                                      
  nGap = oDstRng.Columns(1).Column - 1           '�I�[�g�t�B���^�͈͂̐擪��̍���
  Set oTrgRect.oBgn = oTopRng                    ' �\�t�擪�Z���̎擾
                                                 
  With oDstRng.Columns(oTopRng.Column - nGap)    ' �I�[�g�t�B���^�͈͓��ł̓\�t�擪�Z���Ɠ���̍ŏI�Z���̎擾
    Set oTrgRect.oEnd = .Cells(.Cells.Count)
  End With
                                                 ' �I�[�g�t�B���^�͈͂�\�t�擪�Z���`�ŏI��(����)�͈̔͂Ɍ���
  Set oDstRng = Intersect(oDstRng, Range(oTrgRect.oBgn, oTrgRect.oEnd))
                                                 ' ���Z�����擾
  Set oDstRng = oDstRng.SpecialCells(xlCellTypeVisible)
  
  oSocRng.Copy
  
  Set oClipb = New DataObject
  With oClipb
    .GetFromClipboard
    On Error Resume Next
    vClipbStr = .GetText
    On Error GoTo 0
  End With

  If Not IsEmpty(vClipbStr) Then
    vClipbStr = Split(CStr(vClipbStr), vbCrLf)
    i = 0
    For Each oDst In oDstRng.Cells
      If i > UBound(vClipbStr) Then Exit For
      asCellStr = Split(vClipbStr(i), vbTab)
      For j = 0 To UBound(asCellStr)
        oDst.Offset(, j).Value = asCellStr(j)
      Next
      
      i = i + 1
    Next
  End If
  
  Set oClipb = Nothing

InputCancel:
End Sub
