 '
   ' Record1 �}�N��
   ' ORG_SHEET_NAME,NEW_SHEET_NAME�̂Q�V�[�g���A�Z���P�ʂŔ�r
   ' ���ق��������Z����F�t������
   '
   Private Const ORG_SHEET_NAME = "�C����"
   Private Const NEW_SHEET_NAME = "tmp9"
   Private Const MY_RANGE = "A1:U69"
   
   Sub Record1()
       Dim x As Integer, y As Integer
           
       For Each c In Sheets(ORG_SHEET_NAME).Range(MY_RANGE)
           x = c.Column
           y = c.Row
               
           If (Sheets(ORG_SHEET_NAME).Cells(y, x) <> Sheets(NEW_SHEET_NAME).Cells(y, x)) Then
               Sheets(ORG_SHEET_NAME).Cells(y, x).Interior.ColorIndex = 6
               Sheets(NEW_SHEET_NAME).Cells(y, x).Interior.ColorIndex = 6
               '.Pattern = xlSolid
           Else
               Sheets(ORG_SHEET_NAME).Cells(y, x).Interior.ColorIndex = 2
               Sheets(NEW_SHEET_NAME).Cells(y, x).Interior.ColorIndex = 2
           End If
       Next
   End Sub
