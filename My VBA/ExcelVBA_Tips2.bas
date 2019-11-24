Attribute VB_Name = "Module3"
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'        �d���������u�a�`  �o������������
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'*******************************************************************************
'        �R�s�[�V�[�g�C�I���W�i���V�[�g����s  �Ԋ|����
'*******************************************************************************
'        �����T�v�F�I���W�i���V�[�g�̊e�s�̃X�e�[�^�X��'�K�v*'�̂Ƃ�
'                  �R�s�[�V�[�g�̊e�s���r���č��ڗ�̂Ȃ��ɓ���̌��o����
'                  ���݂����ꍇ�� �R�s�[�V�[�g�̊Y���s�ɖԊ|�����s��
'*******************************************************************************
Sub ListShadowCreate1()

    Dim orgSheet     As Worksheet
    Dim cpySheet     As Worksheet
    Dim cpyRange     As Range

    Dim orgIdxRow    As Integer
    Dim orgStartRow  As Integer
    Dim orgStatusCol As Integer
    Dim orgNoCol     As Integer
    Dim orgNameCol   As Integer    '   ��r�Ώۂ̌��o����

    Dim cpyIdxRow    As Integer
    Dim cpyStartRow  As Integer
    Dim cpyNoCol     As Integer
    Dim cpyNameCol   As Integer    '   ��r�Ώۂ̌��o����

    Set orgSheet = Workbooks("Book1").Worksheets("Sheet1")
    Set cpySheet = Workbooks("Book2").Worksheets("Sheet1")

    orgStartRow = 5
    orgStatusCol = 10
    orgNoCol = 1
    orgNameCol = 2

    cpyStartRow = 3
    cpyNoCol = 1
    cpyNameCol = 2

    orgIdxRow = orgStartRow
    '   �I���W�i���V�[�g�s�Q�ƃ��[�v
    Do While orgSheet.Cells(orgIdxRow, orgNoCol).Value <> ""
        If orgSheet.Cells(orgIdxRow, orgStatusCol).Value Like "�K�v*" Then
            cpyIdxRow = cpyStartRow
            '   �R�s�[�V�[�g�s�Q�ƃ��[�v
            Do While cpySheet.Cells(cpyIdxRow, cpyNoCol).Value <> ""
                '   �R�s�[�V�[�g�̌��o����̐擪����+10��܂ł��r����
                cpySheet.Activate
                For Each cpyRange In cpySheet.Range(Cells(cpyIdxRow, cpyNameCol), _
                                                    Cells(cpyIdxRow, cpyNameCol + 10))
                    If cpyRange.Value = orgSheet.Cells(orgIdxRow, orgNameCol).Value Then
                        cpySheet.Activate
                        Range(Cells(cpyIdxRow, cpyNameCol), Cells(cpyIdxRow, 12)).Select
                        Selection.Interior.ColorIndex = 15
                    End If
                Next
                cpyIdxRow = cpyIdxRow + 1
            Loop
        End If
        orgIdxRow = orgIdxRow + 1
    Loop

End Sub

'*******************************************************************************
'        �X�e�[�^�X�ɂ��Ԋ|����
'*******************************************************************************
'        �����T�v�F�V�[�g�̊e�s�̃X�e�[�^�X��'����'�̂Ƃ��A�O���[�Ԋ|��
'                  �V�[�g�̊e�s�̃X�e�[�^�X��'*������'�̂Ƃ��A���F�Ԋ|��
'                  ���s��
'*******************************************************************************
Sub ListShadowCreate2()

    Dim ixRow       As Integer
    Dim startRow    As Integer
    Dim noCol       As Integer
    Dim startCol    As Integer
    Dim endCol      As Integer
    Dim statusCol   As Integer
    
    Dim cellValue   As String

    Workbooks("Book1").Worksheets("Sheet1").Activate
    
    startRow = 5
    noCol = 1
    startCol = 1
    statusCol = 13
    
    ixRow = startRow
    '   �V�[�g�̍s�Q�ƃ��[�v
    Do While Cells(ixRow, noCol).Value <> ""
        cellValue = Cells(ixRow, statusCol).Value
        If cellValue = "����" Then
            '   �V�[�g�̍ŉE�[����A�N�e�B�u�Z���̈�̍ŉE�[�{�Q����擾���đI��
            endCol = ActiveSheet.Columns.Count
            Range(Cells(ixRow, startCol), Cells(ixRow, endCol).End(xlToLeft).Offset(, 2)).Select
            With Selection.Interior
                .ColorIndex = 15
                .Pattern = xlSolid
            End With
        Else
            If cellValue Like "*������" Then
                '   �V�[�g�̍ŉE�[����A�N�e�B�u�Z���̈�̍ŉE�[�{�Q����擾���đI��
                endCol = ActiveSheet.Columns.Count
                Range(Cells(ixRow, startCol), Cells(ixRow, endCol).End(xlToLeft).Offset(, 2)).Select
                With Selection.Interior
                    .ColorIndex = 27
                    .Pattern = xlSolid
                End With
            End If
        End If
        ixRow = ixRow + 1
    Loop

End Sub

'*******************************************************************************
'        ���o�����ڃR�s�[����
'*******************************************************************************
'        �����T�v�F�I���W�i���V�[�g�̊e�s�̃X�e�[�^�X��"Y"�̂Ƃ�
'                  ���o�����ڂ��R�s�[�V�[�g�ɃR�s�[����
'*******************************************************************************
Sub ListItemCopy()

'
'
    Dim orgSheet    As Worksheet
    Dim cpySheet    As Worksheet
    
    Dim ixRow       As Integer         '�I���W�i���V�[�g�̍���
    Dim startRow    As Integer
    Dim noCol       As Integer
    Dim targetCol   As Integer
    Dim statusCol   As Integer

    Dim selRow      As Integer         '�R�s�[�V�[�g�̍���
    
    Set orgSheet = Workbooks("Book1").Worksheets("Sheet1")
    Set cpySheet = Workbooks("Book2").Worksheets("Sheet2")
    
    startRow = 5
    noCol = 1
    tagetCol = 2
    statusCol = 24
    
    selRow = 0
    
    ixRow = startRow
    Do While orgSheet.Cells(ixRow, noCol).Value <> ""
        If orgSheet.Cells(ixRow, statusCol).Value = "Y" Then
            selRow = selRow + 1
            cpySheet.Activate
            Cells(selRow + 2, 2).Select
            Selection.Value = orgSheet.Cells(ixRow, tagetCol).Value
            '   �Z�������ݒ�  �����ܕԂ��Ȃ�
            Selection.WrapText = False
        End If
        ixRow = ixRow + 1
    Loop

End Sub

'*******************************************************************************
'        �V�K�}�����C���̏�����
'*******************************************************************************
'        �����T�v�F�V�K�}�����郉�C�������̕����ŏ���������
'
'*******************************************************************************
Sub NewLineSet()

'   �V�K�}�����C��������������
    Dim lastCol         As Integer
    Dim userSelectCell  As Range

    'ActiveCell�̃o�b�N�A�b�v
    Set userSelectCell = ActiveCell
    
    '�����̍ŏIColumn �擾�i�w�b�_�[�s�j
    Cells.Find(What:="��", after:=Range("A1")).Select
    lastCol = Selection.End(xlToRight).Column

    '�}���s�̇���Ɋ֐���ݒ�
    userSelectCell.Select
    Selection.EntireRow.Select
    Selection.Insert Shift:=xlDown
    Cells(ActiveCell.row, ActiveCell.Column).Formula = _
        "=IF(LEN(B" & ActiveCell.row & ")>1,ROW()-ROW(A$1)-COUNTBLANK(A$2:A" & ActiveCell.row - 1 & "),"""")"

    '�}���s�͓h��Ԃ��Ȃ�
    Range(Cells(ActiveCell.row, ActiveCell.Column), Cells(ActiveCell.row, lastCol)).Select
    Selection.Interior.ColorIndex = xlNone
        
    '�}���s�{�P�s�̇���̊֐� �Đݒ� (��COUNTBLANK(A$2:Annn) �� nnn ���P�A�b�v���Ȃ�)
    '��L���ۂ́A�}���s�{�P�s�ڂ̂�
    Cells(ActiveCell.row + 1, ActiveCell.Column).Formula = _
        "=IF(LEN(B" & ActiveCell.row + 1 & ")>1,ROW()-ROW(A$1)-COUNTBLANK(A$2:A" & ActiveCell.row & "),"""")"
    
    Cells(ActiveCell.row, 1).Select
    
End Sub

'*******************************************************************************
'        �w��F�̕ύX
'*******************************************************************************
'        �����T�v�F�Z���̐F�����F�������琅�F�ɕύX����
'
'*******************************************************************************
Sub InteriorColorChange()

'   ���F���琅�F�ɕύX
    Workbooks("Book1").Worksheets("Sheet1").Activate
    With Application
        .FindFormat.Interior.ColorIndex = 6            ' ���F
        .ReplaceFormat.Interior.ColorIndex = 8         ' ���F
    End With
    
    ActiveSheet.UsedRange.Replace _
        What:="", replacement:="", SearchFormat:=True, ReplaceFormat:=True
        
    With Application
        .FindFormat.Clear
        .ReplaceFormat.Clear
    End With
    
End Sub

'*******************************************************************************
'        �d���s�̍폜�R���g���[��
'*******************************************************************************
'        �����T�v�F���L �d���s�̍폜�� CALL ����
'
'*******************************************************************************
Sub DuplicateRowsDelCall()

    Dim strColumn  As String
    
    Workbooks("Book1").Worksheets("Sheet1").Activate
    
    strColumn = "A"
    Call DuplicateRowsDelete3(strColumn)

End Sub

'*******************************************************************************
'        �d���s�̍폜
'*******************************************************************************
'        �����T�v�F�w�肳�ꂽ�V�[�g�̗�ʒu�ɂă\�[�g(�s�ʒu��1 �Œ�)��
'                  �㉺����̂Ƃ���̍s���폜����(�f�[�^���Ȃ��Ȃ�܂ŏ������s��)
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�V�[�g����       String
'            �����Q�@�F�J�����ʒu       Integer
'*******************************************************************************
Sub DuplicateRowsDelete3(strColumn As String)

    Dim strColumnRange As String
    Dim rngSortArea    As Range
    Dim rngCurrentCell As Range
    Dim rngNextCell    As Range
    
    strColumnRange = strColumn & "1"
    Set rngSortArea = Range(strColumnRange).CurrentRegion
    
    rngSortArea.sort Key1:=Range(strColumnRange), _
                     Order1:=xlAscending, _
                     Header:=xlGuess, _
                     OrderCustom:=1, _
                     MatchCase:=False, _
                     Orientation:=xlTopToBottom, _
                     SortMethod:=xlPinYin, _
                     DataOption1:=xlSortNormal
       
    Set rngCurrentCell = Range(strColumnRange)
    
    Do While Not IsEmpty(rngCurrentCell)
    
        Set rngNextCell = rngCurrentCell.Offset(1)
        If rngNextCell.Value = rngCurrentCell.Value Then
            rngCurrentCell.EntireRow.Delete
        End If
        
        Set rngCurrentCell = rngNextCell
    Loop
    
End Sub

'*******************************************************************************
'        �w��Z���R�s�[
'*******************************************************************************
'        �����T�v�FA��Z���̏��ƍŉE�[�Z���̏����擾����
'
'*******************************************************************************
sub
