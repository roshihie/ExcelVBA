Attribute VB_Name = "Module2"
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'        �d���������u�a�`  �s������
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'*******************************************************************************
'        �f�[�^�ŉ��s�� SUM�֐���}��
'*******************************************************************************
Sub InputFormula()

    Dim i         As Integer
    Dim myLastClm As Integer
    Dim r         As Long
   '�ŏI����擾
    myLastClm = Range("A1").End(xlToRight).Column
    
   '�ŏI�s���擾
    r = Range("A1").End(xlDown).row
    
   'SUM�֐����ŏI�s�̉��ɓ���
    For i = 1 To myLastClm
        Cells(r + 1, i).FormulaR1C1 = "=SUM(R[" & -r & "]C:R[-1]C)"
    Next

End Sub

'*******************************************************************************
'        �V�K���͍s���擾����
'*******************************************************************************
Sub NewInputRowGet()

    Dim lastRow  As Long
    
    lastRow = ActiveSheet.Rows.Count
    Cells(lastRow, 1).End(xlUp).Offset(1).Select
    
End Sub

'*******************************************************************************
'        �I���Z���̃T�C�Y�ύX
'*******************************************************************************
Sub CellResize()

    Workbooks("Book1").Worksheets("Sheet1").Activate
    Range("b2:c5").Select
    MsgBox "�I���Z���͈͂̃T�C�Y��ύX���A�ʒu���Q�s���ɂ��炵�܂�"
    Selection.Offset(2).Resize(Selection.Rows.Count + 2, Selection.Columns.Count + 3).Select
    
End Sub

'*******************************************************************************
'        �A�N�e�B�u�Z���̈�S�̂�Select
'*******************************************************************************
'        ���@���@�F�A�N�e�B�u�Z���̈恁�󔒍s�Ƌ󔒗�ň͂܂ꂽ�Z���͈́��f�[�^�x�[�X
'                  �A�N�e�B�u�Z���̈�S�̂�Select �� CurrentRegion�ɂ��Select
'*******************************************************************************
Sub CurrentRegionSelect()

    Range("A1").CurrentRegion.Select
    
'   �f�[�^�x�[�X�̃f�[�^�������Z�o
    Dim myData���� As Long
    myData���� = Range("A1").CurrentRegion.Rows.Count - 1   ' ���o���s�����Z���Ă���
    
End Sub

'*******************************************************************************
'        �f�[�^�x�[�X���E�ׂ̃��[�N�V�[�g�ɃR�s�[
'*******************************************************************************
Sub Database()
    
'   �f�[�^�x�[�X�S�̂�I������
    Range("A1").CurrentRegion.Select
    
'   �f�[�^�x�[�X����s1�i���o���s�j���������f�[�^�͈͂��R�s�[����
    Selection.Offset(1).Resize(Selection.Rows.Count - 1).Copy
    
'   �E�ׂ̃��[�N�V�[�g���A�N�e�B�u�ɂ���
    ActiveSheet.Next.Activate
    
'   �R�s�[�����Z���͈͂�\��t����
    Range("A1").PasteSpecial

'   �R�s�[���[�h����������
    Application.CutCopyMode = False

End Sub

'*******************************************************************************
'        �f�[�^�x�[�X���̓���s��I������
'*******************************************************************************
Sub databaseRowSelect()

    Range("A6", Range("A6").End(xlToRight)).Select
    
End Sub

'*******************************************************************************
'        �f�[�^�x�[�X�ɊO�g�̌r��������
'*******************************************************************************
Sub DrawLineOfDatabese()

    With Range("B3").CurrentRegion
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
    End With
    
End Sub

'*******************************************************************************
'        �g�p����Ă���Z���͈͑S�̂�Select
'*******************************************************************************
'        ���@���@�F�g�p����Ă���Z���͈͑S�̂�Select �� UsedRange�ɂ��Select
'                  UsedRange�v���p�e�B�́uRange("A1").CurrentRegion�v�̂悤��
'                  �I���̊�ƂȂ�Z�����ӎ�����K�v�͂Ȃ�
'*******************************************************************************
Sub UsedRangeSelect()

    ActiveSheet.UsedRange.Select

End Sub

'*******************************************************************************
'        �d���s�̍폜�i������ւ̃��[�v�j
'*******************************************************************************
'        �����T�v�F�Z��An(A��̍ŏI�s)����X�^�[�g�� ���� �Z��An�`A2 �̓��e���r��
'                  �㉺����̂Ƃ����̍s���폜����
'*******************************************************************************
Sub DuplicateRowsDelete1()
    
    Dim lastRow   As Long
    Dim myLastRow As Long
    Dim i         As Long
    
'   ��ʂ̂������}�~���Ď��s���x�����コ����
    Application.ScreenUpdating = False
'   �A�N�e�B�u�V�[�g�̍ŏI�s���擾
    lastRow = ActiveSheet.Rows.Count
'   �f�[�^�s�̍ŏI�s���擾
    myLastRow = Cells(lastRow, 1).End(xlUp).row
    
    For i = myLastRow To 3 Step -1
        If Cells(i, 1).Value = Cells(i - 1, 1).Value Then
            Cells(i, 1).EntireRow.Delete
        End If
    Next i
    
End Sub

'*******************************************************************************
'        �d���s�̍폜�i�������ւ̃��[�v�j
'*******************************************************************************
'        �����T�v�F�Z��A2����X�^�[�g�� ���� �Z��A2�`An �̓��e���r��
'                  �㉺����̂Ƃ���̍s���폜����
'
'                  myLastRow�ɂ͍폜�O�̃f�[�^�i�[�ŏI�s�� ���i�[����邪
'                  ����� �������s���ׂ��s�� �ł���
'                  ����s�̂Ƃ� �s�͍폜����邪 �폜��̌��ݍs���Ƃ͊֌W�Ȃ�
'*******************************************************************************
Sub DuplicateRowsDelete2()

    Dim lastRow   As Long
    Dim myLastRow As Long
    Dim i         As Long
    
'   ��ʂ̂������}�~���Ď��s���x�����コ����
    Application.ScreenUpdating = False
'   �A�N�e�B�u�Z���̍ŏI�s���擾
    lastRow = ActiveSheet.Rows.Count
'   �f�[�^�s�̍ŏI�s���擾
    myLastRow = Cells(lastRow, 1).End(xlUp).row
    'Debug.Print "myLastRow = (" & myLastRow & ")"
    Range("A2").Select
    
    For i = 2 To myLastRow
        If Selection.Value = Selection.Offset(1).Value Then
            Selection.EntireRow.Delete
        Else
            Selection.Offset(1).Select
        End If
    Next i
    'Debug.Print "myLastRow = (" & myLastRow & ")"
    
End Sub

'*******************************************************************************
'        �A�N�e�B�u�Z���̈�̒��̉��Z���������R�s�[����
'*******************************************************************************
Sub CopyVisibleRange()

'   �A�N�e�B�u�Z���̈��I������
    Range("A1").CurrentRegion.Select

'   ���Z���������R�s�[����
    Selection.SpecialCells(xlCellTypeVisible).Copy

'   �ʂ̃V�[�g�ɓ\��t����
    Worksheets("Sheet2").Select
    ActiveSheet.Paste

    Application.CutCopyMode = False
    
End Sub

'*******************************************************************************
'        �I�[�g�t�B���^���ꂽ�Y���f�[�^�̌������擾
'*******************************************************************************
'        ���@���@�FRows�v���p�e�B�ɂ́A�����̗̈�̒��ōŏ��̗̈�̍s�������Q�Ƃł��Ȃ�
'                  VBA�́u�̈�v�Ƃ����אڂ��Ă���Z���͈͂� Areas�R���N�V���� �Ƃ���
'                  ������
'*******************************************************************************
Sub CountSelectedData()

    Dim myArea As Range
    Dim myRow  As Integer
    '�S�Ẳ���Ԃ̗̈��I������
    Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Select

   '�e�̈悲�Ƃɍs�����擾���ĉ��Z����
    For Each myArea In Selection.Areas
        myRow = myRow + myArea.Rows.Count
    Next

   '�������s������������ ���Z����
    MsgBox "���o������ " & (myRow - 1) & "���ł�"
    
End Sub

'*******************************************************************************
'        �f�[�^�����͂��ꂽ�ŏI�Z��(UsedRange�̍ŏI�Z��)��Select
'*******************************************************************************
Sub SelectLastCell()

    Dim myLastCell As Range
    Dim lastRow    As Long
    Dim lastCol    As Integer
    
    Set myLastCell = Range("A1").SpecialCells(xlCellTypeLastCell)
    
    With myLastCell
        If .row = 1 And .Column = 1 Then
            myLastCell.Select
            Exit Sub
        End If
    End With
    
    'SpecialCells(xlCellTypeLastCell)���\�b�h�͈�x�f�[�^����͂����Z�����N���A���Ă�
    '���̂܂܍Ō�̃Z���ƔF�����Ă��܂����� �␳����
    If myLastCell.Value = "" Then
    
'       Find(��What:=����������,
'            After:=�w�肵���Z���̎����猟��,
'            LookIn:=�������e[  ����[xlFormulas],
'                             ���l[xlValue],
'                               �R�����g[xlComments]],
'            LookAt:=�����Ώە���[���ꕔ����v�Ō���[xlPart],
'                                   �S�Ĉ�v�Ō���[xlWhole]],
'            SearchOrder:=��������[  ��[xlByColumns],
'                                  ���s[xlByRows]],
'            SearchDirection:=��(��������)�s�̂Ƃ������E,(��������)��̂Ƃ��と��[xlNext],
'                               (��������)�s�̂Ƃ��E����,(��������)��̂Ƃ�������[xlPrevious]
'            MatchCase:=  �啶��������������[True],
'                       ���啶���E�������C�� [False]
'            MatchByte:=  �S�p����p�����[True],
'                       ���S�p�E���p�C�� [False]
        lastRow = Cells.Find(What:="*", after:=myLastCell, _
                             SearchOrder:=xlByRows, _
                             SearchDirection:=xlPrevious).row
        lastCol = Cells.Find(What:="*", after:=myLastCell, _
                             SearchOrder:=xlByColumns, _
                             SearchDirection:=xlPrevious).Column
        Cells(lastRow, lastCol).Select
    Else
        myLastCell.Select
    End If
    
End Sub

'*******************************************************************************
'        ��̕\���^��\����؂�ւ���
'*******************************************************************************
Sub ToggleColumn()
    
    With Worksheets("Sheet1").Columns("C:D")
        .Hidden = Not .Hidden
    End With
    
End Sub

'*******************************************************************************
'        �I������Ă���Z���͈͂̏�Ƀe�L�X�g�{�b�N�X�쐬
'*******************************************************************************
Sub DrawTextbox()

'   �I���Z���͈͂̊J�n�ʒu
    Dim myLeft   As Variant
    Dim myTop    As Variant
'   �I���Z���͈͂̑傫��
    Dim myWidth  As Variant
    Dim myHeight As Variant
    
    myLeft = Selection.Left
    myTop = Selection.Top
    myWidth = Selection.Width
    myHeight = Selection.Height
        
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        myLeft, myTop, myWidth, myHeight).Select
        
End Sub

'*******************************************************************************
'        �f�[�^�R�s�[�ɂ��V�K�u�b�N�쐬
'*******************************************************************************
'        �����T�v�F�V�K�u�b�N���쐬���Ċ����u�b�N����f�[�^���R�s�[��
'                  ���O��t���ĕۑ�����
'*******************************************************************************
Sub LookNewBook()
    
    Dim myNewWK As Workbook
    
'   �V�K�u�b�N���쐬�� ���̎Q�Ƃ��I�u�W�F�N�g�^�ϐ��ɑ������
    Set myNewWK = Workbooks.Add
    
'   Add���\�b�h���s�� �ǉ����ꂽ�V�K�u�b�N���A�N�e�B�u�ƂȂ�̂ňȉ��̒ʂ� �擾���\
'   myWSName = ActiveWorkbook.Name
    
'   �}�N�������s���Ă���u�b�N(Sample.xls)���A�N�e�B�u�ɂ���
    Workbooks("Sample.xls").Activate
    Worksheets(1).Range("A1:E10").Copy
    
'   �I�u�W�F�N�g�^�ϐ��𗘗p���ĐV�K�u�b�N���A�N�e�B�u�ɂ���
    myNewWK.Activate
    Worksheets(1).Activate
    ActiveSheet.Paste
    
'   �V�K�u�b�N�𖼑O��t���ĕۑ����ĕ���
    myNewWK.SaveAs "NewBook.xls"
    Workbooks("NewBook.xls").Close

    Application.CutCopyMode = False
    
End Sub

'*******************************************************************************
'        �u�b�N�ύX�� �㏑���ۑ�
'*******************************************************************************
Sub CheckSaved()

'   �u�b�N���ύX���ꂽ�Ƃ� Saved�v���p�e�B��False���Z�b�g
    If ActiveWorkbook.Saved = False Then
       ActiveWorkbook.Save
    End If
    
End Sub

'*******************************************************************************
'        �ۑ��m�F���b�Z�[�W��\�������Ƀu�b�N�ۑ��E�N���[�Y
'*******************************************************************************
Sub SaveClose()

'   SaveChanges �� False��������ƃu�b�N�͕ۑ����ꂸ�ɕ���
    ActiveWorkbook.Close SaveChanges:=True
    
End Sub

'*******************************************************************************
'        �J�����u�b�N�Ɠ����t�H���_�ɂ���ʂ̃u�b�N(DataBook.xls)���J��
'*******************************************************************************
Private Sub Workbook_Open()

    Application.ScreenUpdating = False

'   �J�����g�h���C�u��ύX
    ChDrive ActiveWorkbook.Path
'   �J�����g�t�H���_��ύX
    ChDir ActiveWorkbook.Path
    
    Workbooks.Open Filename:="DataBook.xls"

'   �J�����g�h���C�u�C�t�H���_��ύX�������Ȃ��Ƃ�
'   Path�v���p�e�B���Ԃ�������� Open���\�b�h�̈����Ɏw�肷��
    Dim myPath  As String
    
    myPath = ActiveWorkbook.Path
    Workbooks.Open Filename:=myPath & "\DataBook.xls"
    
'   ��x���ۑ�����Ă��Ȃ��V�K�u�b�N�̏ꍇ�ɂ�
'   Path�v���p�e�B�͋�̕�����(Null�l)��Ԃ�
    
End Sub

'*******************************************************************************
'        ���̃��[�U�g�p���u�b�N  �����I�N���[�Y
'*******************************************************************************
Sub CloseReadOnlyBook()

'   Excel�� ���̃��[�U���g�p���̃u�b�N���J���� ���̃u�b�N�͎����I�Ɂu�ǎ��p�v�ƂȂ�
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Workbooks.Open "����.xls"

    '�ǂݎ���p�������炻�̃u�b�N�����
    If ActiveWorkbook.ReadOnly = True Then
        ActiveWorkbook.Close
        MsgBox "����.xls�͑��̃��[�U�[���g�p���ł�"
    End If
    
'   �u�b�N���n�߂���ǂݎ���p�Ɏw�肳��Ă����� ���̃��[�U�[���g�p���Ă��Ȃ��Ă�
'   �J�����Ƃ͂ł��Ȃ�
'   ������ VBA�̎d�l���l����� ���̕��@���ł��������ǂ� �܂��\�����p�ɑς�����}�N���ƌ�����
    
End Sub

'*******************************************************************************
'        �B���V�[�g
'*******************************************************************************
'        �����T�v�F�V�[�g�����[�U�[���\���ł��Ȃ��悤�ɉB��
'*******************************************************************************
Sub HiddenSheet()

'   xlVeryHidden�ɂ��
'   Excel�́m����(O)�n���m�V�[�g(H)�n���m�ĕ\��(U)...�n�R�}���h�������ƂȂ�
    Worksheets("Sheet3").Visible = xlVeryHidden
    
'   xlVeryHidden�ɂ���Ĕ�\���ɂȂ����V�[�g��
'   Visible�v���p�e�B��True��������΍ĕ\�������
    
End Sub

'*******************************************************************************
'        �u�b�N��\��
'*******************************************************************************
Sub HiddenBook()

'   �u�b�N���\���ɂ���
'   Workbook�I�u�W�F�N�g�ɂ�Visible�v���p�e�B���Ȃ�����
'   �u�b�N�̂��ׂĂ�Window�I�u�W�F�N�g��Visible�v���p�e�B��False�������Ȃ���΂Ȃ�Ȃ�
    Dim myWindow  As Window

    For Each myWindow In ActiveWorkbook.Windows
        myWindow.Visible = False
    Next myWindow

'   ���[�U�[�́m�E�B���h�E(W)�n�|�m�ĕ\��(U)...�n�R�}���h��
'   �u�b�N���ĕ\�����邱�Ƃ��ł���

End Sub

'*******************************************************************************
'        �}�N���ɂ��ύX�͎󂯓����悤�ɃV�[�g��ی삷��
'*******************************************************************************
Sub ProtectWSheet()

'   �p�X���[�h"pswd1961"�ŃV�[�g��ی삷��
    Worksheets("Sheet1").Protect _
        Password:="pswd1961", _
        UserInterfaceOnly:=True

'   �Z���̓��e���}�N������ύX����
    Range("A1:C10").Value = Array("ABC", "DEF", "HIJ")
    
'   UserInterfaceOnly �� False ���w�肵��Protect���\�b�h�����s�������[�N�V�[�g�̏ꍇ
'  �iUserInterfaceOnly���ȗ������ꍇ�����l�ł���j
'   ���[�U�[�̓��b�N���ꂽ�Z���̓��e���蓮�ŕύX���邱�Ƃ͂ł��Ȃ�
'   ������ �}�N���ŃZ���̓��e��ύX���邱�Ƃ��ł��Ȃ��Ȃ�
    
'   �V�[�g�̕ی����������
    Worksheets("Sheet1").Unprotect Password:="pswd1961"
    
End Sub

'*******************************************************************************
'        ���͔͈͐���
'*******************************************************************************
Sub LimitArea()

    With Worksheets("�[�i��")
        '�X�N���[���͈͂ɖ��O���`
        .Range("A1:I14").Name = "���͔͈�"
        
        '�X�N���[���͈͂𐧌�
        '�Z���͈�A1:I14�ȊO�̃Z����I��������
        '�Z���͈�A1:I14���B��Ă��܂��悤�ȉ�ʃX�N���[���͕s�\�ɂȂ�
        .ScrollArea = "���͔͈�"
        
'   �S�Z����I���ł���悤�ɐݒ��߂��Ƃ��ɂ� ScrollArea�v���p�e�B�ɋ�̕��������
        
        '���b�N���������ꂽ�Z���̂ݓ��͉\�Ƃ���
        .EnableSelection = xlUnlockedCells
        
        '�V�[�g��ی삷��
        .Protect Contents:=True, UserInterfaceOnly:=True
    End With
    
End Sub

'*******************************************************************************
'        �A�N�e�B�u�E�B���h�E�ȊO�̃E�B���h�E���ŏ�������
'*******************************************************************************
Sub WindowIcon()

    Dim myWindow  As Window
    Dim myWndName As String
    
    '�A�N�e�B�u�E�B���h�E�̖��O���擾 (Name�v���p�e�B�ł͂Ȃ����Ƃɒ���)
    myWndName = ActiveWindow.Caption
    
    For Each myWindow In Windows
        '�A�N�e�B�u�E�B���h�E�łȂ�������ŏ�������
        If myWindow.Caption <> myWndName Then
            myWindow.WindowState = xlMinimized
        End If
        
    Next myWindow
    
End Sub

'*******************************************************************************
'        Excel �E�B���h�E�ŏ����E�\��
'*******************************************************************************
Sub WindowSize()

'   Excel�̃E�B���h�E���ŏ�������
    Application.WindowState = xlMinimized
    
'   Excel�̃E�B���h�E��\������
    Application.WindowState = xlNormal

End Sub

'*******************************************************************************
'        Sort����Test
'*******************************************************************************
Sub SortTest()

    Dim rngSort   As Range
    Dim strColumn As String
    
    Workbooks("Book1").Worksheets("Sheet1").Activate
    
    strColumn = "A1"
    Set rngSort = Range(strColumn).CurrentRegion
    
    'rngSort.Select
    'Worksheets(strSheetName).Range(strColumnRange).sort
    rngSort.sort _
                             Key1:=Range(strColumn), _
                             Order1:=xlAscending, _
                             Header:=xlGuess, _
                             OrderCustom:=1, _
                             MatchCase:=False, _
                             Orientation:=xlTopToBottom, _
                             SortMethod:=xlPinYin, _
                             DataOption1:=xlSortNormal

End Sub

