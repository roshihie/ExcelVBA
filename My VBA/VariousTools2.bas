Attribute VB_Name = "Module2"
Sub �Z���R�����g�ꗗ�쐬()
    '�I���Z�����̃R�����g�ꗗ���쐬����}�N��
    '2004.10.29
    
    '�X�e�[�^�X�o�[�̕ۑ�
    ORG_BAR = Application.DisplayStatusBar
    Application.StatusBar = True
    '�I���Z���̃A�h���X���擾
    �J�n�� = Selection.Cells.Column
    �J�n�s = Selection.Cells.Row
    Set myRange = Selection.Cells
    �I���� = myRange.Columns(myRange.Columns.Count).Column
    �I���s = myRange.Rows(myRange.Rows.Count).Row
    '�A�N�e�B�u�u�b�N�̏���ۑ�
    Myworkbook = ActiveWorkbook.Name
    Myworksheet = ActiveSheet.Name
    '�V�K�u�b�N�̒ǉ�
    Workbooks.Add
    '�V�K�u�b�N�̃V�[�g���P�������ɂ���
    Call �V�[�g�폜
    '�V�K�u�b�N�̃V�[�g����ύX����
    Worksheets(1).Activate
    Worksheets(1).Name = Workbooks(Myworkbook).Worksheets(Myworksheet).Name
    '�ꗗ�w�b�_�ҏW�ƃV�[�g�̏����ݒ�
    ActiveSheet.Cells(1, 1) = "���΃A�h���X"
    ActiveSheet.Cells(1, 2) = "�s�ԍ�"
    ActiveSheet.Cells(1, 3) = "��ԍ�"
    ActiveSheet.Cells(1, 4) = "�Z���̓��e"
    ActiveSheet.Cells(1, 5) = "�R�����g�̓��e"
    Columns("A:E").VerticalAlignment = xlTop
    '�ꗗ�ҏW�J�n�s�ݒ�
    �s = 2
    For m = �J�n�s To �I���s
        For n = �J�n�� To �I����
            '�X�e�[�^�X�o�[���X�V
            Application.StatusBar = Cells(m, n).Address & "��������"
            '�Z���ɃR�����g���܂܂�Ă��邩�⍇��
            If TypeName(Workbooks(Myworkbook).Worksheets(Myworksheet).Cells(m, n).Comment) = "Nothing" Then
            Else
                '�Z���ɃR�����g���܂܂�Ă���ꍇ�A�ꗗ��ҏW����
                ��΃A�h���X = Workbooks(Myworkbook).Worksheets(Myworksheet).Cells(m, n).Address
                '��΃A�h���X�𑊑΃A�h���X�ɕҏW
                Call ���΃A�h���X�擾(��΃A�h���X)
                '�Z������ҏW
                ActiveSheet.Cells(�s, 1) = ��΃A�h���X
                ActiveSheet.Cells(�s, 2) = m
                ActiveSheet.Cells(�s, 3) = n
                ActiveSheet.Cells(�s, 4) = Workbooks(Myworkbook).Worksheets(Myworksheet).Cells(m, n)
                ActiveSheet.Cells(�s, 5) = Workbooks(Myworkbook).Worksheets(Myworksheet).Cells(m, n).Comment.Text
                '�ꗗ�s�C���N�������g
                �s = �s + 1
            End If
        Next
    Next
    '�ꗗ�s�̍s�Ɨ�̕��𐮂���
    Columns("A:E").EntireColumn.AutoFit
    Rows.EntireRow.AutoFit
    '�X�e�[�^�X�o�[�̕��A
    Application.StatusBar = False
    ORG_BAR = Application.DisplayStatusBar
End Sub

Sub �V�[�g�폜()
    '�P���̃V�[�g�݂̂ɂ���������W���[��
    Do While Sheets.Count > 1
        Application.DisplayAlerts = False
        Sheets(1).Delete
        Application.DisplayAlerts = True
    Loop
End Sub

Sub ���΃A�h���X�擾(��΃A�h���X)
    '��΃A�h���X���灐���폜����������W���[��
    ���΃A�h���X = ""
    For i = 1 To Len(��΃A�h���X) Step 1
        If Mid(��΃A�h���X, i, 1) = "$" Then
        Else
            ���΃A�h���X = ���΃A�h���X & Mid(��΃A�h���X, i, 1)
        End If
    Next
    ��΃A�h���X = ���΃A�h���X
End Sub

Sub �V�[�g�ꗗ�쐬()
    '�A�N�e�B�u�V�[�g�̃V�[�g�ꗗ���쐬����
    '2004.12.01
    
    '�A�N�e�B�u�V�[�g�̃u�b�N���ۊ�
    �u�b�N�� = ActiveWorkbook.Name
    '�V�K�u�b�N�ǉ�
    Workbooks.Add
    '�ꖇ�̃V�[�g�݂̂ɂ���
    Call �V�[�g�폜
    '�^�C�g���s�̕ҏW
    ActiveWorkbook.Sheets(1).Cells(1, 1) = "�u�b�N��"
    ActiveWorkbook.Sheets(1).Cells(1, 2) = "�V�[�g��"
    ActiveWorkbook.Sheets(1).Cells(1, 3) = "���l"
    ActiveWorkbook.Sheets(1).Cells(2, 1) = �u�b�N��
    '���݂���V�[�g�̐��������[�v
    For i = 1 To Workbooks(�u�b�N��).Sheets.Count
        '�V�[�g��
        ActiveWorkbook.Sheets(1).Cells(i + 1, 2) = Workbooks(�u�b�N��).Sheets(i).Name
        '��\���V�[�g���ǂ�������
        If Workbooks(�u�b�N��).Sheets(i).Visible = xlSheetVisible Then
        Else
            ActiveWorkbook.Sheets(1).Cells(i + 1, 3) = "��\��"
        End If
        '�󂫃V�[�g���ǂ�������
        If Workbooks(�u�b�N��).Sheets(i).UsedRange.Address = "$A$1" And _
           Workbooks(�u�b�N��).Sheets(i).Range("A1") = "" Then
            ActiveWorkbook.Sheets(1).Cells(i + 1, 3) = _
            ActiveWorkbook.Sheets(1).Cells(i + 1, 3) & "�i�󂫁j"
        End If
    Next
    '�J�������킹
    ActiveWorkbook.Sheets(1).Columns("A:C").EntireColumn.AutoFit
End Sub

Sub �V�[�g�����N�ꗗ�쐬()
    '�A�N�e�B�u�V�[�g�̃V�[�g�ꗗ���쐬����i�n�C�p�[�����N�t�j
    '2005.01.06
    '�V�[�g�ǉ����邩�ۂ��q�˂�
    Res = MsgBox("�ꗗ�p�V�[�g�ǉ�", vbYesNo)
    If Res = vbYes Then
        Sheets.Add Before:=Sheets(1)
    End If
    For i = 2 To ActiveWorkbook.Sheets.Count
        Sheets(1).Select
        Cells(i, 2).Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", _
        SubAddress:=Sheets(i).Name & "!A1", TextToDisplay:=Sheets(i).Name
        '�e�V�[�g�ɁuReturn�v�n�C�p�[�����N��t����
        'Sheets(i).Select
        'Cells(1, Sheets(i).UsedRange.Column + 1).Select
        'ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", _
        'SubAddress:=Sheets(1).Name & "!A1", TextToDisplay:="Return"
    Next
    Sheets(1).Select
End Sub

Sub �V�[�g�ǉ�()
    '�����V�[�g���i�[�z��
    Dim �����V�[�g��()
    '�������̖⍇��
    ������� = InputBox("�������", , ActiveSheet.Name)
    '������񂪖����Ƃ��A�������f
    If ������� = "" Then
        Exit Sub
    End If
    '�J��Ԃ����̖⍇��
    �J��Ԃ��� = InputBox("�J��Ԃ���", , "1")
    '�J��Ԃ����������Ƃ��A�������f
    If �J��Ԃ��� = "" Then
        Exit Sub
    End If
    '�����V�[�g���̎擾
    For j = 1 To ActiveWorkbook.Sheets.Count
        ReDim Preserve �����V�[�g��(j)
        �����V�[�g��(j) = ActiveWorkbook.Sheets(j).Name
    Next
    �����V�[�g�� = ActiveWorkbook.Sheets.Count
    '�J��Ԃ������A���[�v
    For i = 1 To �J��Ԃ��� Step 1
        '�A�N�e�B�u�V�[�g�̌�ɃV�[�g�ǉ�
        Sheets.Add After:=ActiveSheet
        '�V�[�g�������ݒ�
        �V�[�g�� = ������� & i
        '���ݒ肵���V�[�g���������V�[�g���Əd�����Ă��Ȃ����`�F�b�N
        Call �����V�[�g��r(�V�[�g��, �����V�[�g��(), �����V�[�g��)
        '�d�����Ȃ��V�[�g���ōX�V
        ActiveSheet.Name = �V�[�g��
    Next
End Sub

Sub �����V�[�g��r(�V�[�g��, �����V�[�g��(), �����V�[�g��)
'���x����`
level1:
    '�V�[�g���������V�[�g���Əd�����Ȃ��܂Ń��[�v
    For j = 1 To �����V�[�g��
        If �V�[�g�� = �����V�[�g��(j) Then
            �V�[�g�� = �V�[�g�� & "@"
            GoTo level1
        End If
    Next
End Sub

Sub �V�[�g���ύX()
    '�V�[�g�����������{�A�ԂŕύX����
    '�����������[�`�������������܂��B
    Randomize
    '��������⍇��
    ������� = InputBox("�������", , "Sheet1")
    '�������NULL�ł����s����\�������邽�߁A���s�L����⍇��
    If MsgBox("���s����", vbYesNo) = vbNo Then
        Exit Sub
    End If
    '��������
    ���� = Int((9999 * Rnd) + 1)
    '��x�V�[�g���������_���ɕύX
    For i = 1 To ActiveWorkbook.Sheets.Count
        Sheets(i).Name = ���� & i
    Next
    '���̌�������{�A�ԂŃV�[�g���ύX
    For i = 1 To ActiveWorkbook.Sheets.Count
        Sheets(i).Name = ������� & i
    Next
End Sub

Sub �V�[�g�w��폜()
    '�A�N�e�B�u�V�[�g�ȊO���폜
    '2004.12.06
    
    '�A�N�e�B�u�u�b�N�̃V�[�g�����P�̂Ƃ��A�����I��
    If ActiveWorkbook.Sheets.Count = 1 Then
        Exit Sub
    End If
    '�C���f�b�N�X�P�̃V�[�g���ƃA�N�e�B�u�V�[�g��������̂Ƃ�
    '�C���f�b�N�X�Q�ȍ~���폜
    If Sheets(1).Name = ActiveSheet.Name Then
        Do While ActiveWorkbook.Sheets.Count > 1
            Application.DisplayAlerts = False
            Sheets(2).Delete
            Application.DisplayAlerts = True
        Loop
    '�A�N�e�B�u�V�[�g�̃C���f�b�N�X���P�łȂ��Ƃ�
    Else
        '�A�N�e�B�u�V�[�g�̃C���f�b�N�X���P�ɂ��邽�߂ɃV�[�g�ړ�
        ActiveSheet.Move Before:=Sheets(1)
        Do While ActiveWorkbook.Sheets.Count > 1
            Application.DisplayAlerts = False
            Sheets(2).Delete
            Application.DisplayAlerts = True
        Loop
    End If
End Sub
