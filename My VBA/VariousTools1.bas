Attribute VB_Name = "Module1"
Option Compare Text
Sub ���������()
    '�I���Z�����̎�����������}�N��
    '2004.08.30
    
    '�X�e�[�^�X�o�[�̕ۑ�
    ORG_BAR = Application.DisplayStatusBar
    Application.StatusBar = True
    '�I���Z���̃A�h���X���擾
    �J�n�� = Selection.Cells.Column
    �J�n�s = Selection.Cells.Row
    Set myRange = Selection.Cells
    �I���� = myRange.Columns(myRange.Columns.Count).Column
    �I���s = myRange.Rows(myRange.Rows.Count).Row
    '�X�e�[�^�X�o�[��\��
    Application.StatusBar = Cells(�J�n�s, �J�n��).Address & "����" & Cells(�I���s, �I����).Address & "�܂ł�I�����Ă��܂�"
    '�ꊇ�u�����ۂ��₢���킹
    resp = MsgBox("�ꊇ�u���H", vbYesNo)
    '�I���Z�������ԂɎQ��
    For m = �J�n�s To �I���s
        For n = �J�n�� To �I����
            '�X�e�[�^�X�o�[��\��
            Application.StatusBar = Cells(m, n).Address & "��������"
            j = ""
            k = 0
            '�Z�����Ɏ����������Ƃ��A��������������l�𒊏o
            For i = 1 To Len(Cells(m, n)) Step 1
                If Cells(m, n).Characters(i, 1).Font.Strikethrough = False Then
                    j = j + Mid(Cells(m, n), i, 1)
                Else
                    k = k + 1
                End If
            Next
            '��������������ꍇ�̂ݒu������l��₢���킹
            If k > 0 Then
                '�ꊇ�u���łȂ��̂Ƃ��̂݁A�ϊ��O����m�F
                If resp = vbNo Then
                    j = InputBox(m & "�s" & n & "��̕ύX�O�F" & Cells(m, n), , j)
                End If
                Cells(m, n) = j
                Cells(m, n).Font.Strikethrough = False
            End If
        Next
    Next
    '�X�e�[�^�X�o�[�̕��A
    Application.StatusBar = False
    ORG_BAR = Application.DisplayStatusBar
    '�I�����b�Z�[�W�̏o��
    MsgBox "Finished"
End Sub
Sub �����F�t()
    '�I���Z�����̕����񌟍������ĐF��t����}�N��
    '2004.11.26
    On Error Resume Next
    '�X�e�[�^�X�o�[�̕ۑ�
    ORG_BAR = Application.DisplayStatusBar
    Application.StatusBar = True
    '�I���Z���̃A�h���X���擾
    �J�n�� = Selection.Cells.Column
    �J�n�s = Selection.Cells.Row
    Set myRange = Selection.Cells
    �I���� = myRange.Columns(myRange.Columns.Count).Column
    �I���s = myRange.Rows(myRange.Rows.Count).Row
    '�X�e�[�^�X�o�[��\��
    Application.StatusBar = Cells(�J�n�s, �J�n��).Address & "����" & Cells(�I���s, �I����).Address & "�܂ł�I�����Ă��܂�"
    '�����������₢���킹
    response1 = InputBox("������������w�肵�Ă�������" & vbCrLf & _
                         "�啶���������S�p���p�͋�ʂ���܂�" & vbCrLf & _
                         "�u*,?�v�͂��̂܂܎w�肵�Ă�������", , "(����������)")
    '���������񂪁iNULL�j�̏ꍇ�������~
    If response1 = "" Then
        MsgBox "�����𒆎~���܂�"
        Exit Sub
    End If
    response2 = InputBox("�����F���w�肵�Ă�������" & vbCrLf & _
                         "��F��[3]��[5]��[7]��[4]��", response1 & " ���������܂�", "(�����F1�`56)")
    '�����F���w��͈͊O�̏ꍇ�������~
    If response2 < 1 Or _
       response2 > 56 Then
        MsgBox "�����𒆎~���܂�"
        Exit Sub
    End If
    j = 0
    '�I���Z�������ԂɎQ��
    For m = �J�n�s To �I���s
        For n = �J�n�� To �I����
            '�X�e�[�^�X�o�[��\��
            Application.StatusBar = Cells(m, n).Address & "��������"
            '�����J�n�ʒu�̏����l�ݒ�
            i = 1
            '�Z���̒l�����[�N�֊i�[
            SearchString = Cells(m, n).Value
            '�Z���̒l�̒������������[�v
            Do While i <= Len(SearchString)
                '�Z���̒��Ɍ��������񂪂��邩����
                If InStr(i, SearchString, response1, vbBinaryCompare) > 0 Then
                    '�Z���̒��Ɍ��������񂪂���ʒu���L��
                    response3 = InStr(i, SearchString, response1, vbBinaryCompare)
                    '���������񂪂���ʒu���猟��������̒��������������ݒ�
                    '�F�ύX
                    Cells(m, n).Characters(response3, Len(response1)).Font.ColorIndex = response2
                    '�{�[���h
                    Cells(m, n).Characters(response3, Len(response1)).Font.Bold = True
                    '�����񌟍��̊J�n�ʒu���V�t�g����
                    i = response3 + Len(response1)
                    j = j + 1
                Else
                    '���������񂪌�����Ȃ������̂ŁA�����񌟍��̊J�n�ʒu���ő�ɃV�t�g����
                    i = i + Len(SearchString)
                End If
            Loop
        Next
    Next
    '�X�e�[�^�X�o�[�̕��A
    Application.StatusBar = False
    ORG_BAR = Application.DisplayStatusBar
    '�I�����b�Z�[�W�̏o��
    MsgBox "Finished" & vbCrLf & _
           "Changed �~ " & j
End Sub
Sub ���s�폜()
    '�J�E���^������
    l = 0
    '�g�p�Z���I��
    ActiveSheet.UsedRange.Select
    '�g�p�Z���͈̔̓A�h���X�擾
    �J�n�� = Selection.Cells.Column
    �J�n�s = Selection.Cells.Row
    Set myRange = Selection.Cells
    �I���� = myRange.Columns(myRange.Columns.Count).Column
    �I���s = myRange.Rows(myRange.Rows.Count).Row
    '�g�p�Z���������[�v
     For i = �J�n�s To �I���s
        For j = �J�n�� To �I����
            '���[�N�G���A������
            aaa = ""
            '�P�������������ĉ��s�R�[�h������ꍇ�������A
            '����ȊO�̏ꍇ�ڑ�����
            For k = 1 To Len(Cells(i, j))
                If Mid(Cells(i, j), k, 1) = Chr(10) Or _
                   Mid(Cells(i, j), k, 1) = Chr(13) Then
                    l = l + 1
                Else
                    aaa = aaa & Mid(Cells(i, j), k, 1)
                End If
            Next
            Cells(i, j) = aaa
        Next
    Next
    '�I�����b�Z�[�W���o�͂���
    MsgBox "���s�폜�~ " & l
End Sub
Sub CSV�쐬()
    '�I���Z����CSV������}�N��
    '2004.12.10
    
    Dim �o�̓��R�[�h() As String
    '�I���Z���̃A�h���X���擾
    �J�n�� = Selection.Cells.Column
    �J�n�s = Selection.Cells.Row
    Set myRange = Selection.Cells
    �I���� = myRange.Columns(myRange.Columns.Count).Column
    �I���s = myRange.Rows(myRange.Rows.Count).Row
    '�I���Z�������ԂɎQ��
    p = 0
    For m = �J�n�s To �I���s
        p = p + 1
        ReDim Preserve �o�̓��R�[�h(p)
        For n = �J�n�� To �I����
            �o�̓��R�[�h(p) = �o�̓��R�[�h(p) & """" & Cells(m, n) & ""","
        Next
    Next
    ���R�[�h�� = p
    �o�̓t�@�C�� = Application.GetSaveAsFilename(fileFilter:="�e�L�X�g �t�@�C�� (*.csv), *.csv")
    ' �o�̓t�@�C�����I�[�v������
    OutputFile = FreeFile
    Open �o�̓t�@�C�� For Output As #OutputFile
    For p = 1 To ���R�[�h��
        Print #OutputFile, �o�̓��R�[�h(p)
    Next
    ' �o�̓t�@�C�����N���[�Y����
    Close #OutputFile
    '�I�����b�Z�[�W�̏o��
    MsgBox "Finished"
End Sub
Sub �e�L�X�g�r���쐬()
    Dim ��() As Variant
    '�I���Z�����e�L�X�g�r���\�`���ɕϊ�����}�N��
    '2004.12.10
    
    '�I���Z���̃A�h���X���擾
    �J�n�� = Selection.Cells.Column
    �J�n�s = Selection.Cells.Row
    Set myRange = Selection.Cells
    �I���� = myRange.Columns(myRange.Columns.Count).Column
    �I���s = myRange.Rows(myRange.Rows.Count).Row
    '�񕝂̃J�E���g
    ���� = �I���� - �J�n�� + 1
    ReDim Preserve ��(����)
    '�񕝏����l�ݒ�
    For p = 1 To ����
        ��(p) = 2
    Next
    '�I���Z�������ԂɎQ�Ƃ��񕝂��L������
    For m = �J�n�s To �I���s
        p = 0
        For n = �J�n�� To �I����
            p = p + 1
            If ��(p) < Len(Cells(m, n).Value) Then
                ��(p) = Len(Cells(m, n).Value)
                If ��(p) Mod 2 > 0 Then
                    ��(p) = ��(p) + 1
                End If
            End If
        Next
    Next
    �o�̓t�@�C�� = Application.GetSaveAsFilename(fileFilter:="�e�L�X�g �t�@�C�� (*.txt), *.txt")
    ' �o�̓t�@�C�����I�[�v������
    OutputFile = FreeFile
    Open �o�̓t�@�C�� For Output As #OutputFile
    '�P�s��
    �o�̓��R�[�h = "��"
    For p = 1 To ���� Step 1
        For q = 1 To (��(p) / 2) Step 1
            �o�̓��R�[�h = �o�̓��R�[�h & "��"
        Next
        If p = ���� Then
            �o�̓��R�[�h = �o�̓��R�[�h & "��"
        Else
            �o�̓��R�[�h = �o�̓��R�[�h & "��"
        End If
    Next
    Print #OutputFile, �o�̓��R�[�h
    '�q��
    �o�̓��R�[�h = "��"
    For p = 1 To ���� Step 1
        For q = 1 To (��(p) / 2) Step 1
            �o�̓��R�[�h = �o�̓��R�[�h & "��"
        Next
        If p = ���� Then
            �o�̓��R�[�h = �o�̓��R�[�h & "��"
        Else
            �o�̓��R�[�h = �o�̓��R�[�h & "��"
        End If
    Next
    '�q�����R�[�h���Z�[�u
    �q�����R�[�h = �o�̓��R�[�h
    '�I���Z�������ԂɎQ��
    For m = �J�n�s To �I���s
        �o�̓��R�[�h = "��"
        p = 0
        For n = �J�n�� To �I����
            p = p + 1
            �Z�����e�L�X�g = Cells(m, n)
            Do While Len(�Z�����e�L�X�g) < ��(p)
                �Z�����e�L�X�g = �Z�����e�L�X�g & " "
            Loop
            �o�̓��R�[�h = �o�̓��R�[�h & �Z�����e�L�X�g & "��"
        Next
        Print #OutputFile, �o�̓��R�[�h
        If m = �I���s Then
        Else
            �o�̓��R�[�h = �q�����R�[�h
            Print #OutputFile, �o�̓��R�[�h
        End If
    Next
    '�ŏI�s
    �o�̓��R�[�h = "��"
    For p = 1 To ���� Step 1
        For q = 1 To (��(p) / 2) Step 1
            �o�̓��R�[�h = �o�̓��R�[�h & "��"
        Next
        If p = ���� Then
            �o�̓��R�[�h = �o�̓��R�[�h & "��"
        Else
            �o�̓��R�[�h = �o�̓��R�[�h & "��"
        End If
    Next
    Print #OutputFile, �o�̓��R�[�h
    ' �o�̓t�@�C�����N���[�Y����
    Close #OutputFile
    '�I�����b�Z�[�W�̏o��
    MsgBox "Finished"
End Sub
