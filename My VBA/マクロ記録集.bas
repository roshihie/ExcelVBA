Attribute VB_Name = "Module4"
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'        �}�N���L�^�W
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'
Sub �Z�������ݒ�_������ܕԂ��Ȃ�Macro()
'
' �Z�������ݒ�_������ܕԂ��Ȃ� Macro
' �}�N���L�^�� : 2005/1/10  ���[�U�[�� :
'
    Range("A1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
    '   ������ܕԂ��Ȃ�
        .WrapText = False
        
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub �Z�������ݒ�_������ܕԂ�����Macro()
'
' �Z�������ݒ�_������ܕԂ����� Macro
' �}�N���L�^�� : 2005/1/10  ���[�U�[�� :
'
    Range("A1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
    '   ������ܕԂ�����
        .WrapText = True
        
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub �Z�������ݒ�_�Ԋ|��_����Macro()
'
' �Z�������ݒ�_�Ԋ|������ Macro
' �}�N���L�^�� : 2005/1/10  ���[�U�[�� :
'
    Range("A1").Select
'   �Ԋ|������
    Selection.Interior.ColorIndex = xlNone
    
End Sub
Sub �Z�������ݒ�_�Ԋ|��_�O���[Macro()
'
' �Z�������ݒ�_�Ԋ|��_�O���[ Macro
' �}�N���L�^�� : 2005/1/10  ���[�U�[�� :
'
    Range("A1").Select
    With Selection.Interior
    '   �Ԋ|��(�O���[)
        .ColorIndex = 15
        
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
End Sub
Sub �֐�Find�g�p���@Macro()
'
' Find�֐��g�p���@ Macro
' �}�N���L�^�� : 2005/1/16  ���[�U�[�� :
'
' Find(What:=����������,
'      After:=�w�肵���Z���̎����猟��,
'      LookIn:=�������e[����[xlFormulas],
'                       �l[xlValues],  ��
'                       �R�����g[xlComments]],
'      LookAt:=�����Ώە���[�ꕔ����v�Ō���[xlPart],  ��
'                           �S�Ĉ�v�Ō���[xlWhole]],
'      SearchOrder:=��������[��[xlByColumns],
'                            �s[xlByRows]],  ��
'      SearchDirection:=�s�̂Ƃ������E,��̂Ƃ��と��[xlNext],  ��
'                       �s�̂Ƃ��E����,��̂Ƃ�������[xlPrevious]
'      MatchCase:=�啶��������������[True],
'                 �啶���E�������C�� [False]  ��
'      MatchByte:=�S�p����p�����[True],
'                 �S�p�E���p�C�� [False]  ��
    Workbooks("�e�X�gBook.xls").Activate
    Cells.Find(What:="������z", after:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, _
               SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
               MatchCase:=False, SearchFormat:=False).Activate
    Cells.FindNext(after:=ActiveCell).Activate
End Sub
Sub �s�}��Macro()
'
' �s�}�� Macro
' �}�N���L�^�� : 2005/2/27  ���[�U�[�� :
'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
End Sub
Sub �����񁨐��l�ϊ�Macro()
'
' �����񁨐��l�ϊ� Macro
' �}�N���L�^�� : 2005/5/15  ���[�U�[�� :
'
    Range("K5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J7:J18").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
End Sub
Sub Sort����Macro()
'
' Sort����Macro
' �}�N���L�^�� : 2009/2/1  ���[�U�[�� :
'
    Range("A1:A27").sort Key1:=Range("A1"), _
                         Order1:=xlAscending, _
                         Header:=xlGuess, _
                         OrderCustom:=1, _
                         MatchCase:=False, _
                         Orientation:=xlTopToBottom, _
                         SortMethod:=xlPinYin, _
                         DataOption1:=xlSortNormal
End Sub

