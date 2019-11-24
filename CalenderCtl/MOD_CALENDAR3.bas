Attribute VB_Name = "MOD_CALENDAR3"
'*******************************************************************************
'   �J�����_�[�t�H�[��3(���t���͕��i)   ���Ăяo���v���V�[�W��
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'*******************************************************************************
Option Explicit
Public Const cnsDateFormat = "YYYY/MM/DD"   ' �f�t�H���g�̓��tFormat
Private Const cnsCaption = "���t�I��"       ' �f�t�H���g��Caption
Public g_swCalendar1Loaded As Boolean       ' Load����X�C�b�`

'*******************************************************************************
' ���[�U�[�t�H�[���̃e�L�X�g�{�b�N�X(MsForms.TextBox)����\��������
'*******************************************************************************
' [����]
' �@�E�e�L�X�g�{�b�N�X(Object�A�V�[�g����̏ꍇ�̓R���g���[���c�[���{�b�N�X�̕�)
' �@�E�J�����_�[�t�H�[����Caption(String) ��Option�A�f�t�H���g��"���t�I��"
' �@�E�l��Ԃ�����Format(String) ��Option�A�f�t�H���g��"YYYY/MM/DD"
' �@�E�J�����_�[�t�H�[���̕\���ʒu�F��(Long) ��Option
' �@�E�J�����_�[�t�H�[���̕\���ʒu�F�c(Long) ��Option
'*******************************************************************************
Public Sub ShowCalendarFromTextBox2(objTextBox As MSForms.TextBox, _
                                    Optional strCaption As String, _
                                    Optional strFormat As String, _
                                    Optional lngLeft As Long, _
                                    Optional lngTop As Long)
    Dim dteDate As Date
    
    ' ���ƂȂ���t���e�L�X�g�{�b�N�X����擾
    If IsDate(Trim(objTextBox.Text)) Then
        dteDate = CDate(Trim(objTextBox.Text))
    End If
    ' Caption(�^�C�g��)�w�肪�Ȃ��ꍇ�̓f�t�H���g("���t�I��")���w��
    If strCaption = "" Then strCaption = cnsCaption
    ' �\���t�H�[�}�b�g�w�肪�Ȃ��ꍇ�̓f�t�H���g("YYYY/MM/DD")���w��
    If strFormat = "" Then strFormat = cnsDateFormat
    ' �J�����_�[�t�H�[��
    With FRM_CALENDAR3
        ' Tag�Ɍ����t(�V���A���l)���Z�b�g
        .Tag = CLng(dteDate)
        ' Caption���Z�b�g
        .Caption = strCaption
        ' �t�H�[���\���ʒu�̊m�F
        If ((lngLeft <> 0) And (lngTop <> 0)) Then
            ' �w�肪����ꍇ�̓}�j���A���w��
            .StartUpPosition = 0
            .Left = lngLeft
            .Top = lngTop
        Else
            ' �w�肪�Ȃ��ꍇ�̓I�[�i�[�t�H�[���̒���
            .StartUpPosition = 1
        End If
        ' �J�����_�[�t�H�[����\��
        .Show
        ' �t�H�[����Unload���ꂽ�ꍇ�͈ȍ~�̏����𖳎�����
        On Error Resume Next
        ' Tag�̓��t���m�F
        If IsNumeric(.Tag) <> True Then Exit Sub
        If Err.Number <> 0 Then Exit Sub
        On Error GoTo 0
        ' Tag����I����t�����o���ăe�L�X�g�{�b�N�X�ɃZ�b�g
        dteDate = CDate(.Tag)
        objTextBox.Text = Format(dteDate, strFormat)
    End With
End Sub

'*******************************************************************************
' �Z��(Range)����\��������
'*******************************************************************************
' [����]
' �@�E�Z��(Object) �������Ƃ��ĒP��Z��
' �@�E�J�����_�[�t�H�[����Caption(String) ��Option�A�f�t�H���g��"���t�I��"
' �@�E�J�����_�[�t�H�[���̕\���ʒu�F��(Long) ��Option
' �@�E�J�����_�[�t�H�[���̕\���ʒu�F�c(Long) ��Option
'*******************************************************************************
Public Sub ShowCalendarFromRange2(objRange As Range, _
                                  Optional strCaption As String, _
                                  Optional lngLeft As Long, _
                                  Optional lngTop As Long)
    Dim dteDate As Date

    ' ���ƂȂ���t���Z������擾
    If IsDate(Trim(objRange.Value)) Then
        dteDate = CDate(Trim(objRange.Value))
    End If
    ' Caption(�^�C�g��)�w�肪�Ȃ��ꍇ�̓f�t�H���g("���t�I��")���w��
    If strCaption = "" Then strCaption = cnsCaption
    ' �J�����_�[�t�H�[��
    With FRM_CALENDAR3
        ' Tag�Ɍ����t(�V���A���l)���Z�b�g
        .Tag = CLng(dteDate)
        ' Caption���Z�b�g
        .Caption = strCaption
        ' �t�H�[���\���ʒu�̊m�F
        If ((lngLeft <> 0) And (lngTop <> 0)) Then
            ' �w�肪����ꍇ�̓}�j���A���w��
            .StartUpPosition = 0
            .Left = lngLeft
            .Top = lngTop
        Else
            ' �w�肪�Ȃ��ꍇ�̓I�[�i�[�t�H�[���̒���
            .StartUpPosition = 1
        End If
        ' �J�����_�[�t�H�[����\��
        .Show
        ' �t�H�[����Unload���ꂽ�ꍇ�͈ȍ~�̏����𖳎�����
        On Error Resume Next
        ' Tag�̓��t���m�F
        If IsNumeric(.Tag) <> True Then Exit Sub
        If Err.Number <> 0 Then Exit Sub
        On Error GoTo 0
        ' Tag����I����t�����o���ăZ���ɃZ�b�g
        dteDate = CDate(.Tag)
        objRange.Value = dteDate
    End With
End Sub

'--------------------------------<< End of Source >>----------------------------



