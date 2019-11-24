VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_CALENDAR3 
   Caption         =   "���t�I��"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2385
   OleObjectBlob   =   "FRM_CALENDAR3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FRM_CALENDAR3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'   �J�����_�[�t�H�[��3(���t���͕��i)    �����[�U�[�t�H�[��
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'*******************************************************************************
Option Explicit
'-------------------------------------------------------------------------------
' [�N�Z�j��] ���J�����_�[�����j�J�n(�j�����[)�ɂ���ꍇ�͢2��ɕύX���ĉ������B
Private Const g_cnsStartYobi = 1                ' 1=���j��,2=���j��(���͕s��)
'-------------------------------------------------------------------------------
' [�N�̕\�����x(From/To)]
Private Const g_cnsYearFrom = 1947              ' �j���@�{�s
Private Const g_cnsYearToAdd = 3                ' �V�X�e�����̔N+n�N�܂ł̎w��
'-------------------------------------------------------------------------------
' �t�H�[����̐F�w�蓙�̒萔
Private Const cnsBC_Select = &HFFCC33           ' �I����t�̔w�i�F
Private Const cnsBC_Other = &HE0E0E0            ' �����ȊO�̔w�i�F
Private Const cnsBC_Sunday = &HFFDDFF           ' ���j�̔w�i�F
Private Const cnsBC_Saturday = &HDDFFDD         ' �y�j�̔w�i�F
Private Const cnsBC_Month = &HFFFFFF            ' �����y���ȊO�̔w�i�F
Private Const cnsFC_Hori = &HFF                 ' �j���̕����F
Private Const cnsFC_Normal = &HC00000           ' �j���ȊO�̕����F
Private Const cnsDefaultGuide = "���L�[�ő���ł��܂��B"
'-------------------------------------------------------------------------------
' �t�H�[���\�����ɕێ����郂�W���[���ϐ�
Private tblDate(1 To 45) As MSForms.Label       ' ���t���x��
Private tblDate2(1 To 45) As Date               ' ���t
Private tblYobi(1 To 45) As Integer             ' �j��
Private tblGuide(1 To 45) As String             ' �K�C�h
Private g_intCurYear As Integer                 ' ���ݕ\���N
Private g_intCurMonth As Integer                ' ���ݕ\����
Private g_FormDate1 As Date                     ' ���ݓ��t
Private g_CurPos As Integer                     ' ���ݓ��t�ʒu
Private g_POS_F As Integer                      ' �������ʒu
Private g_POS_T As Integer                      ' �������ʒu
Private g_swBatch As Boolean                    ' �C�x���g�}��SW
Private g_VisibleYear As Boolean                ' Conbo�̔N�\���X�C�b�`
Private g_VisibleMonth As Boolean               ' Combo�̌��\���X�C�b�`
Private g_intSunday As Integer                  ' ���j���̗j���R�[�h
Private g_intSaturday As Integer                ' �y�j���̗j���R�[�h

'*******************************************************************************
' ���t�H�[����̃C�x���g
'*******************************************************************************
' �����R���{�̑���C�x���g
'*******************************************************************************
Private Sub CBO_MONTH_Change()
    Dim intMonth As Integer

    If g_swBatch Then Exit Sub
    intMonth = CInt(CBO_MONTH.Text)
    g_FormDate1 = DateSerial(g_intCurYear, intMonth, 1)
    Call ERASE_YEAR_MONTH                       ' �N���R���{�̔�\����
    ' �J�����_�[�쐬
    Call GP_MakeCalendar
End Sub

'*******************************************************************************
' ��N��R���{�̑���C�x���g
'*******************************************************************************
Private Sub CBO_YEAR_Change()
    Dim intYear As Integer
    
    If g_swBatch Then Exit Sub
    intYear = CInt(CBO_YEAR.Text)
    g_FormDate1 = DateSerial(intYear, g_intCurMonth, 1)
    Call ERASE_YEAR_MONTH                       ' �N���R���{�̔�\����
    ' �J�����_�[�쐬
    Call GP_MakeCalendar
End Sub

'*******************************************************************************
' �e���t���x���̃C�x���g(�N���X�����͂��Ȃ��ł��ꂼ��Click�C�x���g���Ŏ󂯂�)
'*******************************************************************************
' �e���t���x��(7�j�~6�T=42���A�Ή����t�͕\�����_�Ŕz�񉻂���Ă���)
Private Sub LBL_01_Click(): Call GP_ClickCalendar(tblDate2(1)):  End Sub
Private Sub LBL_02_Click(): Call GP_ClickCalendar(tblDate2(2)):  End Sub
Private Sub LBL_03_Click(): Call GP_ClickCalendar(tblDate2(3)):  End Sub
Private Sub LBL_04_Click(): Call GP_ClickCalendar(tblDate2(4)):  End Sub
Private Sub LBL_05_Click(): Call GP_ClickCalendar(tblDate2(5)):  End Sub
Private Sub LBL_06_Click(): Call GP_ClickCalendar(tblDate2(6)):  End Sub
Private Sub LBL_07_Click(): Call GP_ClickCalendar(tblDate2(7)):  End Sub
Private Sub LBL_08_Click(): Call GP_ClickCalendar(tblDate2(8)):  End Sub
Private Sub LBL_09_Click(): Call GP_ClickCalendar(tblDate2(9)):  End Sub
Private Sub LBL_10_Click(): Call GP_ClickCalendar(tblDate2(10)): End Sub
Private Sub LBL_11_Click(): Call GP_ClickCalendar(tblDate2(11)): End Sub
Private Sub LBL_12_Click(): Call GP_ClickCalendar(tblDate2(12)): End Sub
Private Sub LBL_13_Click(): Call GP_ClickCalendar(tblDate2(13)): End Sub
Private Sub LBL_14_Click(): Call GP_ClickCalendar(tblDate2(14)): End Sub
Private Sub LBL_15_Click(): Call GP_ClickCalendar(tblDate2(15)): End Sub
Private Sub LBL_16_Click(): Call GP_ClickCalendar(tblDate2(16)): End Sub
Private Sub LBL_17_Click(): Call GP_ClickCalendar(tblDate2(17)): End Sub
Private Sub LBL_18_Click(): Call GP_ClickCalendar(tblDate2(18)): End Sub
Private Sub LBL_19_Click(): Call GP_ClickCalendar(tblDate2(19)): End Sub
Private Sub LBL_20_Click(): Call GP_ClickCalendar(tblDate2(20)): End Sub
Private Sub LBL_21_Click(): Call GP_ClickCalendar(tblDate2(21)): End Sub
Private Sub LBL_22_Click(): Call GP_ClickCalendar(tblDate2(22)): End Sub
Private Sub LBL_23_Click(): Call GP_ClickCalendar(tblDate2(23)): End Sub
Private Sub LBL_24_Click(): Call GP_ClickCalendar(tblDate2(24)): End Sub
Private Sub LBL_25_Click(): Call GP_ClickCalendar(tblDate2(25)): End Sub
Private Sub LBL_26_Click(): Call GP_ClickCalendar(tblDate2(26)): End Sub
Private Sub LBL_27_Click(): Call GP_ClickCalendar(tblDate2(27)): End Sub
Private Sub LBL_28_Click(): Call GP_ClickCalendar(tblDate2(28)): End Sub
Private Sub LBL_29_Click(): Call GP_ClickCalendar(tblDate2(29)): End Sub
Private Sub LBL_30_Click(): Call GP_ClickCalendar(tblDate2(30)): End Sub
Private Sub LBL_31_Click(): Call GP_ClickCalendar(tblDate2(31)): End Sub
Private Sub LBL_32_Click(): Call GP_ClickCalendar(tblDate2(32)): End Sub
Private Sub LBL_33_Click(): Call GP_ClickCalendar(tblDate2(33)): End Sub
Private Sub LBL_34_Click(): Call GP_ClickCalendar(tblDate2(34)): End Sub
Private Sub LBL_35_Click(): Call GP_ClickCalendar(tblDate2(35)): End Sub
Private Sub LBL_36_Click(): Call GP_ClickCalendar(tblDate2(36)): End Sub
Private Sub LBL_37_Click(): Call GP_ClickCalendar(tblDate2(37)): End Sub
Private Sub LBL_38_Click(): Call GP_ClickCalendar(tblDate2(38)): End Sub
Private Sub LBL_39_Click(): Call GP_ClickCalendar(tblDate2(39)): End Sub
Private Sub LBL_40_Click(): Call GP_ClickCalendar(tblDate2(40)): End Sub
Private Sub LBL_41_Click(): Call GP_ClickCalendar(tblDate2(41)): End Sub
Private Sub LBL_42_Click(): Call GP_ClickCalendar(tblDate2(42)): End Sub
'-------------------------------------------------------------------------------
' ����A�����A�������x��
Private Sub LBL_43_Click(): Call GP_ClickCalendar(tblDate2(43)): End Sub
Private Sub LBL_44_Click(): Call GP_ClickCalendar(tblDate2(44)): End Sub
Private Sub LBL_45_Click(): Call GP_ClickCalendar(tblDate2(45)): End Sub
'-------------------------------------------------------------------------------
' �σK�C�h���b�Z�[�W
Private Sub LBL_01_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(1): End Sub
Private Sub LBL_02_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(2): End Sub
Private Sub LBL_03_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(3): End Sub
Private Sub LBL_04_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(4): End Sub
Private Sub LBL_05_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(5): End Sub
Private Sub LBL_06_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(6): End Sub
Private Sub LBL_07_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(7): End Sub
Private Sub LBL_08_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(8): End Sub
Private Sub LBL_09_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(9): End Sub
Private Sub LBL_10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(10): End Sub
Private Sub LBL_11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(11): End Sub
Private Sub LBL_12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(12): End Sub
Private Sub LBL_13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(13): End Sub
Private Sub LBL_14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(14): End Sub
Private Sub LBL_15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(15): End Sub
Private Sub LBL_16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(16): End Sub
Private Sub LBL_17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(17): End Sub
Private Sub LBL_18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(18): End Sub
Private Sub LBL_19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(19): End Sub
Private Sub LBL_20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(20): End Sub
Private Sub LBL_21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(21): End Sub
Private Sub LBL_22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(22): End Sub
Private Sub LBL_23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(23): End Sub
Private Sub LBL_24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(24): End Sub
Private Sub LBL_25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(25): End Sub
Private Sub LBL_26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(26): End Sub
Private Sub LBL_27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(27): End Sub
Private Sub LBL_28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(28): End Sub
Private Sub LBL_29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(29): End Sub
Private Sub LBL_30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(30): End Sub
Private Sub LBL_31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(31): End Sub
Private Sub LBL_32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(32): End Sub
Private Sub LBL_33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(33): End Sub
Private Sub LBL_34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(34): End Sub
Private Sub LBL_35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(35): End Sub
Private Sub LBL_36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(36): End Sub
Private Sub LBL_37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(37): End Sub
Private Sub LBL_38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(38): End Sub
Private Sub LBL_39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(39): End Sub
Private Sub LBL_40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(40): End Sub
Private Sub LBL_41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(41): End Sub
Private Sub LBL_42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(42): End Sub
Private Sub LBL_43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(43): End Sub
Private Sub LBL_44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(44): End Sub
Private Sub LBL_45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = tblGuide(45): End Sub
'-------------------------------------------------------------------------------
' �Œ�K�C�h(�j�����x����)
Private Sub LBL_SUN_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_MON_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_TUE_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_WED_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_THU_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_FRI_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_SAT_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = cnsDefaultGuide: End Sub
Private Sub LBL_PREV_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = "�O���ɖ߂�܂�(PageUp)": End Sub
Private Sub LBL_NEXT_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = "�����ɐi�݂܂�(PageDown)": End Sub
Private Sub LBL_YM_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = "�N������I�����܂��B": End Sub
Private Sub LBL_YEAR_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = "�N���I���ł��܂��B": End Sub
Private Sub LBL_MONTH_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): LBL_GUIDE.Caption = "�����I���ł��܂��B": End Sub

'*******************************************************************************
' �u��(�O��)�vClick�C�x���g
'*******************************************************************************
Private Sub LBL_PREV_Click()
    Call ERASE_YEAR_MONTH                       ' �N���R���{�̔�\����
    g_FormDate1 = DateSerial(g_intCurYear, g_intCurMonth - 1, 1)
    ' �J�����_�[�쐬
    Call GP_MakeCalendar
End Sub

'*******************************************************************************
' �u��(����)�vClick�C�x���g
'*******************************************************************************
Private Sub LBL_NEXT_Click()
    Call ERASE_YEAR_MONTH                       ' �N���R���{�̔�\����
    g_FormDate1 = DateSerial(g_intCurYear, g_intCurMonth + 1, 1)
    ' �J�����_�[�쐬
    Call GP_MakeCalendar
End Sub

'*******************************************************************************
' �u���vClick�C�x���g
'*******************************************************************************
Private Sub LBL_MONTH_Click()
    Dim intMonth As Integer
    Dim IX As Long, CUR As Long
    
    Call ERASE_YEAR            ' �N�R���{���\������Ă��������
    ' �N�R���{�̕\��
    g_swBatch = True
    With CBO_MONTH
        .Clear
        For intMonth = 1 To 12
            .AddItem Format(intMonth, "00")
            If intMonth = g_intCurMonth Then CUR = IX
            IX = IX + 1
        Next intMonth
        .ListIndex = CUR
        .Visible = True
        g_VisibleMonth = True
    End With
    g_swBatch = False
End Sub

'*******************************************************************************
' �u�N�vClick�C�x���g
'*******************************************************************************
Private Sub LBL_YEAR_Click()
    Dim intYear As Integer, intYearSTR As Integer, intYearEND As Integer
    Dim IX As Long, CUR As Long
    
    Call ERASE_MONTH            ' ���R���{���\������Ă��������
    ' �N�R���{�̕\��
    g_swBatch = True
    With CBO_YEAR
        .Clear
        intYearSTR = g_intCurYear - 10
        If intYearSTR < g_cnsYearFrom Then intYearSTR = g_cnsYearFrom
        intYearEND = g_intCurYear + 10
        intYear = Year(Date) + g_cnsYearToAdd
        If intYearEND > intYear Then intYearEND = intYear
        For intYear = intYearSTR To intYearEND
            .AddItem CStr(intYear)
            If intYear = g_intCurYear Then CUR = IX
            IX = IX + 1
        Next intYear
        .ListIndex = CUR
        .Visible = True
        g_VisibleYear = True
    End With
    g_swBatch = False
End Sub

'*******************************************************************************
' �t�H�[���\��(�J��Ԃ��\���̏ꍇ��Hide�݂̂̂���Initialize�͋N���Ȃ�)
'*******************************************************************************
Private Sub UserForm_Activate()
    ' Tag������t�����o��
    g_FormDate1 = CDate(Me.Tag)
    ' Tag�͔񐔒l��Ԃɂ��Ă���
    Me.Tag = False
    ' �R���{�͔�\��
    CBO_YEAR.Visible = False
    CBO_MONTH.Visible = False
    g_VisibleYear = False
    g_VisibleMonth = False
    ' �����̔N�����Z�b�g
    If g_FormDate1 = 0 Then g_FormDate1 = Date
    ' �J�����_�[�쐬
    Call GP_MakeCalendar
    LBL_GUIDE.Caption = cnsDefaultGuide             ' �K�C�h�\��
End Sub

'*******************************************************************************
' �t�H�[��������(�J��Ԃ��\���̏ꍇ��Hide�݂̂̂���Initialize�͋N���Ȃ�)
'*******************************************************************************
Private Sub UserForm_Initialize()
    Dim IX As Integer
    Dim strName As String
    
    ' �N�Z�j���ɂ��j�����o���̈ʒu�C��
    If g_cnsStartYobi = 2 Then
        ' ���j�N�Z
        LBL_MON.Left = 2
        LBL_TUE.Left = 18.5
        LBL_WED.Left = 35
        LBL_THU.Left = 51.5
        LBL_FRI.Left = 68
        LBL_SAT.Left = 84.5
        LBL_SUN.Left = 101
        ' �j���R�[�h�̐ݒ�
        g_intSunday = 7
        g_intSaturday = 6
    Else
        ' ���j�N�Z
        LBL_SUN.Left = 2
        LBL_MON.Left = 18.5
        LBL_TUE.Left = 35
        LBL_WED.Left = 51.5
        LBL_THU.Left = 68
        LBL_FRI.Left = 84.5
        LBL_SAT.Left = 101
        ' �j���R�[�h�̐ݒ�
        g_intSunday = 1
        g_intSaturday = 7
    End If
    ' ���t���x����Object�^�z��ϐ��ɃZ�b�g(�������ł͂��̕ϐ��Œl��o�^)
    For IX = 1 To 45
        strName = "LBL_" & Format(IX, "00")
        Set tblDate(IX) = Me.Controls(strName)
    Next IX
    g_swCalendar1Loaded = True              ' Load����X�C�b�`(��Load)
End Sub

'*******************************************************************************
' �t�H�[����L�[�{�[�h����
'*******************************************************************************
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                             ByVal Shift As Integer)

    ' KeyCode(Shift���p)�ɂ�鐧��
    Select Case KeyCode
        Case vbKeyReturn, vbKeyExecute, vbKeySeparator  ' Enter(����)
            Call GP_ClickCalendar(g_FormDate1)
        Case vbKeyCancel, vbKeyEscape                   ' Cancel, Esc(�I��)
            Me.Hide
        Case vbKeyPageDown                              ' PageDown(����)
            Call LBL_NEXT_Click
        Case vbKeyPageUp                                ' KeyPageUp(�O��)
            Call LBL_PREV_Click
        Case vbKeyRight, vbKeyNumpad6, vbKeyAdd         ' ��(����)
            Call GP_MOVE_DAY(1)
        Case vbKeyLeft, vbKeyNumpad4, vbKeySubtract     ' ��(�O��)
            Call GP_MOVE_DAY(-1)
        Case vbKeyUp, vbKeyNumpad8                      ' ��(7����)
            Call GP_MOVE_DAY(-7)
        Case vbKeyDown, vbKeyNumpad2                    ' ��(7���O)
            Call GP_MOVE_DAY(7)
        Case vbKeyHome                                  ' Home(����)
            Call GP_MOVE_DAY(g_POS_F - g_CurPos)
        Case vbKeyEnd                                   ' End(����)
            Call GP_MOVE_DAY(g_POS_T - g_CurPos)
        Case vbKeyTab                                   ' Tab(Shift�ɂ��)
            If Shift = 1 Then
                Call GP_MOVE_DAY(-1)            ' �O��
            Else
                Call GP_MOVE_DAY(1)             ' ����
            End If
        Case vbKeyF11                                   ' F11(�O�N)
            g_FormDate1 = DateSerial(g_intCurYear - 1, g_intCurMonth, 1)
            Call ERASE_YEAR_MONTH                       ' �N���R���{�̔�\����
            ' �J�����_�[�쐬
            Call GP_MakeCalendar
        Case vbKeyF12                                   ' F12(���N)
            g_FormDate1 = DateSerial(g_intCurYear + 1, g_intCurMonth, 1)
            Call ERASE_YEAR_MONTH                       ' �N���R���{�̔�\����
            ' �J�����_�[�쐬
            Call GP_MakeCalendar
    End Select
End Sub

'*******************************************************************************
' �t�H�[����}�E�X�ړ�
'*******************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    Me.LBL_GUIDE.Caption = cnsDefaultGuide
End Sub

'*******************************************************************************
' �t�H�[���I��
'*******************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call ERASE_YEAR_MONTH                       ' �N���R���{�̔�\����
    ' ����[�~]�{�^���������ꂽ���AUnload����Ȃ��悤�ɂ���
    Cancel = True
    Me.Hide
End Sub

Private Sub UserForm_Terminate()
    g_swCalendar1Loaded = False             ' Load����X�C�b�`(��Unload)
End Sub

'*******************************************************************************
' �����ʃT�u����
'*******************************************************************************
' �J�����_�[�\������
'*******************************************************************************
Private Sub GP_MakeCalendar()
    Dim dteDate As Date, dteDate2 As Date, dteDateF As Date, dteDateT As Date
    Dim intYOBI As Integer, intYear As Integer
    Dim IX As Long, IX2 As Long, IXH As Long, IXH_MAX As Long
    Dim tblTODAY As Variant
    
    ' �w��N�������p�\���`�F�b�N
    intYear = Year(g_FormDate1)                      ' �w��N
    If ((intYear < g_cnsYearFrom) Or _
        (intYear > (Year(Date) + g_cnsYearToAdd))) Then
        MsgBox "�j���v�Z�͈͂𒴂��Ă��܂��B", vbExclamation, Me.Caption
        g_FormDate1 = tblDate2(g_CurPos)
    End If
    g_intCurYear = Year(g_FormDate1)                 ' �w��N
    g_intCurMonth = Month(g_FormDate1)               ' �w�茎
    dteDateF = DateSerial(g_intCurYear, g_intCurMonth, 1)       ' ������
    dteDateT = DateSerial(g_intCurYear, g_intCurMonth + 1, 0)   ' ������
    LBL_YM.Caption = g_intCurYear & "�N" & Format(g_intCurMonth, "00") & "��"
    ' �O��3�����̏j���e�[�u���擾(���ʊ֐����)
    IXH_MAX = modGetSyukujitsu2.FP_GetHoliday3(g_intCurYear, g_intCurMonth)
    ' �e�[�u���v�f��1���ǉ����Ă���(���[�v���̗v�f�����f��s�v�ɂ���)
    IX = IXH_MAX + 1
    ReDim Preserve g_tblSyuku(IX)
    g_tblSyuku(IX).dteDate = DateSerial(g_intCurYear + 2, 1, 1)
    ' �w����t�����U�A�O�T�̍ŏI��(�y�j��)�ɖ߂�
    dteDate = DateSerial(g_intCurYear, g_intCurMonth, 1)        ' ������
    If g_cnsStartYobi = 2 Then
        intYOBI = Weekday(dteDate, vbMonday)                    ' �j���̎擾
    Else
        intYOBI = Weekday(dteDate, vbSunday)                    ' �j���̎擾
    End If
    dteDate = dteDate - intYOBI
    intYOBI = 0
    ' �擪�̏j���e�[�u���ʒu����(�}�b�`���O���p�̂���)
    IXH = 0
    dteDate2 = dteDate + 1      ' �J�����_�[���̏���
    Do While IXH <= IXH_MAX
        If g_tblSyuku(IXH).dteDate >= dteDate2 Then Exit Do
        IXH = IXH + 1
    Loop
    '---------------------------------------------------------------------------
    ' �t�H�[����̓��t�Z�b�g(7�j�~6�T=42���Œ�)
    For IX = 1 To 42
        ' ���ʒu�̓��t�A�j�����Z�o
        intYOBI = intYOBI + 1
        If intYOBI > 7 Then intYOBI = 1
        dteDate = dteDate + 1
        ' ���t�͕ʃe�[�u���ɃZ�b�g
        tblDate2(IX) = dteDate
        tblYobi(IX) = intYOBI
        tblGuide(IX) = Format(dteDate, cnsDateFormat) & _
            "(" & Format(dteDate, "aaa") & ")"
        If dteDate = dteDateF Then
            ' ��������
            g_POS_F = IX
        ElseIf dteDate = dteDateT Then
            ' ��������
            g_POS_T = IX
        End If
        ' ���x���R���g���[����z�񉻂����ϐ�
        With tblDate(IX)
            ' ���x���ɓ��t���Z�b�g
            .Caption = Day(dteDate)
            ' ���x�A�j���ɂ�胉�x���̏������Z�b�g
            .Font.Bold = False
            .ForeColor = cnsFC_Normal
            If dteDate = g_FormDate1 Then
                ' �����I����t
                .BackColor = cnsBC_Select
                g_CurPos = IX
            ElseIf Month(dteDate) = g_intCurMonth Then
                ' ����
                Select Case intYOBI
                    Case g_intSunday    ' ���j��
                        .BackColor = cnsBC_Sunday
                    Case g_intSaturday  ' �y�j��
                        .BackColor = cnsBC_Saturday
                    Case Else
                        .BackColor = cnsBC_Month
                End Select
            Else
                ' �����ȊO
                .BackColor = cnsBC_Other
            End If
            ' �j��(�ܐU�֋x��)�̔���
            If g_tblSyuku(IXH).dteDate = dteDate Then
                ' �����F��ԂƂ���
                .ForeColor = cnsFC_Hori
                If Month(dteDate) = g_intCurMonth Then .Font.Bold = True
                tblGuide(IX) = tblGuide(IX) & " " & g_tblSyuku(IXH).strName
                ' �j���e�[�u���̎Q��Index�����Z
                IXH = IXH + 1
            End If
        End With
    Next IX
    '---------------------------------------------------------------------------
    ' ��������������̏���
    dteDate = Date              ' ����
    If ((Year(dteDate) <> g_intCurYear) Or (Month(dteDate) < g_intCurMonth)) Then
        IXH_MAX = modGetSyukujitsu2.FP_GetHoliday3(Year(dteDate), Month(dteDate))
    End If
    IXH = 0
    dteDate = Date - 1          ' ���
    Do While IXH <= IXH_MAX
        If g_tblSyuku(IXH).dteDate >= dteDate Then Exit Do
        IXH = IXH + 1
    Loop
    tblTODAY = Array("[���]", "[����]", "[����]")
    For IX = 43 To 45
        tblDate2(IX) = dteDate
        tblGuide(IX) = tblTODAY(IX2) & Format(dteDate, cnsDateFormat) & _
            "(" & Format(dteDate, "aaa") & ")"
        ' �j��(�ܐU�֋x��)�̔���
        If IXH <= IXH_MAX Then
            ' �j���e�[�u���̓��t�Ƃ̈�v�𔻒�
            If g_tblSyuku(IXH).dteDate = dteDate Then
                tblGuide(IX) = tblGuide(IX) & " " & g_tblSyuku(IXH).strName
                ' �j���e�[�u���̎Q��Index�����Z
                IXH = IXH + 1
            End If
        End If
        dteDate = dteDate + 1
        IX2 = IX2 + 1
    Next IX
    LBL_GUIDE.Caption = tblGuide(g_CurPos)      ' �K�C�h�\��
End Sub

'*******************************************************************************
' �J�����_�[��̈ړ�����
'*******************************************************************************
Private Sub GP_MOVE_DAY(intIDOU As Integer)
    Dim intPOS As Integer
    Dim dteDate As Date
    
    Call ERASE_YEAR_MONTH                       ' �N���R���{�̔�\����
    ' �ړ���̈ʒu,���t���Z�o
    intPOS = g_CurPos + intIDOU             ' �ړ���ʒu
    dteDate = g_FormDate1 + intIDOU         ' �ړ�����t
    If ((intPOS < 1) Or (intPOS > 42)) Then
        ' �O�����͗����Ɉړ�
        g_FormDate1 = dteDate
        Call GP_MakeCalendar
        Exit Sub
    End If
    '---------------------------------------------------------------------------
    ' �ȑO�̈ʒu�̓��t���x���̔w�i�F�����ɖ߂�
    With tblDate(g_CurPos)
        If ((g_CurPos >= g_POS_F) And (g_CurPos <= g_POS_T)) Then
            ' ������
            Select Case tblYobi(g_CurPos)
                Case g_intSunday:   .BackColor = cnsBC_Sunday
                Case g_intSaturday: .BackColor = cnsBC_Saturday
                Case Else: .BackColor = cnsBC_Month
            End Select
        Else
            ' �O�㌎
            .BackColor = cnsBC_Other
        End If
    End With
    '---------------------------------------------------------------------------
    ' ����̈ʒu�̓��t���x���̔w�i�F��I����ԂɕύX
    With tblDate(intPOS)
        .BackColor = cnsBC_Select
    End With
    ' ���ݓ��t(�ޔ�)���X�V
    g_FormDate1 = dteDate
    g_CurPos = intPOS
    LBL_GUIDE.Caption = tblGuide(g_CurPos)      ' �K�C�h�\��
End Sub

'*******************************************************************************
' �J�����_�[�N���b�N����
'*******************************************************************************
Private Sub GP_ClickCalendar(dteDate As Date)
    Call ERASE_YEAR_MONTH                       ' �N���R���{�̔�\����
    Me.Tag = CLng(dteDate)  ' ���݂̑I����t(�V���A���l)
    Me.Hide
End Sub

'*******************************************************************************
' ��N������R���{�̔�\����
'*******************************************************************************
Private Sub ERASE_YEAR_MONTH()
    Call ERASE_YEAR
    Call ERASE_MONTH
End Sub

'*******************************************************************************
' ��N��R���{�̔�\����
'*******************************************************************************
Private Sub ERASE_YEAR()
    If g_VisibleYear Then
        CBO_YEAR.Visible = False
        g_VisibleYear = False
    End If
End Sub

'*******************************************************************************
' �����R���{�̔�\����
'*******************************************************************************
Private Sub ERASE_MONTH()
    If g_VisibleMonth Then
        CBO_MONTH.Visible = False
        g_VisibleMonth = False
    End If
End Sub

'--------------------------------<< End of Source >>----------------------------
