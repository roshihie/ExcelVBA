Attribute VB_Name = "modGetSyukujitsu2"
'*******************************************************************************
'   �j�����菈��        ���N���w��ɂ��j��(�U�x�␳��)��z��ŕԂ��A
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'*******************************************************************************
Option Explicit
Private Const g_cnsFURI = "(�U�֋x��)"
Private Const g_cnsKYU2 = "�����̋x��"
' �j���e�[�u��(���[�U�[��`)
Public Type typSyuku
    dteDate As Date                 ' ���t
    intFuri As Integer              ' �U�֋x��SW(1=�U�֋x��, 0=�ʏ�)
    strName As String               ' �j������
End Type
' ���L�����ō쐬�����j���e�[�u��
Public g_tblSyuku() As typSyuku     ' �j���e�[�u��(�Ăь��ŗ��p����)

'*******************************************************************************
' ���Y�N���̏j�����̃e�[�u�����쐬����(����1�����p)
'
' �߂�l�F�j���e�[�u���̗v�f��(�}�C�i�X���͏j���Ȃ�)
' �����@�FArg1=�N(Integer)
' �@�@�@�@Arg2=��(Integer)
'*******************************************************************************
Public Function FP_GetHoliday1(intY As Integer, _
                               intM As Integer) As Long
    Dim IX As Long              ' �z���Index

    ' �z��̏�����(�v�f��)
    IX = -1
    ReDim g_tblSyuku(0)         ' ��U�������
    ' �j�����̃e�[�u�����쐬(1���������ʏ���)
    Call GP_GetHolidaySub(intY, intM, IX)
    ' �߂�l�̃Z�b�g
    FP_GetHoliday1 = IX
End Function

'*******************************************************************************
' �O����3�����̏j�����̃e�[�u�����쐬����(����+�O���3�����p)
'
' �߂�l�F�j���e�[�u���̗v�f��
' �����@�FArg1=�N(Integer)
' �@�@�@�@Arg2=��(Integer)
'*******************************************************************************
Public Function FP_GetHoliday3(intYear As Integer, _
                               intMonth As Integer) As Long
    Dim intY As Integer, intM As Integer
    Dim IX As Long, IX2 As Long
    
    ' �z��̏�����(�v�f��)
    IX = -1
    ReDim g_tblSyuku(0)         ' ��U�������
    ' �O���̔N�����Z�o
    If intMonth = 1 Then
        intY = intYear - 1
        intM = 12
    Else
        intY = intYear
        intM = intMonth - 1
    End If
    ' �O�E���E����3�������J��Ԃ�
    For IX2 = 1 To 3
        ' �j�����̃e�[�u�����쐬(1���������ʏ���)
        Call GP_GetHolidaySub(intY, intM, IX)
        ' �����̔N�����Z�o
        If intM = 12 Then
            intY = intY + 1
            intM = 1
        Else
            intM = intM + 1
        End If
    Next IX2
    ' �߂�l���Z�b�g
    FP_GetHoliday3 = IX
End Function

'*******************************************************************************
' ���ȉ��̓T�u����
'*******************************************************************************
' �j�����̃e�[�u�����쐬(1���������ʏ���)
'
' �߂�l�F(�Ȃ�)
' �����@�FArg1=�N(Integer)
' �@�@�@�@Arg2=��(Integer)
' �@�@�@�@Arg3=�e�[�u���ŏI�ʒu(Long)  �����O���ڂ̓o�^�ʒu
'*******************************************************************************
Private Sub GP_GetHolidaySub(intY As Integer, _
                             intM As Integer, _
                             IX As Long)
    Dim strName As String, strName2 As String
    
    ' ���ɂ�镪��
    Select Case intM
        '-----------------------------------------------------------------------
        ' 1��
        Case 1
            ' ���U(1/1)
            Call GP_GetHolidaySub2(DateSerial(intY, intM, 1), IX, "���U")
            ' ���l�̓�
            strName = "���l�̓�"
            If intY < 2000 Then
                ' 1999�N�܂ł�15���Œ�
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 15), IX, strName)
            Else
                ' 2000�N�ȍ~�͑�2���j��
                Call GP_GetHolidaySub3(intY, intM, 2, 2, IX, strName)
            End If
        '-----------------------------------------------------------------------
        ' 2��
        Case 2
            ' �����L�O�̓�(2/11)
            Call GP_GetHolidaySub2(DateSerial(intY, intM, 11), IX, "�����L�O�̓�")
        '-----------------------------------------------------------------------
        ' 3��
        Case 3
            ' �t���̓�(����p����)
            Call GP_GetSyunbun(intY, IX)
        '-----------------------------------------------------------------------
        ' 4��
        Case 4
            ' �݂ǂ�̓�(4/29) �� ���a�̓�(2007�N�`)
            If intY >= 2007 Then
                strName = "���a�̓�"
            Else
                strName = "�݂ǂ�̓�"
            End If
            Call GP_GetHolidaySub2(DateSerial(intY, intM, 29), IX, strName)
        '-----------------------------------------------------------------------
        ' 5��
        Case 5
            strName = "���@�L�O��"
            strName2 = "�q���̓�"
            If intY >= 1985 Then
                IX = IX + 3
                ReDim Preserve g_tblSyuku(IX)
                ' ���@�L�O��(5/3)
                g_tblSyuku(IX - 2).dteDate = DateSerial(intY, intM, 3)
                g_tblSyuku(IX - 2).strName = strName
                ' �����̋x��(5/4) �� �݂ǂ�̓�(2007�N�`)
                g_tblSyuku(IX - 1).dteDate = DateSerial(intY, intM, 4)
                If intY >= 2007 Then
                    g_tblSyuku(IX - 1).strName = "�݂ǂ�̓�"
                Else
                    g_tblSyuku(IX - 1).strName = g_cnsKYU2
                End If
                ' �q���̓�(5/5)
                If intY < 2007 Then
                    IX = IX - 1     ' ��U���Z(����Proc�ŉ��Z����邽��)
                    Call GP_GetHolidaySub2(DateSerial(intY, intM, 5), IX, strName2)
                Else
                    g_tblSyuku(IX).dteDate = DateSerial(intY, intM, 5)
                    g_tblSyuku(IX).strName = strName2
                    ' 2007�N�ȍ~��5/3,5/4�����j�̏ꍇ���A5/6���U��Ԃ���
                    If ((Weekday(g_tblSyuku(IX - 2).dteDate, vbSunday) = vbSunday) Or _
                        (Weekday(g_tblSyuku(IX - 1).dteDate, vbSunday) = vbSunday) Or _
                        (Weekday(g_tblSyuku(IX).dteDate, vbSunday) = vbSunday)) Then
                        IX = IX + 1
                        ReDim Preserve g_tblSyuku(IX)
                        g_tblSyuku(IX).dteDate = DateSerial(intY, intM, 6)
                        g_tblSyuku(IX).intFuri = 1
                        g_tblSyuku(IX).strName = g_cnsFURI
                    End If
                End If
            Else
                ' ���@�L�O��(5/3)
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 3), IX, strName)
                ' �q���̓�(5/5)
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 5), IX, strName2)
            End If
        '-----------------------------------------------------------------------
        ' 6��
        Case 6
            ' �j���Ȃ�
        '-----------------------------------------------------------------------
        ' 7��
        Case 7
            If intY >= 1996 Then
                strName = "�C�̓�"
                If intY >= 2003 Then
                    ' �C�̓�(��3���j��)
                    Call GP_GetHolidaySub3(intY, intM, 3, 2, IX, strName)
                Else
                    ' �C�̓�(7/20)
                    Call GP_GetHolidaySub2(DateSerial(intY, intM, 20), IX, strName)
                End If
            End If
        '-----------------------------------------------------------------------
        ' 8��
        Case 8
            ' �j���Ȃ�
        '-----------------------------------------------------------------------
        ' 9��
        Case 9
            strName = "�h�V�̓�"
            If intY >= 2003 Then
                ' �h�V�̓�(��3���j��)
                Call GP_GetHolidaySub3(intY, intM, 3, 2, IX, strName)
            Else
                ' �h�V�̓�(9/15)
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 15), IX, strName)
            End If
            ' �H���̓�(����p����)
            Call GP_GetSyuubun(intY, IX)
        '-----------------------------------------------------------------------
        ' 10��
        Case 10
            strName = "�̈�̓�"
            If intY >= 2000 Then
                ' �̈�̓�(��2���j��)
                Call GP_GetHolidaySub3(intY, intM, 2, 2, IX, strName)
            Else
                ' �̈�̓�(10/10)
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 10), IX, strName)
            End If
        '-----------------------------------------------------------------------
        ' 11��
        Case 11
            ' �����̓�(11/3)
            Call GP_GetHolidaySub2(DateSerial(intY, intM, 3), IX, "�����̓�")
            ' �ΘJ���ӂ̓�(11/23)
            Call GP_GetHolidaySub2(DateSerial(intY, intM, 23), IX, "�ΘJ���ӂ̓�")
        '-----------------------------------------------------------------------
        ' 12��
        Case 12
            If intY >= 1989 Then
                ' �V�c�a����(12/23)
                Call GP_GetHolidaySub2(DateSerial(intY, intM, 23), IX, "�V�c�a����")
            End If
    End Select
End Sub

'*******************************************************************************
' ���Y�j�������j�Ȃ痂����U�֋x���ɂ��ăe�[�u���Z�b�g(����Sub����)
'
' �߂�l�F(�Ȃ�)
' �����@�FArg1=�j�����t(Date)
' �@�@�@�@Arg2=�e�[�u���ŏI�ʒu(Long)  �����O���ڂ̓o�^�ʒu
' �@�@�@�@Arg3=�j���̖���(String)
'*******************************************************************************
Private Sub GP_GetHolidaySub2(dteHoliday As Date, _
                              IX As Long, _
                              strName As String)
    ' ���Y�j��
    IX = IX + 1
    ReDim Preserve g_tblSyuku(IX)
    g_tblSyuku(IX).dteDate = dteHoliday
    g_tblSyuku(IX).strName = strName
    If Weekday(dteHoliday, vbSunday) = vbSunday Then
        ' ���j�Əd�Ȃ����ꍇ�̗�����U�֋x���Ƃ���
        IX = IX + 1
        ReDim Preserve g_tblSyuku(IX)
        g_tblSyuku(IX).dteDate = dteHoliday + 1
        g_tblSyuku(IX).intFuri = 1          ' �U�֋x��
        g_tblSyuku(IX).strName = g_cnsFURI
    End If
End Sub

'*******************************************************************************
' �N����n�T��m�j�����Z�o���ăe�[�u���Z�b�g(����Sub����)
'
' �߂�l�F(�Ȃ�)
' �����@�FArg1=�N(Integer)
' �@�@�@�@Arg2=��(Integer)
' �@�@�@�@Arg3=�T(Integer)
' �@�@�@�@Arg4=�j���R�[�h(Integer)     ��1=���j, 2=���j�D�D�D7=�y�j(2�̂ݗ��p)
' �@�@�@�@Arg5=�e�[�u���ŏI�ʒu(Long)  �����O���ڂ̓o�^�ʒu
' �@�@�@�@Arg6=�j���̖���(String)
'*******************************************************************************
Private Sub GP_GetHolidaySub3(intY As Integer, _
                              intM As Integer, _
                              intW As Integer, _
                              intG As Integer, _
                              IX As Long, _
                              strName As String)
    Dim dteDate As Date
    Dim intG2 As Integer
    
    IX = IX + 1
    ReDim Preserve g_tblSyuku(IX)
    dteDate = DateSerial(intY, intM, 1)     ' ������
    intG2 = Weekday(dteDate, vbSunday)      ' �������̗j��
    If intG2 > intG Then intW = intW + 1    ' ���T����
    g_tblSyuku(IX).dteDate = dteDate - intG2 + (intW - 1) * 7 + intG
    g_tblSyuku(IX).strName = strName
End Sub

'*******************************************************************************
' �t���̓��̎Z�o(�ȈՌv�Z����)
'
' �߂�l�F(�Ȃ�)
' �����@�FArg1=�N(Integer)
' �@�@�@�@Arg2=�e�[�u���ŏI�ʒu(Long)  �����O���ڂ̓o�^�ʒu
'*******************************************************************************
Private Sub GP_GetSyunbun(intY As Integer, _
                          IX As Long)
    Dim intD As Integer, intY2 As Integer, dteDate As Date
    
    ' �j���@�{�s(1947�N)�ȑO,2151�N�ȍ~(�ȈՌv�Z�s��)�͖���
    IX = IX + 1
    ReDim Preserve g_tblSyuku(IX)
    intY2 = intY - 1980
    Select Case intY
        Case Is <= 1979
            intD = Int(20.8357 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Is <= 2099
            intD = Int(20.8431 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Else
            intD = Int(21.851 + (0.242194 * intY2) - Int(intY2 / 4))
    End Select
    dteDate = DateSerial(intY, 3, intD)
    ' ���Y���t�����j�̏ꍇ�͗����Ƃ���(�U�֋x���Ƃ͂��Ȃ�)
    If Weekday(dteDate, vbSunday) = vbSunday Then
        g_tblSyuku(IX).dteDate = dteDate + 1
    Else
        g_tblSyuku(IX).dteDate = dteDate
    End If
    g_tblSyuku(IX).strName = "�t���̓�"
End Sub

'*******************************************************************************
' �H���̓��̎Z�o(�ȈՌv�Z����)
'
' �߂�l�F(�Ȃ�)
' �����@�FArg1=�N(Integer)
' �@�@�@�@Arg2=�e�[�u���ŏI�ʒu(Long)  �����O���ڂ̓o�^�ʒu
'*******************************************************************************
Private Sub GP_GetSyuubun(intY As Integer, _
                          IX As Long)
    Dim intD As Integer, intY2 As Integer, dteDate As Date
    
    ' �j���@�{�s(1947�N)�ȑO,2151�N�ȍ~(�ȈՌv�Z�s��)�͖���
    intY2 = intY - 1980
    Select Case intY
        Case Is <= 1979
            intD = Int(23.2588 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Is <= 2099
            intD = Int(23.2488 + (0.242194 * intY2) - Int(intY2 / 4))
        Case Else
            intD = Int(24.2488 + (0.242194 * intY2) - Int(intY2 / 4))
    End Select
    dteDate = DateSerial(intY, 9, intD)
    ' ���Y���t�����j�̏ꍇ�͗����Ƃ���(�U�֋x���Ƃ͂��Ȃ�)
    If Weekday(dteDate, vbSunday) = vbSunday Then
        dteDate = dteDate + 1
    End If
    ' 2003�N�ȍ~�͌h�V�̓��̗��X�����H���̓��̏ꍇ��Ԃ̓��͢�����̋x����ɂȂ�
    If ((intY >= 2003) And ((dteDate - g_tblSyuku(IX).dteDate) = 2)) Then
        IX = IX + 2
        ReDim Preserve g_tblSyuku(IX)
        g_tblSyuku(IX - 1).dteDate = dteDate - 1
        g_tblSyuku(IX - 1).strName = g_cnsKYU2
    Else
        IX = IX + 1
        ReDim Preserve g_tblSyuku(IX)
    End If
    g_tblSyuku(IX).dteDate = dteDate
    g_tblSyuku(IX).strName = "�H���̓�"
End Sub

'--------------------------------<< End of Source >>----------------------------

