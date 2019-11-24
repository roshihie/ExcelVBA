Attribute VB_Name = "Module4"
Option Explicit

Dim isERR As Integer                                                ' ERROR �t���O

Public Type typDatNam                                               ' �j���i�[�^
    datDate As Date
    strName As String
End Type

Public Type typCellPos                                              ' �Z���|�W�V�����^
    lngRow As Long
    lngCol As Long
End Type

Dim lngSheetLastRow As Long                                         ' �Z���ŏI�s

'*******************************************************************************
'        �b�r�j�[�i���Z�b�g�A�b�v����
'*******************************************************************************
Public Sub �[�i���Z�b�g�A�b�v()

    Dim posStart    As typCellPos
    Dim posEnd      As typCellPos
    Dim intEigyoCnt As Integer
    
    lngSheetLastRow = ActiveSheet.Rows.Count
    
    Call ProcInit(posStart, posEnd)
    If isERR Then
        Exit Sub
    End If
    
    Call ProcDetailInitSet(posStart, posEnd)
                           
    Call ProcHolidaySet(posStart, posEnd, intEigyoCnt)
    
    Call ProcDetailTimeSet(posStart, posEnd, intEigyoCnt)

End Sub

'*******************************************************************************
'        ���@���@���@��
'*******************************************************************************
'        �����T�v�F���t�G���A�̊J�n�s,��A����яI���s,����擾����
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�J�n�|�W�V����   typCellPos
'            �����Q�@�F�I���|�W�V����   typCellPos
'*******************************************************************************
Private Sub ProcInit(posStart As typCellPos, posEnd As typCellPos)
    
    Dim rngFind As Range
    Dim intRtn As Integer
    
    isERR = False
    Set rngFind = Columns(2).Find(What:="21", _
                                  After:=Range("B1"), _
                                  Lookat:=xlWhole, _
                                  Matchbyte:=True)
    If rngFind Is Nothing Then
        intRtn = MsgBox(prompt:="���t�F21�� �� ������܂���", Buttons:=vbOKOnly + vbCritical)
        If intRtn = vbOK Then
            isERR = True
            Exit Sub
        End If
    Else
        posStart.lngRow = rngFind.Row
        posStart.lngCol = rngFind.Column
    End If
        
    Set rngFind = Columns(2).Find(What:="20", _
                                  After:=Range("B1"), _
                                  Lookat:=xlWhole, _
                                  Matchbyte:=True)
    If rngFind Is Nothing Then
        intRtn = MsgBox(prompt:="���t�F20�� �� ������܂���", Buttons:=vbOKOnly + vbCritical)
        If intRtn = vbOK Then
            isERR = True
            Exit Sub
        End If
    Else
        posEnd.lngRow = rngFind.Row
        posEnd.lngCol = rngFind.Column
    End If
        
End Sub

'*******************************************************************************
'        �[�i���@���t�s�@�����ݒ�
'*******************************************************************************
'        �����T�v�F�������C�j���̐ݒ�  �����
'                  �J�n���ԁC�I�����ԁC�x�e���ԁC�R�����g���C��Ɠ��e
'                  �N���A���t�H���g���ɐݒ肷��
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�J�n�|�W�V����   typCellPos
'            �����Q�@�F�I���|�W�V����   typCellPos
'*******************************************************************************
Private Sub ProcDetailInitSet(posStart As typCellPos, posEnd As typCellPos)

    Const conWeekday As String = "�����ΐ��؋��y"
    Const conBlack   As Variant = 1
    
    Dim idxRow As Long
    Dim datRowdat As Date, datEndMonth As Date
    Dim lngEndMonth As Long
                                        ' ���N���� �� �������Z�o
    datEndMonth = DateSerial(Year(Range("L3").Value), Month(Range("L3").Value), 0)
    lngEndMonth = Day(datEndMonth)
    
    For idxRow = posStart.lngRow To posEnd.lngRow
                                        ' ���t�C�j���ݒ�
        If Cells(idxRow, 2).Value = "" Then                         ' ���s�̓��t���u�����N�̂Ƃ�
            If Cells(idxRow - 1, 2).Value = "" Then                     ' �O�s�̓��t���u�����N�̂Ƃ�
                                                                            ' �����Ȃ�
            ElseIf Cells(idxRow - 1, 2).Value < lngEndMonth Then        ' �O�s�̓��t����̂Ƃ�
                Cells(idxRow, 2).Value = Cells(idxRow - 1, 2).Value + 1     '���s�̓��t �� �O�s�̓��t�{1
            End If
        End If
        Select Case Cells(idxRow, 2).Value
            Case Is > 20                                            ' ���s�̓��t��20 �̂Ƃ�
                If Cells(idxRow, 2).Value > lngEndMonth Then               ' ���s�̓��t�����N�����̌����� �̂Ƃ�
                    Range(Cells(idxRow, 2), Cells(idxRow, 3)).ClearContents       ' ���s�̓��t�C�j���N���A
                Else                                                    ' ���s�̓��t�����N�����̌����� �̂Ƃ�
                    datRowdat = DateSerial(Year(Range("F3").Value), _
                                          Month(Range("F3").Value), _
                                          Cells(idxRow, 2).Value)              ' ���s�̗j���Z�b�g
                    Cells(idxRow, 3).Value = Mid(conWeekday, Weekday(datRowdat, vbSunday), 1)
                End If
                                                                    ' ���s�̓��t��1���`20���܂�
            Case Is >= 1
                datRowdat = DateSerial(Year(Range("L3").Value), _
                                      Month(Range("L3").Value), _
                                      Cells(idxRow, 2).Value)                  ' ���s�̗j���Z�b�g
                Cells(idxRow, 3).Value = Mid(conWeekday, Weekday(datRowdat, vbSunday), 1)
            Case Else                                               ' ���s�̓��t���u�����N�̂Ƃ� ����
        End Select
                                        ' ���t�C�j���@�t�H���g��
        With Range(Cells(idxRow, 2), Cells(idxRow, 3))
            .Font.ColorIndex = conBlack
        End With
                                        ' �J�n���ԁC�I������ �N���A���t�H���g��
        With Range(Cells(idxRow, 4), Cells(idxRow, 7))
            .ClearContents
            .Font.ColorIndex = conBlack
        End With
                                        ' �x�e���� �N���A���t�H���g��
        With Range(Cells(idxRow, 10), Cells(idxRow, 11))
            .ClearContents
            .Font.ColorIndex = conBlack
        End With
                                        ' �R�����g���C��Ɠ��e �N���A���t�H���g��
        With Range(Cells(idxRow, 14), Cells(idxRow, 17))
            .ClearContents
            .Font.ColorIndex = conBlack
        End With
    Next

End Sub

'*******************************************************************************
'        �[�i���@���t�s�@�y���C�x���ݒ�
'*******************************************************************************
'*******************************************************************************
'        �����T�v�F�j���̔z����擾���A���t�s�̊Y�����t���j���̂Ƃ�
'                  ��Ɠ��e���ɏj�����̂�\�����t�H���g�Ԃɐݒ肵
'                  �y�j���̓t�H���g�C���j���̓t�H���g�Ԃɐݒ肷��
'                  �܂��A�c�Ɠ����J�E���g���ĕԂ�
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�J�n�|�W�V����   typCellPos
'            �����Q�@�F�I���|�W�V����   typCellPos
'            �����R�@�F�c�Ɠ��J�E���g   Integer
'*******************************************************************************
Private Sub ProcHolidaySet(posStart As typCellPos, posEnd As typCellPos, _
                           intEigyoCnt As Integer)

    Const conHoliday    As Integer = 1
    Const conNotHoliday As Integer = 0
    Const conBlue       As Variant = 5
    Const conRed        As Variant = 3
    
    Dim aryHoliday() As typDatNam
    Dim intHolidayCnt As Integer
    
    Dim intYYYY As Integer, intMM As Integer, intMMCnt As Integer
    Dim isHoliday As Boolean
    Dim stsHoliday As Integer
    Dim stsWeekday As Integer

    Dim datRowDate As Date
    Dim idxRow As Long, idxArray As Long
    
    intYYYY = Year(Range("F3").Value)
    intMM = Month(Range("F3").Value)
    intMMCnt = 2
    intEigyoCnt = 0
    intHolidayCnt = -1
    isHoliday = FuncHolidayGet(intYYYY, intMM, intMMCnt, aryHoliday, intHolidayCnt)
    
    For idxRow = posStart.lngRow To posEnd.lngRow
        If Cells(idxRow, 2).Value <> "" Then
            If Cells(idxRow, 2).Value > 20 Then
                datRowDate = DateSerial(Year(Range("F3").Value), Month(Range("F3").Value), Cells(idxRow, 2).Value)
            Else
                datRowDate = DateSerial(Year(Range("L3").Value), Month(Range("L3").Value), Cells(idxRow, 2).Value)
            End If
            
            stsHoliday = conNotHoliday
            idxArray = 0
            
            Do While (idxArray <= intHolidayCnt)                    ' VBA �̔���� ���S�]��(Complete)�^�ł���A���ׂĂ̏�������
                                                                    ' ����w�肵�� True/False �����肷��i�̒Z���]��(Short-circuit)�^�j
                If aryHoliday(idxArray).datDate > datRowDate Then
                    Exit Do
                End If
            
                If datRowDate = aryHoliday(idxArray).datDate Then
                
                    stsHoliday = conHoliday
                    Range(Cells(idxRow, 2), Cells(idxRow, 3)).Font.ColorIndex = conRed
                    Cells(idxRow, 15).Value = aryHoliday(idxArray).strName
                    Cells(idxRow, 15).Font.ColorIndex = conRed
                    
                End If
                idxArray = idxArray + 1
            
            Loop
            
            stsWeekday = Weekday(datRowDate, vbSunday)
            
            If stsHoliday = conHoliday Then
            ElseIf stsWeekday = vbSaturday Then
                Range(Cells(idxRow, 2), Cells(idxRow, 3)).Font.ColorIndex = conBlue
            ElseIf stsWeekday = vbSunday Then
                Range(Cells(idxRow, 2), Cells(idxRow, 3)).Font.ColorIndex = conRed
            Else
                intEigyoCnt = intEigyoCnt + 1
            End If
        End If
    Next

End Sub

'*******************************************************************************
'        �j �� �� ��  �� �� �� ��
'*******************************************************************************
'        �����T�v�F�J�n�N������w�肳�ꂽ�������̏j�����擾���邽�߂�
'                  ������s���A�j���̔z���Ԃ�
'
'            �߂�l�@�F�j������Ȃ�     Boolean    (True:����CFalse:�Ȃ�)
'            �����P�@�F�J�n�N           Integer
'            �����Q�@�F�J�n��           Integer
'            �����R�@�F�擾����         Integer
'            �����S�@�F�j���̔z��       typDatNam
'            �����T�@�F�z��i�[����     Integer
'*******************************************************************************
Private Function FuncHolidayGet(intYYYY As Integer, _
                                intMM As Integer, _
                                intMMCnt As Integer, _
                                aryHoliday() As typDatNam, _
                                intHolidayCnt As Integer) As Boolean
    Dim idxRow        As Integer
    Dim intCurYYYY As Integer, intCurMM As Integer

    If IsMissing(intMMCnt) Then
        intMMCnt = 1
    End If
    
    intCurYYYY = intYYYY
    intCurMM = intMM
    ReDim aryHoliday(0)
    
    For idxRow = 1 To intMMCnt
    
        Call ProcHolidayGet(intCurYYYY, intCurMM, aryHoliday, intHolidayCnt)
        If intCurMM >= 12 Then
            intCurYYYY = intCurYYYY + 1
            intCurMM = 1
        Else
            intCurMM = intCurMM + 1
        End If
        
    Next
    
    If intHolidayCnt >= 0 Then
        FuncHolidayGet = True
    Else
        FuncHolidayGet = False
    End If
    
End Function

'*******************************************************************************
'        �j �� �� �� �� ��
'*******************************************************************************
'        �����T�v�F�J�n�N������w�肳�ꂽ�������̏j�����擾���邽�߂�
'                  ������s���A�j���̔z���Ԃ�
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�N               Integer
'            �����Q�@�F��               Integer
'            �����R�@�F�j���z��         Date       array
'            �����S�@�F�z��i�[����     Integer
'*******************************************************************************
Private Sub ProcHolidayGet(intCurYYYY As Integer, _
                   intCurMM As Integer, _
                   aryHoliday() As typDatNam, _
                   intHolidayCnt As Integer)
                                        ' ������ɂ�菈������
    Select Case intCurMM
        Case 1                              ' 1��
            intHolidayCnt = intHolidayCnt + 4
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' ���U(1/1) (�U�֋x���Ȃ�)
            aryHoliday(intHolidayCnt - 3).strName = "���U"
            aryHoliday(intHolidayCnt - 3).datDate = DateSerial(intCurYYYY, intCurMM, 1)
                                                ' ��Ћx��(1/2) (�U�֋x���Ȃ�)
            aryHoliday(intHolidayCnt - 2).strName = "��Ћx��(�N���N�n)"
            aryHoliday(intHolidayCnt - 2).datDate = DateSerial(intCurYYYY, intCurMM, 2)
                                                ' ��Ћx��(1/3) (�U�֋x���Ȃ�)
            aryHoliday(intHolidayCnt - 1).strName = "��Ћx��(�N���N�n)"
            aryHoliday(intHolidayCnt - 1).datDate = DateSerial(intCurYYYY, intCurMM, 3)
                                                ' ���l�̓�
            aryHoliday(intHolidayCnt).strName = "���l�̓�"
            If intCurYYYY < 2000 Then               ' 1999�N�܂� 15���Œ�(1/15)
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 15))
            Else                                    ' 2000�N�ȍ~ ��2���j��(�n�b�s�[�}���f�[)
                aryHoliday(intHolidayCnt).datDate = FuncHappyMonday(intCurYYYY, intCurMM, 2, vbMonday)
            End If
                
        Case 2                              ' 2��
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' �����L�O��(2/11)(�U�֋x������)
            aryHoliday(intHolidayCnt).strName = "�����L�O�̓�"
            Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 11))
            
        Case 3                              ' 3��
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' �t���̓� �擾����
            aryHoliday(intHolidayCnt).strName = "�t���̓�"
            Call ProcSyunbunDay(aryHoliday, intHolidayCnt, intCurYYYY)
            
        Case 4                              ' 4��
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' �݂ǂ�̓�(4/29)��(2007�N�ȍ~)���a�̓�(4/29) (�U�֋x������)
            If intCurYYYY < 2007 Then
                aryHoliday(intHolidayCnt).strName = "�݂ǂ�̓�"
            Else
                aryHoliday(intHolidayCnt).strName = "���a�̓�"
            End If
            Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 29))
            
        Case 5                              ' 5��
            If intCurYYYY >= 1985 Then          ' 1985�N�ȏ�
                
                If intCurYYYY < 2007 Then       ' 2007�N����
                
                    intHolidayCnt = intHolidayCnt + 3
                    ReDim Preserve aryHoliday(intHolidayCnt)
                                                    ' ���@�L�O��(5/3) (�U�֋x���Ȃ�)
                    aryHoliday(intHolidayCnt - 2).strName = "���@�L�O��"
                    aryHoliday(intHolidayCnt - 2).datDate = DateSerial(intCurYYYY, intCurMM, 3)
                                                    ' �����̋x��(5/4)��(2007�N�ȍ~)�݂ǂ�̓�(5/4) (�U�֋x���Ȃ�)
                    aryHoliday(intHolidayCnt - 1).strName = "�����̋x��"
                    aryHoliday(intHolidayCnt - 1).datDate = DateSerial(intCurYYYY, intCurMM, 4)
                                                    ' ���ǂ��̓�(5/5) (�U�֋x������)
                    aryHoliday(intHolidayCnt).strName = "���ǂ��̓�"
                    Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 5))
                Else                            ' 2007�N�ȍ~
                                                    '5/3,5/4,5/5 �������ꂩ�����j���̂Ƃ� 5/6 �U�֋x��
                    If (Weekday(DateSerial(intCurYYYY, intCurMM, 3), vbSunday) = vbSunday Or _
                        Weekday(DateSerial(intCurYYYY, intCurMM, 4), vbSunday) = vbSunday) Then
                        
                        intHolidayCnt = intHolidayCnt + 4
                        ReDim Preserve aryHoliday(intHolidayCnt)
                        
                        aryHoliday(intHolidayCnt - 3).strName = "���@�L�O��"
                        aryHoliday(intHolidayCnt - 3).datDate = DateSerial(intCurYYYY, intCurMM, 3)
                        aryHoliday(intHolidayCnt - 2).strName = "�݂ǂ�̓�"
                        aryHoliday(intHolidayCnt - 2).datDate = DateSerial(intCurYYYY, intCurMM, 4)
                        aryHoliday(intHolidayCnt - 1).strName = "���ǂ��̓�"
                        aryHoliday(intHolidayCnt - 1).datDate = DateSerial(intCurYYYY, intCurMM, 5)
                        aryHoliday(intHolidayCnt).strName = "�U�֋x��"
                        aryHoliday(intHolidayCnt).datDate = DateSerial(intCurYYYY, intCurMM, 6)
                    Else
                        intHolidayCnt = intHolidayCnt + 3
                        ReDim Preserve aryHoliday(intHolidayCnt)
                        
                        aryHoliday(intHolidayCnt - 2).strName = "���@�L�O��"
                        aryHoliday(intHolidayCnt - 2).datDate = DateSerial(intCurYYYY, intCurMM, 3)
                        aryHoliday(intHolidayCnt - 1).strName = "�݂ǂ�̓�"
                        aryHoliday(intHolidayCnt - 1).datDate = DateSerial(intCurYYYY, intCurMM, 4)
                        aryHoliday(intHolidayCnt).strName = "���ǂ��̓�"
                        Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 5))
                    End If
                End If
            Else                                ' 1985�N����
                                                    ' ���@�L�O��(5/3) (�U�֋x������)
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "���@�L�O��"
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 3))
                                                    ' ���ǂ��̓�(5/5) (�U�֋x������)
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "���ǂ��̓�"
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 5))
            End If
                
        Case 6                              ' 6��
                                                ' �j���Ȃ�
        Case 7                              ' 7��
            If intCurYYYY >= 1996 Then          ' 1996�N�ȍ~
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                                                    ' �C�̓�
                aryHoliday(intHolidayCnt).strName = "�C�̓�"
                If intCurYYYY >= 2003 Then              ' 2003�N�ȍ~ ��3���j��(�n�b�s�[�}���f�[)
                    aryHoliday(intHolidayCnt).datDate = FuncHappyMonday(intCurYYYY, intCurMM, 3, vbMonday)
                Else                                    ' 2002�N�܂� 20���Œ�(7/20) (�U�֋x������)
                    Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 20))
                End If
            End If
            
        Case 8                              ' 8��
                                                ' �j���Ȃ�
        Case 9                              ' 9��
            
            If intCurYYYY >= 2003 Then          ' 2003�N�ȍ~
                                                    ' �h�V�̓� ��3���j��(�n�b�s�[�}���f�[)
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "�h�V�̓�"
                aryHoliday(intHolidayCnt).datDate = FuncHappyMonday(intCurYYYY, intCurMM, 3, vbMonday)
                                                    ' �H���̓� �擾����
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "�H���̓�"
                Call ProcSyuubunDay(aryHoliday, intHolidayCnt, intCurYYYY)
                                                    ' �����̋x�� ����(�h�V�̓��ƏH���̓��̊Ԃ�1���󂢂��Ƃ� �����̋x��)
                If (aryHoliday(intHolidayCnt).datDate - aryHoliday(intHolidayCnt - 1).datDate) = 2 Then
                
                    intHolidayCnt = intHolidayCnt + 1
                    ReDim Preserve aryHoliday(intHolidayCnt)
                    
                    aryHoliday(intHolidayCnt) = aryHoliday(intHolidayCnt - 1)
                                                        ' �����̋x��
                    aryHoliday(intHolidayCnt - 1).strName = "�����̋x��"
                    aryHoliday(intHolidayCnt - 1).datDate = aryHoliday(intHolidayCnt - 2).datDate + 1
                End If
            Else                                ' 2003�N����
                                                    ' �h�V�̓� 15���Œ�(9/15) (�U�֋x������)
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "�h�V�̓�"
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 15))
                                                    ' �H���̓� �擾����
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                aryHoliday(intHolidayCnt).strName = "�H���̓�"
                Call ProcSyuubunDay(aryHoliday, intHolidayCnt, intCurYYYY)
            End If
            
        Case 10                             ' 10��
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' �̈�̓�
            aryHoliday(intHolidayCnt).strName = "�̈�̓�"
            If intCurYYYY >= 2000 Then              ' 2000�N�ȍ~ ��2���j��(�n�b�s�[�}���f�[)
                aryHoliday(intHolidayCnt).datDate = FuncHappyMonday(intCurYYYY, intCurMM, 2, vbMonday)
            Else                                    ' 1999�N���� 10���Œ�(10/10) (�U�֋x������)
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 10))
            End If
            
        Case 11                             ' 11��
                                                ' �����̓�(11/3) (�U�֋x������)
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
            aryHoliday(intHolidayCnt).strName = "�����̓�"
            Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 3))
                                                ' �ΘJ���ӂ̓�(11/23) (�U�֋x������)
            intHolidayCnt = intHolidayCnt + 1
            ReDim Preserve aryHoliday(intHolidayCnt)
            aryHoliday(intHolidayCnt).strName = "�ΘJ���ӂ̓�"
            Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 23))
            
        Case 12                             ' 12��
            If intCurYYYY >= 1989 Then          ' 1989�N�ȍ~
                intHolidayCnt = intHolidayCnt + 1
                ReDim Preserve aryHoliday(intHolidayCnt)
                                                    ' �V�c�a����(12/23) (�U�֋x������)
                aryHoliday(intHolidayCnt).strName = "�V�c�a����"
                Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, intCurMM, 23))
            End If
            
            intHolidayCnt = intHolidayCnt + 3
            ReDim Preserve aryHoliday(intHolidayCnt)
                                                ' ��Ћx��(12/29) (�U�֋x���Ȃ�)
            aryHoliday(intHolidayCnt - 2).strName = "��Ћx��(�N���N�n)"
            aryHoliday(intHolidayCnt - 2).datDate = DateSerial(intCurYYYY, intCurMM, 29)
                                                ' ��Ћx��(12/30) (�U�֋x���Ȃ�)
            aryHoliday(intHolidayCnt - 1).strName = "��Ћx��(�N���N�n)"
            aryHoliday(intHolidayCnt - 1).datDate = DateSerial(intCurYYYY, intCurMM, 30)
                                                ' ��Ћx��(12/31) (�U�֋x���Ȃ�)
            aryHoliday(intHolidayCnt).strName = "��Ћx��(�N���N�n)"
            aryHoliday(intHolidayCnt).datDate = DateSerial(intCurYYYY, intCurMM, 31)
            
    End Select
        
End Sub

'*******************************************************************************
'        �U�֋x������E�ݒ菈��
'*******************************************************************************
'        �����T�v�F�j�������j���Ȃ�� �U�֋x���Ƃ��� ���j�����x���Ƃ���
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�j���z��         Date       array
'            �����Q�@�F�z��i�[����     Integer
'            �����R�@�F�U�֑O�x��       Date
'*******************************************************************************
Private Sub ProcFuriHoliday(aryHoliday() As typDatNam, _
                            intHolidayCnt As Integer, _
                            datHoliday As Date)

    aryHoliday(intHolidayCnt).datDate = datHoliday                  ' �U�֑O�x�� �ݒ�
    
    If Weekday(datHoliday, vbSunday) = vbSunday Then                ' �U�֑O�x�������j���̂Ƃ�
    
        intHolidayCnt = intHolidayCnt + 1
        ReDim Preserve aryHoliday(intHolidayCnt)
                
        aryHoliday(intHolidayCnt).datDate = datHoliday + 1              ' �U�֌�x�� �ݒ�
        aryHoliday(intHolidayCnt).strName = "�U�֋x��"
    End If

End Sub

'*******************************************************************************
'        �n�b�s�[�}���f�[�擾����
'*******************************************************************************
'        �����T�v�F�j���̎w�肳�ꂽ�T�̎w��j�����Z�o����
'                  (�n�b�s�[�}���f�[�̎w��j���͂��ׂČ��j���ł���)
'
'            �߂�l�@�F�n�b�s�[�}���f�[ Date
'            �����P�@�F�w��N           Integer
'            �����Q�@�F�w�茎           Integer
'            �����R�@�F�w��T           Integer
'            �����S�@�F�w��j��         Integer
'*******************************************************************************
Private Function FuncHappyMonday(intCurYYYY As Integer, _
                         intCurMM As Integer, _
                         intWeekNo As Integer, _
                         intWeekday As Integer) As Date

    Dim dat1stMonth  As Date
    Dim int1stWeekday As Integer
    
    dat1stMonth = DateSerial(intCurYYYY, intCurMM, 1)               ' ������      �Z�o
    int1stWeekday = Weekday(dat1stMonth, vbSunday)                  ' ������ �j�� �Z�o

    If intWeekday < int1stWeekday Then                              ' �w��j�����������̗j�� �̂Ƃ�
        intWeekNo = intWeekNo + 1                                       ' �w��T�{1
    End If
                                        ' �O���̍ŏI�y�j�� �� �w��T�̑O�T�̓y�j�� �� �w��T�̎w��j�� �Z�o
    FuncHappyMonday = dat1stMonth - int1stWeekday + (intWeekNo - 1) * 7 + intWeekday

End Function

'*******************************************************************************
'        �t���̓��@�擾����
'*******************************************************************************
'        �����T�v�F�w�肳�ꂽ�N�̏t���̓����擾���āA�U�֋x������E�ݒ���s��
'
'                  20.8357�F20��20��3���C20.8341�F20��20��1���C21.851�F21��20��25��
'                  0.242194�F�P�N�Ԃ�365���𒴂��鎞�ԁ�5����48��
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�j���z��         Date       array
'            �����Q�@�F�z��i�[����     Integer
'            �����R�@�F�w��N           Integer
'*******************************************************************************
Private Sub ProcSyunbunDay(aryHoliday() As typDatNam, _
                           intHolidayCnt As Integer, _
                           intCurYYYY As Integer)

    Const conMM  As Integer = 3
    Dim intDD As Integer, intYYDiff1980 As Integer
                                        ' �j���@�{�s(1947�N)�ȑO�C2151�N�ȍ~(�ȈՌv�Z�s��)�͖���
    intYYDiff1980 = intCurYYYY - 1980                               ' �w��N��1980�N�̍������Z�o
    
    Select Case intCurYYYY
        Case Is <= 1979                                             ' �w��N��1979�N �̂Ƃ�
            intDD = Int(20.8357 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
        Case Is <= 2099                                             ' �w��N��2099�N �̂Ƃ�
            intDD = Int(20.8341 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
        Case Else                                                   ' �w��N��2100�N �̂Ƃ�
            intDD = Int(21.851 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
    End Select
    
    aryHoliday(intHolidayCnt).strName = "�t���̓�"
    Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, conMM, intDD))

End Sub

'*******************************************************************************
'        �H���̓��@�擾����
'*******************************************************************************
'        �����T�v�F�w�肳�ꂽ�N�̏H���̓����擾���āA�U�֋x������E�ݒ���s��
'
'                  23.2588�F23��6��12���C23.2488�F23��5��58���C24.2488�F24��5��58��
'                  0.242194�F�P�N�Ԃ�365���𒴂��鎞�ԁ�5����48��
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�j���z��         Date       array
'            �����Q�@�F�z��i�[����     Integer
'            �����R�@�F�w��N           Integer
'*******************************************************************************
Private Sub ProcSyuubunDay(aryHoliday() As typDatNam, _
                           intHolidayCnt As Integer, _
                           intCurYYYY As Integer)

    Const conMM  As Integer = 9
    Dim intYYDiff1980 As Integer, intDD As Integer
                                        ' �j���@�{�s(1947�N)�ȑO�C2151�N�ȍ~(�ȈՌv�Z�s��)�͖���
    intYYDiff1980 = intCurYYYY - 1980                                  ' �w��N��1980�N�̍������Z�o
    
    Select Case intCurYYYY
        Case Is <= 1979                                             ' �w��N��1979�N �̂Ƃ�
            intDD = Int(23.2588 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
        Case Is <= 2099                                             ' �w��N��2099�N �̂Ƃ�
            intDD = Int(23.2488 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
        Case Else                                                   ' �w��N��2100�N �̂Ƃ�
            intDD = Int(24.2488 + (0.242194 * intYYDiff1980) - Int(intYYDiff1980 / 4))
           
    End Select
    
    aryHoliday(intHolidayCnt).strName = "�H���̓�"
    Call ProcFuriHoliday(aryHoliday, intHolidayCnt, DateSerial(intCurYYYY, conMM, intDD))

End Sub

'*******************************************************************************
'        �[�i���@�J�n���ԁC�I�����ԁC�x�e���ԁ@�ݒ�
'*******************************************************************************
'        �����T�v�F�[�i���̓��t�s�̊J�n���ԁC�I�����ԁC�x�e���Ԃ̐ݒ���s��
'
'            �߂�l�@�F�Ȃ�
'            �����P�@�F�J�n�|�W�V����   typCellPos
'            �����Q�@�F�I���|�W�V����   typCellPos
'            �����R�@�F�c�Ɠ����J�E���g Integer
'*******************************************************************************
Private Sub ProcDetailTimeSet(posStart As typCellPos, posEnd As typCellPos, _
                              intEigyoCnt As Integer)
                              
    Const conMonthlyTime As Integer = 175
    Const varFromTime    As Variant = "08:50:00"
    Const con1dayToTime  As Date = "17:20:00"
    Const conBlack       As Variant = 1

    Dim dblCalcTime   As Double
    Dim dblTimeSho    As Double
    Dim int1dayHour   As Integer
    Dim dbl1dayTime   As Double
    Dim timToTime     As Date
    Dim is1dayMini    As Boolean
    Dim idxRow        As Long
    
    dblCalcTime = conMonthlyTime / intEigyoCnt
    int1dayHour = Int(dblCalcTime)
    
    dblTimeSho = dblCalcTime - int1dayHour
    
    Select Case dblTimeSho
        Case Is > 0.75
            dbl1dayTime = int1dayHour + 1
        Case Is > 0.5
            dbl1dayTime = int1dayHour + 0.75
        Case Is > 0.25
            dbl1dayTime = int1dayHour + 0.5
        Case Is > 0
            dbl1dayTime = int1dayHour + 0.25
        Case Else
            dbl1dayTime = int1dayHour
    End Select
    
    is1dayMini = True
    If dbl1dayTime = 7.5 Then
        dbl1dayTime = dbl1dayTime + 1                               ' ���x�݁@�@�@�@�F1����
    ElseIf dbl1dayTime > 7.5 Then
        is1dayMini = False
        dbl1dayTime = dbl1dayTime + 1.25                            ' ���x�݁{�[�x�݁F1.25����
    Else
    End If
    timToTime = DateAdd("n", dbl1dayTime * 60, varFromTime)
    
    For idxRow = posStart.lngRow To posEnd.lngRow
    
        If Cells(idxRow, 2).Value <> "" Then
            If Cells(idxRow, 2).Font.ColorIndex = conBlack Then
                Cells(idxRow, 4).Value = Hour(varFromTime)
                Cells(idxRow, 5).Value = Minute(varFromTime)
                Cells(idxRow, 6).Value = Hour(timToTime)
                Cells(idxRow, 7).Value = Minute(timToTime)
                
                Cells(idxRow, 10).Value = "1"
                If is1dayMini Then
                    Cells(idxRow, 11).Value = "00"
                Else
                    Cells(idxRow, 11).Value = "15"
                End If
            End If
        End If
    Next
    
End Sub

