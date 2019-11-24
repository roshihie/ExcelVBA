Attribute VB_Name = "Module1"
Option Explicit
'***********************************************************************
'       �t�@�C���t�H�[�}�b�g�ϊ��i�b�k�쐬
'***********************************************************************
'       �����T�v�F�Œ蒷�t�@�C��(FB) ���� �ϒ��t�@�C��(VB)�ɕϊ�����
'                 JCL �ƕϊ������ϒ��t�@�C�� ���� �Œ蒷�t�@�C����
'                 �ĕϊ�����JCL ���쐬���܂��B
'
'***********************************************************************
Public Type StCellPos
  iRow    As Integer
  iCol    As Integer
End Type

Dim g_bERR  As Boolean

Const cn������  As String = "�� �� �� �� ��"

Public Sub �t�H�[�}�b�g�ϊ�JCL�쐬()

  Const cnSheet_Input As String = "���̓V�[�g"

  Dim sFileName  As String
  Dim sVolSer    As String
  Dim iPrim      As Integer
  Dim iSecd      As Integer

  Call prInit(sFileName, sVolSer, iPrim, iSecd)
  
  If g_bERR Then
    Exit Sub
  End If
  
  Call prFB_VB_CnvJCL_Out(sFileName, sVolSer, iPrim, iSecd)
  Call prVB_FB_CnvJCL_Out(sFileName)

  Worksheets(cnSheet_Input).Activate
  Application.StatusBar = False

End Sub
'***********************************************************************
'       ��������
'***********************************************************************
'       �����T�v�F���͂��ꂽ�t�@�C�����́CVOL=SER�CSPACE�ʂ̂P���ʁC
'                 �Q���ʂ��擾����B
'***********************************************************************
Private Sub prInit(argsFileName As String, _
                   argsVolSer As String, _
                   argiPrim As Integer, _
                   argiSecd As Integer)
  
  Const cnFileNameRow As Integer = 12
  Const cnFileNameCol As Integer = 10
  Const cnVolSerRow   As Integer = 14
  Const cnVolSerCol   As Integer = 10
  Const cnPrimRow     As Integer = 16
  Const cnPrimCol     As Integer = 10
  Const cnSecdRow     As Integer = 16
  Const cnSecdCol     As Integer = 18
                    
  Application.ScreenUpdating = False
  Application.StatusBar = False
  Application.StatusBar = cn������
                      
  If Cells(cnFileNameRow, cnFileNameCol).Value = "" Then
    MsgBox prompt:="�Œ蒷�t�@�C�����̂���͂��Ă�������", Buttons:=vbOKOnly + vbCritical
    g_bERR = True
    Exit Sub
  Else
    argsFileName = Cells(cnFileNameRow, cnFileNameCol).Value
  End If
                      
  If Cells(cnVolSerRow, cnVolSerCol).Value = "" Then
    MsgBox prompt:="VOL=SER ����͂��Ă�������", Buttons:=vbOKOnly + vbCritical
    g_bERR = True
  Else
    argsVolSer = Cells(cnVolSerRow, cnVolSerCol).Value
  End If
                      
  If Cells(cnPrimRow, cnPrimCol).Value = "" Then
    MsgBox prompt:="SPACE�̂P���ʂ���͂��Ă�������", Buttons:=vbOKOnly + vbCritical
    g_bERR = True
  Else
    argiPrim = Cells(cnPrimRow, cnPrimCol).Value
  End If
                      
  If Cells(cnSecdRow, cnSecdCol).Value = "" Then
    MsgBox prompt:="SPACE�̂Q���ʂ���͂��Ă�������", Buttons:=vbOKOnly + vbCritical
    g_bERR = True
  Else
    argiSecd = Cells(cnSecdRow, cnSecdCol).Value
  End If

End Sub
'***********************************************************************
'       �Œ蒷�t�@�C��(FB)���ϒ��t�@�C��(VB) �ϊ�JCL�쐬
'***********************************************************************
'       �����T�v�FSheet=FB_VB_CNVJCL ��Ǎ��݁A�Œ蒷�t�@�C��(FB)����
'                 �ϒ��t�@�C��(VB)�ɕϊ�����JCL���쐬����
'***********************************************************************
Private Sub prFB_VB_CnvJCL_Out(argsFileName As String, _
                               argsVolSer As String, _
                               argiPrim As Integer, _
                               argiSecd As Integer)
                               
  Const cnSheet_FB_VB_Cnv   As String = "FB_VB_CNVJCL"
  Const cnText_FB_VB_CnvJCL As String = "\FB_VB_CNVJCL.txt"
  Const cnJCL_DSN           As String = ",DSN="
  Const cnsVB               As String = ".VB"
  Const cnJCL_VolSer        As String = ",VOL=SER="
  Const cnJCL_SPACE         As String = ",SPACE=(TRK,("
  Const cnJCL_RLSE          As String = "),RLSE),"
  
  Const cnRow_DSN_VB1       As Integer = 12
  Const cnRow_VolSer        As Integer = 13
  Const cnRow_DSN_FB        As Integer = 20
  Const cnRow_DSN_VB2       As Integer = 21
  
  Dim oFSys    As New FileSystemObject
  Dim oFStr    As TextStream
  Dim sJCL     As String
  Dim xyStart  As StCellPos
  Dim xyEnd    As StCellPos
  Dim xyCur    As StCellPos
  
  Worksheets(cnSheet_FB_VB_Cnv).Activate
  Set oFStr = oFSys.CreateTextFile( _
                  Filename:=ThisWorkbook.Path & cnText_FB_VB_CnvJCL, _
                  Overwrite:=True)

  xyStart.iRow = 1
  xyStart.iCol = 1
  xyEnd.iRow = Cells(ActiveSheet.Rows.Count, xyStart.iCol).End(xlUp).Row
  xyEnd.iCol = xyStart.iCol
  
  xyCur.iCol = xyStart.iCol
  For xyCur.iRow = xyStart.iRow To xyEnd.iRow
  
    sJCL = Cells(xyCur.iRow, xyCur.iCol).Value
    Select Case xyCur.iRow
      Case cnRow_DSN_VB1
        sJCL = sJCL & cnJCL_DSN & argsFileName & cnsVB & ","
      Case cnRow_VolSer
        sJCL = sJCL & cnJCL_VolSer & argsVolSer & cnJCL_SPACE & _
                      argiPrim & "," & argiSecd & cnJCL_RLSE
      Case cnRow_DSN_FB
        sJCL = sJCL & cnJCL_DSN & argsFileName
      Case cnRow_DSN_VB2
        sJCL = sJCL & cnJCL_DSN & argsFileName & cnsVB
    End Select
      
    oFStr.WriteLine sJCL
  Next
  
End Sub
'***********************************************************************
'       �ϒ��t�@�C��(VB)���Œ蒷�t�@�C��(FB) �ϊ�JCL�쐬
'***********************************************************************
'       �����T�v�FSheet=VB_FB_CNVJCL ��Ǎ��݁A�ϒ��t�@�C��(VB)����
'                 �Œ蒷�t�@�C��(FB)�ɕϊ�����JCL���쐬����
'***********************************************************************
Private Sub prVB_FB_CnvJCL_Out(argsFileName As String)
                               
  Const cnSheet_VB_FB_Cnv   As String = "VB_FB_CNVJCL"
  Const cnText_VB_FB_CnvJCL As String = "\VB_FB_CNVJCL.txt"
  Const cnJCL_DSN           As String = ",DSN="
  Const cnsVB               As String = ".VB"
  
  Const cnRow_DSN_VB        As Integer = 9
  Const cnRow_DSN_FB        As Integer = 10
  
  Dim oFSys    As New FileSystemObject
  Dim oFStr    As TextStream
  Dim sJCL     As String
  Dim xyStart  As StCellPos
  Dim xyEnd    As StCellPos
  Dim xyCur    As StCellPos
  
  Worksheets(cnSheet_VB_FB_Cnv).Activate
  Set oFStr = oFSys.CreateTextFile( _
                  Filename:=ThisWorkbook.Path & cnText_VB_FB_CnvJCL, _
                  Overwrite:=True)

  xyStart.iRow = 1
  xyStart.iCol = 1
  xyEnd.iRow = Cells(ActiveSheet.Rows.Count, xyStart.iCol).End(xlUp).Row
  xyEnd.iCol = xyStart.iCol
  
  xyCur.iCol = xyStart.iCol
  For xyCur.iRow = xyStart.iRow To xyEnd.iRow
  
    sJCL = Cells(xyCur.iRow, xyCur.iCol).Value
    Select Case xyCur.iRow
      Case cnRow_DSN_VB
        sJCL = sJCL & cnJCL_DSN & argsFileName & cnsVB
      Case cnRow_DSN_FB
        sJCL = sJCL & cnJCL_DSN & argsFileName
    End Select
      
    oFStr.WriteLine sJCL
  Next
  
End Sub


