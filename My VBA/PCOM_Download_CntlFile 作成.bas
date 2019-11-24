Attribute VB_Name = "Module1"
Option Explicit
'***********************************************************************
'       PCOM_Download_CntlFile �쐬
'***********************************************************************
'       �����T�v�FPCOM_Download_Cntl�V�[�g����ɁA
'                 PCOM_Download_CntlFile ���쐬����B
'
'***********************************************************************
Public Type StCellPos
  iRow    As Integer
  iCol    As Integer
End Type

Dim g_bERR  As Boolean

Public Sub PCOM_Download_CntlFile�쐬()

  Const cnSheet_PCOM_Download As String = "PCOM_Download"

  Dim sHostFileName  As String
  Dim sDownloadName  As String
  Dim xyStart        As StCellPos
  Dim xyEnd          As StCellPos

  Call prInit(xyStart, xyEnd)
  
  If g_bERR Then
    Exit Sub
  End If
  
  Call prDownload_CntlFile�쐬(xyStart, xyEnd)

  Worksheets(cnSheet_PCOM_Download).Activate
  Application.StatusBar = False

End Sub
'***********************************************************************
'       ��������
'***********************************************************************
'       �����T�v�FPCOM_Download_Cntl�V�[�g�̍ŏI�s���m�肷��
'
'***********************************************************************
Private Sub prInit(argxyStart As StCellPos, _
                   argxyEnd   As StCellPos )
  
  Const cn������  As String = "�� �� �� �� ��"

  Application.ScreenUpdating = False
  Application.StatusBar = False
  Application.StatusBar = cn������

  argxyStart.iRow = 2
  argxyStart.iCol = 2
  argxyEnd.iRow = Cells(ActiveSheet.Rows.Count, xyStart.iCol).End(xlUp).Row
  argxyEnd.iCol = argxyStart.iCol

End Sub
'***********************************************************************
'       Download Cntl File �쐬
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


