Attribute VB_Name = "Module1"
Option Explicit
'***********************************************************************
'       PCOM_Download_CntlFile 作成
'***********************************************************************
'       処理概要：PCOM_Download_Cntlシートを基に、
'                 PCOM_Download_CntlFile を作成する。
'
'***********************************************************************
Public Type StCellPos
  iRow    As Integer
  iCol    As Integer
End Type

Dim g_bERR  As Boolean

Public Sub PCOM_Download_CntlFile作成()

  Const cnSheet_PCOM_Download As String = "PCOM_Download"

  Dim sHostFileName  As String
  Dim sDownloadName  As String
  Dim xyStart        As StCellPos
  Dim xyEnd          As StCellPos

  Call prInit(xyStart, xyEnd)
  
  If g_bERR Then
    Exit Sub
  End If
  
  Call prDownload_CntlFile作成(xyStart, xyEnd)

  Worksheets(cnSheet_PCOM_Download).Activate
  Application.StatusBar = False

End Sub
'***********************************************************************
'       初期処理
'***********************************************************************
'       処理概要：PCOM_Download_Cntlシートの最終行を確定する
'
'***********************************************************************
Private Sub prInit(argxyStart As StCellPos, _
                   argxyEnd   As StCellPos )
  
  Const cn処理中  As String = "◆ 処 理 中 ◆"

  Application.ScreenUpdating = False
  Application.StatusBar = False
  Application.StatusBar = cn処理中

  argxyStart.iRow = 2
  argxyStart.iCol = 2
  argxyEnd.iRow = Cells(ActiveSheet.Rows.Count, xyStart.iCol).End(xlUp).Row
  argxyEnd.iCol = argxyStart.iCol

End Sub
'***********************************************************************
'       Download Cntl File 作成
'***********************************************************************
'       処理概要：Sheet=FB_VB_CNVJCL を読込み、固定長ファイル(FB)から
'                 可変長ファイル(VB)に変換するJCLを作成する
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
'       可変長ファイル(VB)→固定長ファイル(FB) 変換JCL作成
'***********************************************************************
'       処理概要：Sheet=VB_FB_CNVJCL を読込み、可変長ファイル(VB)から
'                 固定長ファイル(FB)に変換するJCLを作成する
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


