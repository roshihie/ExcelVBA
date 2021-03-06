Attribute VB_Name = "Module1"
Option Explicit
'***********************************************************************
'       ファイルフォーマット変換ＪＣＬ作成
'***********************************************************************
'       処理概要：固定長ファイル(FB) から 可変長ファイル(VB)に変換する
'                 JCL と変換した可変長ファイル から 固定長ファイルに
'                 再変換するJCL を作成します。
'
'***********************************************************************
Public Type StCellPos
  iRow    As Integer
  iCol    As Integer
End Type

Dim g_bERR  As Boolean

Const cn処理中  As String = "◆ 処 理 中 ◆"

Public Sub フォーマット変換JCL作成()

  Const cnSheet_Input As String = "入力シート"

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
'       初期処理
'***********************************************************************
'       処理概要：入力されたファイル名称，VOL=SER，SPACE量の１次量，
'                 ２次量を取得する。
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
  Application.StatusBar = cn処理中
                      
  If Cells(cnFileNameRow, cnFileNameCol).Value = "" Then
    MsgBox prompt:="固定長ファイル名称を入力してください", Buttons:=vbOKOnly + vbCritical
    g_bERR = True
    Exit Sub
  Else
    argsFileName = Cells(cnFileNameRow, cnFileNameCol).Value
  End If
                      
  If Cells(cnVolSerRow, cnVolSerCol).Value = "" Then
    MsgBox prompt:="VOL=SER を入力してください", Buttons:=vbOKOnly + vbCritical
    g_bERR = True
  Else
    argsVolSer = Cells(cnVolSerRow, cnVolSerCol).Value
  End If
                      
  If Cells(cnPrimRow, cnPrimCol).Value = "" Then
    MsgBox prompt:="SPACEの１次量を入力してください", Buttons:=vbOKOnly + vbCritical
    g_bERR = True
  Else
    argiPrim = Cells(cnPrimRow, cnPrimCol).Value
  End If
                      
  If Cells(cnSecdRow, cnSecdCol).Value = "" Then
    MsgBox prompt:="SPACEの２次量を入力してください", Buttons:=vbOKOnly + vbCritical
    g_bERR = True
  Else
    argiSecd = Cells(cnSecdRow, cnSecdCol).Value
  End If

End Sub
'***********************************************************************
'       固定長ファイル(FB)→可変長ファイル(VB) 変換JCL作成
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


