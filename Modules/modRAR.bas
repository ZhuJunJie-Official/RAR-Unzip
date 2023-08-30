Attribute VB_Name = "modRAR"

Option Explicit

Const ERAR_END_ARCHIVE = 10
Const ERAR_NO_MEMORY = 11
Const ERAR_BAD_DATA = 12
Const ERAR_BAD_ARCHIVE = 13
Const ERAR_UNKNOWN_FORMAT = 14
Const ERAR_EOPEN = 15
Const ERAR_ECREATE = 16
Const ERAR_ECLOSE = 17
Const ERAR_EREAD = 18
Const ERAR_EWRITE = 19
Const ERAR_SMALL_BUF = 20
 
Const RAR_OM_LIST = 0
Const RAR_OM_EXTRACT = 1
 
Const RAR_SKIP = 0
Const RAR_TEST = 1
Const RAR_EXTRACT = 2
 
Const RAR_VOL_ASK = 0
Const RAR_VOL_NOTIFY = 1

Enum RarOperations
    OP_EXTRACT = 0
    OP_TEST = 1
    OP_LIST = 2
End Enum
 
Public Type RARHeaderData
    ArcName As String * 260
    FileName As String * 260
    Flags As Long
    PackSize As Long
    UnpSize As Long
    HostOS As Long
    FileCRC As Long
    FileTime As Long
    UnpVer As Long
    Method As Long
    FileAttr As Long
    CmtBuf As String
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type
 
Public Type RAROpenArchiveData
    ArcName As String
    OpenMode As Long
    OpenResult As Long
    CmtBuf As String
    CmtBufSize As Long
    CmtSize As Long
    CmtState As Long
End Type
 
Public Declare Function RAROpenArchive Lib "unrar.dll" (ByRef ArchiveData As RAROpenArchiveData) As Long
Public Declare Function RARCloseArchive Lib "unrar.dll" (ByVal hArcData As Long) As Long
Public Declare Function RARReadHeader Lib "unrar.dll" (ByVal hArcData As Long, ByRef HeaderData As RARHeaderData) As Long
Public Declare Function RARProcessFile Lib "unrar.dll" (ByVal hArcData As Long, ByVal Operation As Long, ByVal DestPath As String, ByVal DestName As String) As Long
Public Declare Sub RARSetChangeVolProc Lib "unrar.dll" (ByVal hArcData As Long, ByVal Mode As Long)
Public Declare Sub RARSetPassword Lib "unrar.dll" (ByVal hArcData As Long, ByVal Password As String)

Dim i As Long
Dim Msg As String
Public ExtFolder As String

Public Sub AddFileName(FileName As String, Optional FileSize As Long, Optional FilePackSize As Long, Optional FileCRC32 As Variant)
    On Error Resume Next
    '
    With fMain
    '
        .lvFiles.ListItems.Add i, FileName, FileName, , 1
        .lvFiles.ListItems.Item(i).SubItems(1) = VBA.Replace(Format((FileSize / 1024), "##.0 KB"), ",", ".")
        .lvFiles.ListItems.Item(i).SubItems(2) = VBA.Replace(Format((FilePackSize / 1024), "##.0 KB"), ",", ".")
        .lvFiles.ListItems.Item(i).SubItems(3) = FileCRC32
            i = i + 1
    End With
    '
    On Error GoTo 0
End Sub


Public Sub RARExecute(Mode As RarOperations, RarFile As String, Optional Password As String)
On Error Resume Next
    Dim lHandle As Long
    Dim iStatus As Integer
    Dim uRAR As RAROpenArchiveData
    Dim uHeader As RARHeaderData
    Dim sStat As String, Ret As Long
    '
    i = 1
    '
    uRAR.ArcName = RarFile
    uRAR.CmtBuf = Space(16384)
    uRAR.CmtBufSize = 16384
    '
    If Mode = OP_LIST Then
        uRAR.OpenMode = RAR_OM_LIST
    Else
        uRAR.OpenMode = RAR_OM_EXTRACT
    End If
    '
    lHandle = RAROpenArchive(uRAR)
    If uRAR.OpenResult <> 0 Then OpenError uRAR.OpenResult, RarFile
    '
    If Password <> "" Then RARSetPassword lHandle, Password
    '
    '
    iStatus = RARReadHeader(lHandle, uHeader)
    fMain.Show
        With fMain
        If Mode = OP_LIST Then
        .lvFiles.ListItems.Clear
        Msg = ""
        End If
        Do Until iStatus <> 0
            sStat = Left(uHeader.FileName, InStr(1, uHeader.FileName, vbNullChar) - 1)
            Select Case Mode
                Case RarOperations.OP_EXTRACT
                    .sbStat.Panels(1).Text = "Extracting " & sStat
                    If Dir$("C:\WINDOWS\Temp\WinRAR VB", vbDirectory) = "" Then: MkDir "C:\WINDOWS\Temp\WinRAR VB"
                    Ret = RARProcessFile(lHandle, RAR_EXTRACT, "C:\WINDOWS\Temp\WinRAR VB\", uHeader.FileName)
                Case RarOperations.OP_TEST
                    .sbStat.Panels(1).Text = "Testing " & sStat
                    If Dir$("C:\WINDOWS\Temp\WinRAR VB", vbDirectory) = "" Then: MkDir "C:\WINDOWS\Temp\WinRAR VB"
                    Ret = RARProcessFile(lHandle, RAR_TEST, "C:\WINDOWS\Temp\WinRAR VB\", uHeader.FileName)
                Case RarOperations.OP_LIST
                    AddFileName sStat, uHeader.UnpSize, uHeader.PackSize, uHeader.FileCRC
                    Ret = RARProcessFile(lHandle, RAR_SKIP, "", "")
                End Select
        '
        If Ret = 0 Then
            .sbStat.Panels(1).Text = "测试完整!"
        Else
            ProcessError Ret
        End If
        '
        iStatus = RARReadHeader(lHandle, uHeader)
        .Refresh
    Loop
    '
    If Mode = OP_LIST Then
        If (uRAR.CmtState = 1) Then ShowComment (uRAR.CmtBuf)
    Else
    End If
    If iStatus = ERAR_BAD_DATA Then MakeError ("文件头信息破环")
        .Caption = "AZ Studio RAR 解压工具 - " & uHeader.ArcName
        RARCloseArchive lHandle
    '
    End With
    '
    
        Msg = Msg & "文件名: " & uHeader.ArcName & vbCrLf & "属性: " & _
        uHeader.FileAttr & vbCrLf & "操作系统: " & uHeader.HostOS & vbCrLf & _
        "注释: " & uHeader.CmtBuf
    '
    On Error GoTo 0
End Sub

Public Sub OpenError(ErrorNum As Long, ArcName As String)
On Error Resume Next
    Select Case ErrorNum
    Case ERAR_NO_MEMORY
        MakeError "内存不足"
    Case ERAR_EOPEN:
        MakeError "无法打开 " & ArcName
    Case ERAR_BAD_ARCHIVE:
        MakeError ArcName & " 不是 RAR 压缩文件"
    Case ERAR_BAD_DATA:
        MakeError ArcName & ": 压缩文件头信息丢失"
    End Select
    On Error GoTo 0
End Sub

Public Sub ProcessError(ErrorNum As Long)
On Error Resume Next
    Select Case ErrorNum
    Case ERAR_UNKNOWN_FORMAT
        MakeError "未知压缩文件格式"
    Case ERAR_BAD_ARCHIVE:
        MakeError "坏的卷标"
    Case ERAR_ECREATE:
        MakeError "文件创建错误"
    Case ERAR_EOPEN:
        MakeError "卷标打开错误"
    Case ERAR_ECLOSE:
        MakeError "文件关闭错误"
    Case ERAR_EREAD:
        MakeError "读错误"
    Case ERAR_EWRITE:
        MakeError "写错误"
    Case ERAR_BAD_DATA:
        MakeError "CRC 校验错误"
    End Select
    On Error GoTo 0
End Sub

Public Sub MakeError(Msg As String)
On Error Resume Next
    MsgBox Msg, vbApplicationModal + vbCritical, "错误"
    End
    On Error GoTo 0
End Sub

Private Sub ShowComment(Comment As String)
On Error Resume Next
    '
    With fComment
        .txtComment.Text = Comment
        .Show vbModal, fMain
    End With
    '
    On Error GoTo 0
End Sub

Public Sub ShowProp()
On Error Resume Next
    '
    MsgBox Msg, vbInformation, "属性"
    '
    On Error GoTo 0
End Sub
