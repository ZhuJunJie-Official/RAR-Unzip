VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fExt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "解压缩到..."
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "fExt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5490
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog CD 
      Left            =   2010
      Top             =   1230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   1440
      Width           =   1185
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "解压缩(&E)"
      Default         =   -1  'True
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   1440
      Width           =   1185
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   5070
      TabIndex        =   3
      Top             =   270
      Width           =   315
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "\"
      Top             =   270
      Width           =   4065
   End
   Begin VB.Image imApp 
      Height          =   720
      Left            =   90
      Picture         =   "fExt.frx":37A2
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "fExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim m_transparencyKey As Long

Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, ByVal nIndex&, ByVal dwNewLong&)
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
    (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long

Private Const SC_CLOSE As Long = &HF060&

Private Const GWL_WNDPROC As Long = (-4&)

Private Const MF_BYCOMMAND As Long = &H0&
Private Const MF_BYPOSITION As Long = &H400&
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_CHECKED As Long = &H8&
Private Const MF_GRAYED As Long = &H1&
Private Const MF_BITMAP = &H4&
Private Const MAX_PATH As Long = 260

Private Type BrowseInfo
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Sub cmdBrowse_Click()
On Error Resume Next
        Dim bi As BrowseInfo
        Dim pidlRoot As Long
        Dim pidl As Long
        Dim buffer As String
        
        With bi
            .hwndOwner = Me.hwnd
            .pidlRoot = pidlRoot
            .lpszTitle = "选择要解压的文件夹"
            .ulFlags = &H1
            .lpfnCallback = 0
            .lParam = 0
            .iImage = 0
        End With
        
        pidl = SHBrowseForFolder(bi)
        
        If pidl <> 0 Then
            buffer = Space$(MAX_PATH)
            SHGetPathFromIDList pidl, buffer
            buffer = Left$(buffer, InStr(buffer, vbNullChar) - 1)
            
            txtPath.Text = buffer
        End If
        
        ' 释放 PIDL
        CoTaskMemFree pidlRoot
        On Error GoTo 0
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    '
    Unload Me
    '
    On Error GoTo 0
End Sub

Private Sub SystemMenu()
On Error Resume Next
    Dim hMenu As Long, hID As Long
    hMenu = GetSystemMenu(Me.hwnd, 0)
    
    InsertMenu hMenu, &HFFFFFFFF, MF_BYCOMMAND + MF_SEPARATOR, 0&, vbNullString
    
    InsertMenu hMenu, &HFFFFFFFF, MF_BYPOSITION, IDM.a, "关于(&A)"
    
    '刷新菜单
    DrawMenuBar hMenu
    
    
    
    procOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    On Error GoTo 0
End Sub

Private Sub DwmAPI()
On Error Resume Next
    Dim m_transparencyKey As Long
    
    m_transparencyKey = RGB(255, 255, 1)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributesByColor Me.hwnd, m_transparencyKey, 0, LWA_COLORKEY

   

    On Error GoTo ern

    Dim mg As MARGINS, en As Long
    mg.m_Left = -1
    mg.m_Button = -1
    mg.m_Right = -1
    mg.m_Top = -1
    'MsgBox "1"
    DwmIsCompositionEnabled en
    If en Then
        'MsgBox "2"
        DwmExtendFrameIntoClientArea Me.hwnd, mg
        'MsgBox "OK!"
        

    End If

    Exit Sub

ern:
On Error GoTo 0
End Sub

Private Sub Form_Paint()
On Error Resume Next
    Dim hBrush As Long, m_Rect As RECT, hBrushOld As Long
    hBrush = CreateSolidBrush(m_transparencyKey)
    hBrushOld = SelectObject(Me.hdc, hBrush)
    GetClientRect Me.hwnd, m_Rect

    FillRect Me.hdc, m_Rect, hBrush
    SelectObject Me.hdc, hBrushOld

    DeleteObject hBrush
    On Error GoTo 0
End Sub

Private Sub cmdExtract_Click()
On Error Resume Next
    '
    If Dir$(txtPath.Text, vbDirectory) = "" Then: MkDir txtPath.Text
    Unload Me
        Call fMain.ExtractRAR(txtPath.Text)
    '
    On Error GoTo 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    '
    If KeyAscii = 27 Then: Unload Me
    '
    On Error GoTo 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    '
    SystemMenu
    DwmAPI
    '
    On Error GoTo 0
End Sub
