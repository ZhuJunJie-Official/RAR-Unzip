VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fMain 
   Caption         =   "AZ Studio RAR 解压工具"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8340
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
   Icon            =   "fMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8340
   StartUpPosition =   1  '所有者中心
   Begin ComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   1170
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   2064
      ButtonWidth     =   1455
      ButtonHeight    =   1905
      Appearance      =   1
      ImageList       =   "imMenu"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "打开"
            Key             =   ""
            Object.ToolTipText     =   "打开一个压缩文件"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "关闭"
            Key             =   ""
            Object.ToolTipText     =   "关闭当前压缩文件"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "解压缩"
            Key             =   ""
            Object.ToolTipText     =   "解压缩文件 "
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "测试"
            Key             =   ""
            Object.ToolTipText     =   " 测试压缩文件"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "关于"
            Key             =   ""
            Object.ToolTipText     =   " 关于"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar sbStat 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5370
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView lvFiles 
      Height          =   3825
      Left            =   30
      TabIndex        =   1
      Top             =   1290
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   6747
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "imFile"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "文件名"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "大小"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "压缩后大小"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "CRC32"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1950
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   9495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14415
   End
   Begin VB.Image imApp 
      Height          =   480
      Left            =   960
      Picture         =   "fMain.frx":37A2
      Top             =   9000
      Visible         =   0   'False
      Width           =   480
   End
   Begin ComctlLib.ImageList imMenu 
      Left            =   5640
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":446C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":5FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":7B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":9662
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":B1B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":CD06
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imFile 
      Left            =   4290
      Top             =   3660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fMain.frx":E858
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuopen 
         Caption         =   "打开压缩文件(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuclose 
         Caption         =   "关闭压缩文件(&C)..."
         Shortcut        =   ^C
      End
      Begin VB.Menu mnusep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprop 
         Caption         =   "属性(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "命令(&E)"
      Begin VB.Menu mnuextract 
         Caption         =   "解压缩(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnutest 
         Caption         =   "测试压缩文件(&T)"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuabout 
         Caption         =   "关于(&A)"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "fMain"
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

Private Sub Form_Load()
On Error Resume Next
    '
    SystemMenu
    DwmAPI

    InitCommonControlsVB
    MakeFlatToolbar tbMenu
    ReadCommand (Command)
    '
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Label1.Height = Me.Height
    Label1.Width = Me.Width
        ' 如果应用程序是最小化到任务栏上，则不执行Resize
        If Me.WindowState = vbMinimized Then: Exit Sub
            lvFiles.Top = tbMenu.Top + tbMenu.Height
            lvFiles.Left = 30
            lvFiles.Width = Me.ScaleWidth - 60
            lvFiles.Height = Me.ScaleHeight - (sbStat.Height + 30 + tbMenu.Height)
        Dim OnePart As Long
        OnePart = lvFiles.Width / 10
        lvFiles.ColumnHeaders.Item(1).Width = OnePart * 6
        lvFiles.ColumnHeaders.Item(2).Width = OnePart
        lvFiles.ColumnHeaders.Item(3).Width = OnePart
        lvFiles.ColumnHeaders.Item(4).Width = OnePart
    '
    On Error GoTo 0
End Sub

Private Sub lvFiles_Click()
On Error Resume Next
    '
    If lvFiles.ListItems.Count = 0 Then: Exit Sub
    If lvFiles.SelectedItem.Selected = True Then: Exit Sub
    sbStat.Panels(1).Text = "按 F1 获取更多帮助"
    '
    On Error GoTo 0
End Sub

Private Sub lvFiles_ItemClick(ByVal Item As ComctlLib.ListItem)
On Error Resume Next
    '
    If lvFiles.ListItems.Count = 0 Then: Exit Sub
    sbStat.Panels(1).Text = lvFiles.ListItems(Item.Index).Key
    '
    On Error GoTo 0
End Sub

Private Sub mnuabout_Click()
On Error Resume Next
    '
    Call ShowAbout
    '
    On Error GoTo 0
End Sub

Private Sub mnuclose_Click()
On Error Resume Next
    '
    lvFiles.ListItems.Clear
    sbStat.Panels(1).Text = "按 F1 获取更多帮助"
    '
    mnuclose.Enabled = False
    mnuextract.Enabled = False
    mnutest.Enabled = False
    mnuprop.Enabled = False
    '
    tbMenu.Buttons(2).Enabled = mnuclose.Enabled
    tbMenu.Buttons(4).Enabled = mnuextract.Enabled
    tbMenu.Buttons(5).Enabled = mnutest.Enabled
    '
    Me.Caption = "AZ Studio RAR 解压工具"
    CD.FileName = ""
    '
    On Error GoTo 0
End Sub

Private Sub mnuexit_Click()
    On Error Resume Next
        Unload fComment
        Unload fMain
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

Private Sub mnuextract_Click()
On Error Resume Next
    '
    fExt.Tag = Me.Tag
    fExt.Show vbModal, Me
    '
    On Error GoTo 0
End Sub

Private Sub mnuopen_Click()
On Error Resume Next
    '
    Call OpenArchive
    '
    On Error GoTo 0
End Sub

Private Sub mnuprop_Click()
On Error Resume Next
    '
    Call ShowProp
    '
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

Private Sub mnutest_Click()
On Error Resume Next
    '
    Call RARExecute(OP_TEST, Me.Tag, "")
    '
    On Error GoTo 0
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As ComctlLib.Button)
On Error Resume Next
    '
    Select Case Button.Index
        Case 1
            mnuopen_Click
        Case 2
            mnuclose_Click
        Case 4
            mnuextract_Click
        Case 5
            mnutest_Click
        Case 7
            mnuabout_Click
    End Select
    '
    On Error GoTo 0
End Sub

Public Sub ExtractRAR(sPath As String)
On Error Resume Next
    '
    ExtFolder = sPath
    Call RARExecute(OP_EXTRACT, Me.Tag, "")
    '
    On Error GoTo 0
End Sub

Public Sub CloseArc()
On Error Resume Next
    '
    mnuclose_Click
    '
    On Error GoTo 0
End Sub
