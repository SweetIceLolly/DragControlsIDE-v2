VERSION 5.00
Begin VB.Form frmCreate 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "新建项目"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   Icon            =   "frmCreate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   6840
      Top             =   5160
      _extentx        =   847
      _extenty        =   847
      sizable         =   0
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar_NoDrop 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7545
      _extentx        =   13309
      _extenty        =   873
      font            =   "frmCreate.frx":1BCC2
      caption         =   "新建项目"
      maxbuttonenabled=   0
      minbuttonenabled=   0
      maxbuttonvisible=   0
      minbuttonvisible=   0
      bindcaption     =   -1
      picture         =   "frmCreate.frx":1BCF6
   End
   Begin DragControlsIDE.DarkImageButton cmdNewWindowProgram 
      Height          =   765
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   6855
      _extentx        =   12091
      _extenty        =   1349
      image           =   "frmCreate.frx":1C948
      alignment       =   0
      hasborder       =   0
      caption         =   "       新建窗口程序"
   End
   Begin DragControlsIDE.DarkImageButton cmdNewConsoleProgram 
      Height          =   765
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   6855
      _extentx        =   12091
      _extenty        =   1349
      image           =   "frmCreate.frx":1CA7B
      alignment       =   0
      hasborder       =   0
      caption         =   "       新建控制台程序"
   End
   Begin DragControlsIDE.DarkImageButton cmdNewPlainCpp 
      Height          =   765
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   6855
      _extentx        =   12091
      _extenty        =   1349
      image           =   "frmCreate.frx":1CBFC
      alignment       =   0
      hasborder       =   0
      caption         =   "       新建空白C++程序"
   End
   Begin DragControlsIDE.DarkImageButton cmdOpenProject 
      Height          =   765
      Left            =   480
      TabIndex        =   3
      Top             =   3600
      Width           =   6855
      _extentx        =   12091
      _extenty        =   1349
      image           =   "frmCreate.frx":1CF41
      alignment       =   0
      hasborder       =   0
      caption         =   "       打开工程..."
   End
   Begin VB.Label labTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "创建"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFB00A&
      Height          =   330
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   600
      Width           =   480
   End
   Begin VB.Label labTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFB00A&
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFB00A&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   90
   End
   Begin VB.Label labTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDB1A&
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFB00A&
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   4680
      Width           =   90
   End
   Begin VB.Label labTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "最近"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDB1A&
      Height          =   330
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   4680
      Width           =   480
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'描述:      新建项目窗口，用户在这里选择新建项目的类型
'作者:      冰棍
'文件:      frmCreateOptions.frm
'====================================================

Option Explicit

Private Sub cmdNewWindowProgram_Click()
    On Error Resume Next
    frmCreateOptions.NewProjectType = 1                 '设置工程类型
    frmCreateOptions.Show                               '显示新建选项
    frmCreateOptions.edProjectName.SetFocus
    Unload Me
End Sub

Private Sub cmdNewConsoleProgram_Click()
    On Error Resume Next
    frmCreateOptions.NewProjectType = 2                 '设置工程类型
    frmCreateOptions.Show                               '显示新建选项
    frmCreateOptions.edProjectName.SetFocus
    Unload Me
End Sub

Private Sub cmdNewPlainCpp_Click()
    On Error Resume Next
    frmCreateOptions.NewProjectType = 3                 '设置工程类型
    frmCreateOptions.Show                               '显示新建选项
    frmCreateOptions.edProjectName.SetFocus
    Unload Me
End Sub

Private Sub cmdOpenProject_Click()
    Call frmMain.HideStartupPage
    
    frmMain.Enabled = True
    frmMain.DarkWindowBorderSizer.Bind = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then                      '按下Esc键关闭窗体
        KeyAscii = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    '加载语言字符串
    Me.Caption = Lang_Create_Caption
    Me.labTip(1).Caption = Lang_Create_CreateLabel
    Me.labTip(3).Caption = Lang_Create_RecentLabel
    Me.cmdNewConsoleProgram.Caption = Lang_Create_NewConsoleProgram
    Me.cmdNewPlainCpp.Caption = Lang_Create_NewEmptyCpp
    Me.cmdNewWindowProgram.Caption = Lang_Create_NewWindowProgram
    Me.cmdOpenProject.Caption = Lang_Create_OpenProject
    '---------------------------------------------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmCreateOptions.Visible = False Then            '如果取消新建，则重新激活主窗体
        Unload frmCreateOptions
        frmMain.Enabled = True
        frmMain.DarkWindowBorderSizer.Bind = True
    End If
End Sub
