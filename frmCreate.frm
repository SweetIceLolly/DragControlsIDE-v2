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
      _ExtentX        =   847
      _ExtentY        =   847
      Sizable         =   0   'False
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar_NoDrop 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "新建项目"
      MaxButtonEnabled=   0   'False
      MinButtonEnabled=   0   'False
      MaxButtonVisible=   0   'False
      MinButtonVisible=   0   'False
      BindCaption     =   -1  'True
      Picture         =   "frmCreate.frx":1BCC2
   End
   Begin DragControlsIDE.DarkImageButton cmdNewWindowProgram 
      Height          =   765
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1349
      Image           =   "frmCreate.frx":1C914
      HasBorder       =   0   'False
      Caption         =   "       新建窗口程序"
      Alignment       =   0
   End
   Begin DragControlsIDE.DarkImageButton cmdNewConsoleProgram 
      Height          =   765
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1349
      Image           =   "frmCreate.frx":1CA47
      HasBorder       =   0   'False
      Caption         =   "       新建控制台程序"
      Alignment       =   0
   End
   Begin DragControlsIDE.DarkImageButton cmdNewPlainCpp 
      Height          =   765
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1349
      Image           =   "frmCreate.frx":1CBC8
      HasBorder       =   0   'False
      Caption         =   "       新建空白C++程序"
      Alignment       =   0
   End
   Begin DragControlsIDE.DarkImageButton cmdOpenProject 
      Height          =   765
      Left            =   480
      TabIndex        =   3
      Top             =   3600
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1349
      Image           =   "frmCreate.frx":1CF0D
      HasBorder       =   0   'False
      Caption         =   "       打开工程..."
      Alignment       =   0
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
