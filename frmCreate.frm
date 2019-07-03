VERSION 5.00
Begin VB.Form frmCreate 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "新建项目"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   6840
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      Sizable         =   0   'False
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar 
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
      BindCaption     =   -1  'True
      Picture         =   "frmCreate.frx":0000
   End
   Begin DragControlsIDE.DarkImageButton cmdNewWindowProgram 
      Height          =   765
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1349
      Image           =   "frmCreate.frx":0C52
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
      Image           =   "frmCreate.frx":0D85
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
      Image           =   "frmCreate.frx":0F06
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
      Image           =   "frmCreate.frx":124B
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

Private Sub cmdNewConsoleProgram_Click()
    Call frmMain.HideStartupPage
    
    Unload Me
End Sub

Private Sub cmdNewPlainCpp_Click()
    On Error Resume Next
    
    frmMain.ProjectType = 3                                                                                             '设置工程类型
    Call frmMain.HideStartupPage                                                                                        '隐藏启动界面
    frmMain.DarkMenu.MenuEnabled(29) = False                                                                            '禁用控件箱菜单
    frmMain.DarkMenu.MenuEnabled(30) = False                                                                            '禁用属性菜单
    frmMain.DockingPane.ShowPane 3                                                                                      '显示工程资源管理器
    frmMain.DockingPane.ShowPane 5                                                                                      '显示输出
    frmMain.Caption = "新空白C++程序 - 拖控件大法"                                                                      '更新标题
    
    '构建工程结构
    Dim ParentItem  As Long                                                                                             '树视图的父节点
    frmSolutionExplorer.SolutionTreeView.RemoveItem 0                                                                   '清空树视图
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem("工程")
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem("源文件", ParentItem)
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem("新建空白代码.cpp", ParentItem)
    frmSolutionExplorer.SolutionTreeView.SelectItem ParentItem
    
    frmCodeWindow.Caption = "新建空白代码.cpp"
    frmMain.TabBar.AddForm frmCodeWindow                                                                                '新建一个代码框
    frmMain.picWindowClientArea.Visible = True                                                                          '显示窗口客户区
    frmCodeWindow.SyntaxEdit.SetFocus                                                                                   '让代码框获得焦点
    
    Unload Me
End Sub

Private Sub cmdNewWindowProgram_Click()
    Call frmMain.HideStartupPage
    
    Unload Me
End Sub

Private Sub cmdOpenProject_Click()
    Call frmMain.HideStartupPage
    
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Enabled = True
End Sub
