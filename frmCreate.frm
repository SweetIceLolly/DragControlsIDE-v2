VERSION 5.00
Begin VB.Form frmCreate 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "�½���Ŀ"
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
      Caption         =   "�½���Ŀ"
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
      Caption         =   "       �½����ڳ���"
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
      Caption         =   "       �½�����̨����"
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
      Caption         =   "       �½��հ�C++����"
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
      Caption         =   "       �򿪹���..."
      Alignment       =   0
   End
   Begin VB.Label labTip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
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
      Caption         =   "���"
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
    
    frmMain.ProjectType = 3                                                                                             '���ù�������
    Call frmMain.HideStartupPage                                                                                        '������������
    frmMain.DarkMenu.MenuEnabled(29) = False                                                                            '���ÿؼ���˵�
    frmMain.DarkMenu.MenuEnabled(30) = False                                                                            '�������Բ˵�
    frmMain.DockingPane.ShowPane 3                                                                                      '��ʾ������Դ������
    frmMain.DockingPane.ShowPane 5                                                                                      '��ʾ���
    frmMain.Caption = "�¿հ�C++���� - �Ͽؼ���"                                                                      '���±���
    
    '�������̽ṹ
    Dim ParentItem  As Long                                                                                             '����ͼ�ĸ��ڵ�
    frmSolutionExplorer.SolutionTreeView.RemoveItem 0                                                                   '�������ͼ
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem("����")
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem("Դ�ļ�", ParentItem)
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem("�½��հ״���.cpp", ParentItem)
    frmSolutionExplorer.SolutionTreeView.SelectItem ParentItem
    
    frmCodeWindow.Caption = "�½��հ״���.cpp"
    frmMain.TabBar.AddForm frmCodeWindow                                                                                '�½�һ�������
    frmMain.picWindowClientArea.Visible = True                                                                          '��ʾ���ڿͻ���
    frmCodeWindow.SyntaxEdit.SetFocus                                                                                   '�ô�����ý���
    
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
