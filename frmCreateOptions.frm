VERSION 5.00
Begin VB.Form frmCreateOptions 
   BackColor       =   &H00403D3D&
   BorderStyle     =   0  'None
   Caption         =   "新建项目"
   ClientHeight    =   6870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   Icon            =   "frmCreateOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmCreateOptions.frx":1BCC2
   ScaleHeight     =   6870
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin DragControlsIDE.ImgOptionBox TypeOption 
      Height          =   1380
      Index           =   1
      Left            =   384
      TabIndex        =   0
      Top             =   1392
      Width           =   1308
      _ExtentX        =   2302
      _ExtentY        =   2434
      Image           =   "frmCreateOptions.frx":1C42C
      Content         =   "窗口程序"
   End
   Begin VB.PictureBox picBtnFrame 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00504D4D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   924
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   10110
      TabIndex        =   13
      Top             =   5940
      Width           =   10110
      Begin DragControlsIDE.DarkButton cmdCancel 
         Height          =   492
         Left            =   8568
         TabIndex        =   11
         Top             =   192
         Width           =   1332
         _ExtentX        =   2355
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
         Caption         =   "取消"
         HasBorder       =   0   'False
      End
      Begin DragControlsIDE.DarkButton cmdOK 
         Height          =   492
         Left            =   6936
         TabIndex        =   10
         Top             =   192
         Width           =   1332
         _ExtentX        =   2355
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
         Caption         =   "确定"
         HasBorder       =   0   'False
      End
   End
   Begin VB.PictureBox picProjectFrame 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00302D2D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   5325
      ScaleHeight     =   5445
      ScaleWidth      =   4785
      TabIndex        =   14
      Top             =   495
      Width           =   4788
      Begin DragControlsIDE.DarkCheckBox chkMainArgs 
         Height          =   372
         Left            =   240
         TabIndex        =   7
         Top             =   3048
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "帮我写好main （有参数）"
      End
      Begin DragControlsIDE.DarkCheckBox chkMain 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2610
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "帮我写好main （无参数）"
      End
      Begin DragControlsIDE.DarkEdit edProjectName 
         Height          =   372
         Left            =   288
         TabIndex        =   3
         Top             =   1128
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "新空白C++程序"
      End
      Begin DragControlsIDE.DarkCheckBox chkIncludeStdio 
         Height          =   372
         Left            =   240
         TabIndex        =   9
         Top             =   3912
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "#include <stdio.h>"
      End
      Begin DragControlsIDE.DarkCheckBox chkWinMain 
         Height          =   372
         Left            =   240
         TabIndex        =   8
         Top             =   3480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "帮我写好WinMain"
      End
      Begin DragControlsIDE.DarkButton cmdBrowse 
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   1920
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "浏览..."
         HasBorder       =   0   'False
      End
      Begin DragControlsIDE.DarkEdit edPath 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft YaHei UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "C:\Project"
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设置"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   324
         Index           =   2
         Left            =   288
         TabIndex        =   17
         Top             =   240
         Width           =   480
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目名称:"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Index           =   0
         Left            =   288
         TabIndex        =   16
         Top             =   768
         Width           =   768
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目文件夹:"
         ForeColor       =   &H00FFFFFF&
         Height          =   192
         Index           =   1
         Left            =   288
         TabIndex        =   15
         Top             =   1608
         Width           =   948
      End
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   5355
      Top             =   690
      _ExtentX        =   847
      _ExtentY        =   847
      Sizable         =   0   'False
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar_NoDrop 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
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
      Picture         =   "frmCreateOptions.frx":1C5A8
   End
   Begin DragControlsIDE.ImgOptionBox TypeOption 
      Height          =   1380
      Index           =   2
      Left            =   2016
      TabIndex        =   1
      Top             =   1392
      Width           =   1308
      _ExtentX        =   2302
      _ExtentY        =   2434
      Image           =   "frmCreateOptions.frx":1CD22
      Content         =   "控制台程序"
   End
   Begin DragControlsIDE.ImgOptionBox TypeOption 
      Height          =   1380
      Index           =   3
      Left            =   3648
      TabIndex        =   2
      Top             =   1392
      Width           =   1308
      _ExtentX        =   2302
      _ExtentY        =   2434
      Image           =   "frmCreateOptions.frx":1CED2
      Content         =   "空白C++程序"
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "项目类型"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   390
      TabIndex        =   18
      Top             =   750
      Width           =   960
   End
End
Attribute VB_Name = "frmCreateOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'描述:      新建选项窗口，用户在这里配置项目的名称、路径和一些选项
'作者:      冰棍
'文件:      frmCreateOptions.frm
'====================================================

Option Explicit

'显示“浏览文件夹”对话框
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

'获取特殊目录路径
Private Declare Function SHGetFolderPathA Lib "shell32.dll" (ByVal hwnd As Long, ByVal csidl As Long, ByVal hToken As Long, _
    ByVal dwFlags As Long, pszPath As Any) As Long
    
Public NewProjectType   As Integer                                          '将要新建的项目类型（请见frmMain的ProjectType变量说明）
Dim MyDocPathStr        As String                                           '“我的文档”路径
Dim PathChanged         As Boolean                                          '用户是否更改过路径。如果没更改过，路径会随着项目名称而变化

'描述:      根据不同的工程类型取不同的名字和不同的路径名
Public Sub RefreshName()
    Select Case NewProjectType
        Case 1
            Me.edProjectName.Text = Lang_CreateOptions_WindowProgram
            Me.edPath.Text = MyDocPathStr & "\" & Lang_CreateOptions_WindowProgram
            
        Case 2
            Me.edProjectName.Text = Lang_CreateOptions_ConsoleProgram
            Me.edPath.Text = MyDocPathStr & "\" & Lang_CreateOptions_ConsoleProgram
            
        Case 3
            Me.edProjectName.Text = Lang_CreateOptions_PlainCPP
            Me.edPath.Text = MyDocPathStr & "\" & Lang_CreateOptions_PlainCPP
    End Select
    Me.edProjectName.ToolTipText = Me.edProjectName.Text
    Me.edPath.ToolTipText = Me.edPath.Text
End Sub

Private Sub chkMain_Click()
    Me.chkMainArgs.Value = False
    Me.chkWinMain.Value = False
End Sub

Private Sub chkMainArgs_Click()
    Me.chkMain.Value = False
    Me.chkWinMain.Value = False
End Sub

Private Sub chkWinMain_Click()
    Me.chkMain.Value = False
    Me.chkMainArgs.Value = False
End Sub

Private Sub cmdBrowse_Click()
    '显示“选择文件夹”对话框
    Dim bi      As BROWSEINFO
    Dim pidl    As Long
    Dim NewPath As String * MAX_PATH
    
    With bi
        .hWndOwner = Me.hwnd
        .pidlRoot = 0
        .lpszTitle = Lang_CreateOptions_BrowseCaption
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
    End With
    frmMain.SkinFramework.AutoApplyNewThreads = False                           '禁用皮肤
    frmMain.SkinFramework.AutoApplyNewWindows = False
    pidl = SHBrowseForFolder(bi)
    frmMain.SkinFramework.AutoApplyNewThreads = True                            '启用皮肤
    frmMain.SkinFramework.AutoApplyNewWindows = True
    If pidl <> 0 Then                                                           '如果用户没有取消操作
        If SHGetPathFromIDList(pidl, NewPath) Then
            Me.edPath.Text = Split(NewPath, vbNullChar)(0)
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    '去掉路径里左右两边的空格
    Dim ProjPath        As String                                                                       '工程路径
    Dim ProjName        As String                                                                       '工程名称
    Dim PrevPathChanged As Boolean                                                                      '之前的PathChanged标志
    
    PrevPathChanged = PathChanged
    Me.edPath.Text = Trim(Me.edPath.Text)
    Me.edProjectName.Text = Trim(Me.edProjectName.Text)
    PathChanged = PrevPathChanged
    ProjName = Me.edProjectName.Text
    ProjPath = Me.edPath.Text
    ProjPath = IIf(Right(ProjPath, 1) = "\", ProjPath, ProjPath & "\")                                  '添加"\"到路径末尾
    If Len(Trim(ProjName)) = 0 Then                                                                     '没有输入工程名称
        NoSkinMsgBox Lang_CreateOptions_ProjectNameRequired, vbExclamation, Lang_Msgbox_Error
        Me.edProjectName.SetFocus
        Exit Sub
    End If
    If Not CheckInvalidFileName(ProjName) Then                                                          '检查工程名称是否非法
        NoSkinMsgBox Lang_CreateOptions_InvalidProjectName & ProjName, vbExclamation, Lang_Msgbox_Error
        Me.edProjectName.SetFocus
        Exit Sub
    End If
    
    '检测路径
    If Dir(Me.edPath.Text, vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem) = "" Then     '检测到路径不存在
        MkDir Me.edPath.Text                                                                                '尝试创建文件夹
        If Err.Number <> 0 Then                                                                             '创建文件夹失败
            NoSkinMsgBox Lang_CreateOptions_InvalidProjectPath, vbExclamation, Lang_Msgbox_Error
            Me.cmdBrowse.SetFocus
            Exit Sub
        End If
    Else                                                                                                '检测到路径存在
        If (GetAttr(Me.edPath.Text) And vbDirectory) = 0 Then                                               '目标路径不是文件夹
            NoSkinMsgBox Lang_CreateOptions_InvalidProjectPath, vbExclamation, Lang_Msgbox_Error
            Me.cmdBrowse.SetFocus
            Exit Sub
        End If
    End If
    
    '检测是否有重名文件
    If Dir(ProjPath & ProjName & ".cpp", vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem) <> "" Then      '检测到重名的cpp文件
        If NoSkinMsgBox(Lang_CreateOptions_SameNameReplace_1 & ProjPath & ProjName & ".cpp" & _
           Lang_CreateOptions_SameNameReplace_2, vbQuestion Or vbYesNo, Lang_Msgbox_Confirm) <> vbYes Then
            Exit Sub
        End If
    End If
    If Dir(ProjPath & ProjName & ".myproj", vbDirectory Or vbNormal Or vbReadOnly Or vbHidden Or vbSystem) <> "" Then   '检测到重名的工程文件
        If NoSkinMsgBox(Lang_CreateOptions_SameNameReplace_1 & ProjPath & ProjName & ".myproj" & _
           Lang_CreateOptions_SameNameReplace_2, vbQuestion Or vbYesNo, Lang_Msgbox_Confirm) <> vbYes Then
            Exit Sub
        End If
    End If
    
    '创建项目相关的文件
    Kill ProjPath & ProjName & ".cpp"
    Kill ProjPath & ProjName & ".myproj"
    Err.Clear
    Open ProjPath & ProjName & ".myproj" For Binary As #1                                               '创建工程文件
        If Err.Number <> 0 Then
            MsgBox Lang_CreateOptions_CreationFailure_1 & ProjPath & ProjName & ".myproj" & _
                   Lang_CreateOptions_CreateProjectFailed, vbExclamation, Lang_Msgbox_Error
            Close #1
            Exit Sub
        End If
    Close #1
    Open ProjPath & ProjName & ".cpp" For Binary As #1                                                  '创建cpp文件
        If Err.Number <> 0 Then
            MsgBox Lang_CreateOptions_CreationFailure_1 & ProjPath & ProjName & ".cpp" & _
                   Lang_CreateOptions_CreateProjectFailed, vbExclamation, Lang_Msgbox_Error
            Close #1
            Exit Sub
        End If
    Close #1
    
    '更新窗体状态
    frmMain.ProjectType = NewProjectType                                                                '设置工程类型
    Call frmMain.HideStartupPage                                                                        '隐藏启动界面
    If NewProjectType = 2 Or NewProjectType = 3 Then                                                    '如果不是窗口程序就禁用对应的菜单
        frmMain.DarkMenu.MenuEnabled(29) = False                                                            '禁用控件箱菜单
        frmMain.DarkMenu.MenuEnabled(30) = False                                                            '禁用属性菜单
    End If
    frmMain.DockingPane.ShowPane 3                                                                      '显示工程资源管理器
    frmMain.DockingPane.ShowPane 5                                                                      '显示输出
    frmMain.Caption = ProjName & " - " & Lang_Application_Title                                         '更改标题
    frmMain.SkinFramework.AutoApplyNewThreads = True                                                    '重新激活皮肤，否则代码框的滚动条就不能画到了
    frmMain.SkinFramework.AutoApplyNewWindows = True
    
    '构建工程结构
    Dim ParentItem      As Long                                                                         '树视图的父节点
    Dim GenCode         As String                                                                       '生成的代码
    Dim CodeStartLn     As Long                                                                         '生成代码后光标所在的行
    Dim NewCodeWindow   As frmCodeWindow                                                                '新创建的代码窗口
    
    CurrentProject.ProjectName = ProjName
    frmSolutionExplorer.SolutionTreeView.RemoveItem 0                                                   '清空树视图
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem(CurrentProject.ProjectName)               '添加项目
    ProjectNameTvItem = ParentItem                                                                      '记录工程名称对应的树视图列表项
    ParentItem = frmSolutionExplorer.SolutionTreeView.AddItem(ProjName & ".cpp", ParentItem)
    frmSolutionExplorer.SolutionTreeView.ExpandItems ProjectNameTvItem, 2
    frmSolutionExplorer.SolutionTreeView.SelectItem ParentItem
    ReDim TvItemBinding(0)                                                                              '添加一个树视图列表项和文件序号的绑定
    TvItemBinding(0).FileIndex = 0                                                                      '设置绑定
    TvItemBinding(0).TVITEM = ParentItem
    TvItemBinding(0).IsFolder = False
    
    CodeStartLn = 1
    If Me.chkIncludeStdio.Value = True Then                                                             '#include <stdio.h>
        GenCode = "#include <stdio.h>" & vbCrLf & vbCrLf
        CodeStartLn = CodeStartLn + 2
    End If
    If Me.chkMain.Value = True Then                                                                     'main (无参数)
        'int main() {
        '[Tab]
        '}
        GenCode = GenCode & "int main() {" & vbCrLf & vbTab & vbCrLf & "}" & vbCrLf
        CodeStartLn = CodeStartLn + 1
    ElseIf Me.chkMainArgs.Value = True Then                                                             'main (有参数)
        'int main(int argc, char *argv[]) {
        '[Tab]
        '}
        GenCode = GenCode & "int main(int argc, char *argv[]) {" & vbCrLf & vbTab & vbCrLf & "}" & vbCrLf
         CodeStartLn = CodeStartLn + 1
    ElseIf Me.chkWinMain.Value = True Then                                                              'WinMain
        '#include <windows.h>
        '
        'int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
        '[Tab]
        '}
        GenCode = GenCode & "#include <windows.h>" & vbCrLf & vbCrLf & _
            "int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {" & vbCrLf & vbTab & vbCrLf & "}" & vbCrLf
         CodeStartLn = CodeStartLn + 3
    End If
    'ToDo: replace this with option
    'if use_space_for_tab then                                                                           '检查是否把Tab替换成空格
    GenCode = Replace(GenCode, vbTab, Space(4))
    'end if
    With CurrentProject                                                                                 '设置工程信息
        ReDim .Folders(0)
        ReDim .Files(0)
        With .Files(0)
            .FilePath = ProjPath & ProjName & ".cpp"
            .FolderIndex = 0
            .Changed = True
            .PrevLine = CodeStartLn
            ReDim .Breakpoints(0)
        End With
        .ProjectType = NewProjectType
        .ProjectName = ProjName
        .Changed = True
    End With
    ProjectFolderPath = ProjPath                                                                        '设置项目文件夹路径
    ProjectFilePath = ProjPath & ProjName & ".myproj"                                                   '设置项目工程文件路径
    Set NewCodeWindow = CreateNewCodeWindow(0)                                                          '新建一个代码窗口
    NewCodeWindow.Caption = ProjName & ".cpp"
    frmMain.TabBar.AddForm NewCodeWindow
    CodeWindows.Add NewCodeWindow
    NewCodeWindow.FileIndex = 0                                                                         '设置代码窗口对应的代码文件序号
    NewCodeWindow.SyntaxEdit.Text = GenCode
    frmMain.picWindowClientArea.Visible = True                                                          '显示窗口客户区
    NewCodeWindow.SyntaxEdit.CurrPos.SetPos CodeStartLn, NewCodeWindow.SyntaxEdit.TabSize + 1           '把输入光标移动到适合的位置
    NewCodeWindow.SyntaxEdit.SetFocus                                                                   '让代码框获得焦点
    Unload Me
End Sub

Private Sub edPath_Change()
    PathChanged = True                                                      '用户自行更改过路径
    Me.edPath.ToolTipText = Me.edPath.Text
End Sub

Private Sub edPath_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then                         '响应Ctrl+A
        Me.edPath.SelStart = 0
        Me.edPath.SelLength = Len(Me.edPath.Text)
    End If
End Sub

Private Sub edProjectName_Change()
    Me.edProjectName.ToolTipText = Me.edProjectName.Text
    If Not PathChanged Then                                                 '如果用户没有更改过路径，就自动更改项目文件夹路径
        Me.edPath.Text = MyDocPathStr & "\" & Me.edProjectName.Text
        PathChanged = False
    End If
End Sub

Private Sub edProjectName_GotFocus()
    Me.edProjectName.SelStart = 0
    Me.edProjectName.SelLength = Len(Me.edProjectName.Text)
End Sub

Private Sub edProjectName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then                         '响应Ctrl+A
        Me.edProjectName.SelStart = 0
        Me.edProjectName.SelLength = Len(Me.edProjectName.Text)
    End If
End Sub

Private Sub edProjectName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then                                          '响应回车键
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then                                          '按下Esc键取消新建
        KeyAscii = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_CreateOptions_Caption                                 '设置窗口标题
    Me.labTip(0).Caption = Lang_CreateOptions_ProjectNameLabel
    Me.labTip(1).Caption = Lang_CreateOptions_ProjectFolderLabel
    Me.labTip(3).Caption = Lang_CreateOptions_ProjectType
    Me.TypeOption(1).Content = Lang_CreateOptions_TypeOption_1
    Me.TypeOption(2).Content = Lang_CreateOptions_TypeOption_2
    Me.TypeOption(3).Content = Lang_CreateOptions_TypeOption_3
    Me.chkIncludeStdio.Caption = Lang_CreateOptions_Include
    Me.chkMain.Caption = Lang_CreateOptions_Main_NoArgs
    Me.chkMainArgs.Caption = Lang_CreateOptions_Main_Args
    Me.chkWinMain.Caption = Lang_CreateOptions_WinMain
    Me.cmdBrowse.Caption = Lang_CreateOptions_Browse
    Me.cmdCancel.Caption = Lang_CreateOptions_Cancel
    Me.cmdOK.Caption = Lang_CreateOptions_OK
    '---------------------------------------------------------------------
    
    frmMain.Enabled = False
    frmMain.DarkWindowBorderSizer.Bind = False
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    
    '获取“我的文档”路径作为默认路径
    Dim MyDocPath(MAX_PATH) As Byte
    Dim rtn                 As Long
    
    rtn = SHGetFolderPathA(0, CSIDL_PERSONAL, 0, 0, MyDocPath(0))
    If rtn = S_OK Then
        MyDocPathStr = ByteArrayConv(MyDocPath) & "\MyProjects"
    End If
    
    Call RefreshName
    PathChanged = False                                                         '记录为用户没有更改过路径
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmMain.ProjectType = 0 Then
        frmMain.ShowStartupPage
    Else
        frmMain.HideStartupPage
    End If
    frmMain.Enabled = True
    frmMain.DarkWindowBorderSizer.Bind = True
    frmMain.SkinFramework.AutoApplyNewThreads = True
    frmMain.SkinFramework.AutoApplyNewWindows = True
End Sub

Private Sub TypeOption_Click(Index As Integer)
    On Error Resume Next
    NewProjectType = Index                                                      '更新项目类型
    Call RefreshName
    Me.edProjectName.SetFocus
End Sub
